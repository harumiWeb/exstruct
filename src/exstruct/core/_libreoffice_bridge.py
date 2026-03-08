"""UNO bridge helper executed inside the LibreOffice Python runtime."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import time
from typing import Protocol, cast

import uno

_HMM_PER_POINT = 2540.0 / 72.0


class _ServiceManager(Protocol):
    """Protocol for UNO service-manager access used by the bridge."""

    def createInstanceWithContext(  # noqa: N802
        self, service_name: str, context: object
    ) -> object:
        """Create a UNO service instance using the provided component context."""

        ...


class _UnoContext(Protocol):
    """Protocol for a UNO component context."""

    ServiceManager: _ServiceManager


class _UnoResolver(Protocol):
    """Protocol for resolving remote UNO component contexts."""

    def resolve(self, connection: str) -> object:
        """Resolve a remote UNO component context from a connection string."""

        ...


class _Desktop(Protocol):
    """Protocol for the UNO desktop service used to load documents."""

    def loadComponentFromURL(  # noqa: N802
        self,
        url: str,
        target: str,
        search_flags: int,
        properties: tuple[object, ...],
    ) -> object:
        """Load a document component from a UNO file URL."""

        ...


class _NamedContainer(Protocol):
    """Protocol for UNO containers that expose named members."""

    def getElementNames(self) -> list[str]:  # noqa: N802
        """Return the member names exposed by the container."""

        ...


class _DrawPage(Protocol):
    """Protocol for a LibreOffice draw page."""

    def getCount(self) -> int:  # noqa: N802
        """Return the number of drawable items on the page."""

        ...

    def getByIndex(self, index: int) -> object:  # noqa: N802
        """Return a drawable item by index."""

        ...


class _ShapeLike(Protocol):
    """Protocol for UNO draw objects that report a shape type."""

    def getShapeType(self) -> str:  # noqa: N802
        """Return the UNO shape type name."""

        ...


class _Sheet(Protocol):
    """Protocol for spreadsheet sheets exposed through UNO."""

    def getCharts(self) -> _NamedContainer:  # noqa: N802
        """Return the chart container for the sheet."""

        ...

    def getDrawPage(self) -> _DrawPage:  # noqa: N802
        """Return the draw page for the sheet."""

        ...


class _Sheets(Protocol):
    """Protocol for the spreadsheet sheet collection."""

    def getElementNames(self) -> list[str]:  # noqa: N802
        """Return the sheet names in workbook order."""

        ...

    def getByName(self, name: str) -> object:  # noqa: N802
        """Return a sheet object by name."""

        ...


class _SpreadsheetDocument(Protocol):
    """Protocol for a loaded spreadsheet document."""

    def getSheets(self) -> _Sheets:  # noqa: N802
        """Return the workbook sheet collection."""

        ...

    def close(self, deliver_ownership: bool) -> None:
        """Request a graceful close for the spreadsheet document."""

        ...

    def dispose(self) -> None:
        """Dispose of the spreadsheet document immediately."""

        ...


def main() -> int:
    """Run the bridge entry point and print a JSON payload for the requested extraction."""

    args = _parse_args()
    if args.probe:
        _run_probe()
        return 0
    ctx = _resolve_context(args.host, args.port)
    desktop = cast(
        _Desktop,
        ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx),
    )
    doc = _load_document(desktop, Path(args.file))
    try:
        if args.kind == "draw-page":
            payload = _extract_draw_page_payload(doc)
        else:
            payload = _extract_chart_payload(doc)
    finally:
        _close_document(doc)
    print(json.dumps(payload, ensure_ascii=False))
    return 0


def _parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the LibreOffice bridge."""

    parser = argparse.ArgumentParser()
    parser.add_argument("--probe", action="store_true")
    parser.add_argument("--host")
    parser.add_argument("--port", type=int)
    parser.add_argument("--file")
    parser.add_argument("--kind", choices=("charts", "draw-page"), default="charts")
    args = parser.parse_args()
    if not args.probe and (args.host is None or args.port is None or args.file is None):
        parser.error("--host, --port, and --file are required unless --probe is set.")
    return args


def _run_probe() -> None:
    """Validate that the bridge can be imported and basic UNO types resolve."""

    from com.sun.star.beans import PropertyValue

    prop = PropertyValue()
    prop.Name = "Hidden"
    prop.Value = True
    _ = prop
    _ = _HMM_PER_POINT


def _resolve_context(host: str, port: int) -> _UnoContext:
    """Resolve the remote LibreOffice UNO component context for the requested host and port."""

    local_ctx = uno.getComponentContext()
    resolver = cast(
        _UnoResolver,
        local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver",
            local_ctx,
        ),
    )
    connection = f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
    deadline = time.monotonic() + 15.0
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            return cast(_UnoContext, resolver.resolve(connection))
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            time.sleep(0.1)
    if last_error is not None:
        raise last_error
    raise RuntimeError("Failed to resolve LibreOffice UNO context.")


def _load_document(desktop: _Desktop, file_path: Path) -> _SpreadsheetDocument:
    """Load the target spreadsheet document through the UNO desktop service."""

    from com.sun.star.beans import PropertyValue

    props: list[object] = []
    for name, value in (("Hidden", True), ("ReadOnly", True)):
        prop = PropertyValue()
        prop.Name = name
        prop.Value = value
        props.append(prop)
    return cast(
        _SpreadsheetDocument,
        desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(file_path.resolve())),
            "_blank",
            0,
            tuple(props),
        ),
    )


def _extract_chart_payload(doc: _SpreadsheetDocument) -> dict[str, object]:
    """Collect chart geometry payloads from workbook draw pages."""

    sheets_payload: dict[str, list[dict[str, object]]] = {}
    sheets = doc.getSheets()
    for sheet_name in sheets.getElementNames():
        sheet = cast(_Sheet, sheets.getByName(sheet_name))
        chart_names = set(sheet.getCharts().getElementNames())
        draw_page = sheet.getDrawPage()
        items: list[dict[str, object]] = []
        for index in range(draw_page.getCount()):
            shape = draw_page.getByIndex(index)
            shape_type = _safe_attr(shape, "getShapeType")
            if shape_type != "com.sun.star.drawing.OLE2Shape":
                continue
            persist_name = _safe_attr(shape, "PersistName")
            if persist_name not in chart_names:
                continue
            position = _safe_attr(shape, "Position")
            size = _safe_attr(shape, "Size")
            items.append(
                {
                    "name": _safe_attr(shape, "Name") or persist_name or "Chart",
                    "persist_name": persist_name,
                    "left": _hmm_to_points(getattr(position, "X", None)),
                    "top": _hmm_to_points(getattr(position, "Y", None)),
                    "width": _hmm_to_points(getattr(size, "Width", None)),
                    "height": _hmm_to_points(getattr(size, "Height", None)),
                }
            )
        sheets_payload[str(sheet_name)] = items
    return {"sheets": sheets_payload}


def _extract_draw_page_payload(doc: _SpreadsheetDocument) -> dict[str, object]:
    """Collect non-chart draw-page payloads from workbook sheets."""

    sheets_payload: dict[str, list[dict[str, object]]] = {}
    sheets = doc.getSheets()
    for sheet_name in sheets.getElementNames():
        sheet = cast(_Sheet, sheets.getByName(sheet_name))
        draw_page = sheet.getDrawPage()
        items: list[dict[str, object]] = []
        for index in range(draw_page.getCount()):
            shape = draw_page.getByIndex(index)
            shape_type = _safe_attr(shape, "getShapeType")
            if shape_type == "com.sun.star.drawing.OLE2Shape":
                continue
            position = _safe_attr(shape, "Position")
            size = _safe_attr(shape, "Size")
            items.append(
                {
                    "name": _safe_attr(shape, "Name") or f"Shape {index + 1}",
                    "shape_type": shape_type,
                    "text": _safe_attr(shape, "String") or "",
                    "left": _hmm_to_points(getattr(position, "X", None)),
                    "top": _hmm_to_points(getattr(position, "Y", None)),
                    "width": _hmm_to_points(getattr(size, "Width", None)),
                    "height": _hmm_to_points(getattr(size, "Height", None)),
                    "rotation": _rotation_to_degrees(_safe_attr(shape, "RotateAngle")),
                    "is_connector": shape_type == "com.sun.star.drawing.ConnectorShape",
                    "start_shape_name": _safe_shape_name(shape, "StartShape"),
                    "end_shape_name": _safe_shape_name(shape, "EndShape"),
                }
            )
        sheets_payload[str(sheet_name)] = items
    return {"sheets": sheets_payload}


def _safe_attr(obj: object, name: str) -> object:
    """Read a UNO attribute or helper method while suppressing lookup failures."""

    if name == "getShapeType":
        try:
            return cast(_ShapeLike, obj).getShapeType()
        except Exception:  # noqa: BLE001
            return None
    try:
        return getattr(obj, name)
    except Exception:  # noqa: BLE001
        return None


def _close_document(doc: _SpreadsheetDocument) -> None:
    """Close a UNO spreadsheet document, falling back to dispose when needed."""

    try:
        doc.close(True)
    except Exception:  # noqa: BLE001
        try:
            doc.dispose()
        except Exception:  # noqa: BLE001
            return


def _safe_shape_name(obj: object, attr_name: str) -> str | None:
    """Return the related shape name referenced by a UNO connector endpoint."""

    try:
        ref = getattr(obj, attr_name)
    except Exception:  # noqa: BLE001
        return None
    if ref is None:
        return None
    name = _safe_attr(ref, "Name")
    return name if isinstance(name, str) and name else None


def _hmm_to_points(value: object) -> int | None:
    """Convert hundredths of millimetres to rounded point units."""

    if not isinstance(value, int | float):
        return None
    return int(round(float(value) / _HMM_PER_POINT))


def _rotation_to_degrees(value: object) -> float | None:
    """Convert LibreOffice rotation units to degrees."""

    if not isinstance(value, int | float):
        return None
    return float(value) / 100.0


if __name__ == "__main__":
    raise SystemExit(main())
