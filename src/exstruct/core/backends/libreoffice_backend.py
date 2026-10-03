"""LibreOffice-backed rich shape and chart extraction helpers."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from contextlib import AbstractContextManager
import logging
from pathlib import Path
from typing import Protocol, TypeGuard

from ...models import Arrow, Chart, Shape, SmartArt
from ..libreoffice import (
    LibreOfficeChartGeometry,
    LibreOfficeDrawPageShape,
    LibreOfficeSession,
    LibreOfficeWorkbookHandle,
)
from ..ooxml_drawing import OoxmlConnectorInfo, OoxmlShapeInfo, read_sheet_drawings
from .base import ChartData, RichBackend, ShapeData
from .ooxml_shapes import (  # noqa: F401 - legacy compatibility re-exports
    _build_shapes_from_ooxml as _build_shapes_from_ooxml,
    _classify_connector_resolution as _classify_connector_resolution,  # noqa: F401 - legacy re-export
    _connector_endpoints as _connector_endpoints,
    _direction_from_shape_boxes as _direction_from_shape_boxes,
    _distance_to_box as _distance_to_box,
    _nearest_shape_id as _nearest_shape_id,
    _resolve_connector as _resolve_connector,
    _resolve_direct_connector_ids as _resolve_direct_connector_ids,
    _resolve_direction as _resolve_direction,
    _rotate_connector_delta as _rotate_connector_delta,
    _shape_box_center as _shape_box_center,
    _ShapeBox as _ShapeBox,
    _to_shape_box as _to_shape_box,
)

logger = logging.getLogger(__name__)


class _LibreOfficePathSessionProtocol(Protocol):
    """Structural contract for legacy path-only rich-extraction sessions."""

    def extract_chart_geometries(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeChartGeometry]]:
        """Extract chart geometries for a workbook path."""

    def extract_draw_page_shapes(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeDrawPageShape]]:
        """Extract draw-page shapes for a workbook path."""


class _LibreOfficeWorkbookLifecycleSessionProtocol(Protocol):
    """Structural contract for sessions that support typed workbook handles."""

    def load_workbook(self, file_path: Path) -> LibreOfficeWorkbookHandle:
        """Open a workbook lifecycle handle."""

    def close_workbook(self, workbook: LibreOfficeWorkbookHandle) -> None:
        """Close a workbook lifecycle handle."""

    def extract_chart_geometries(
        self, workbook: Path | LibreOfficeWorkbookHandle
    ) -> dict[str, list[LibreOfficeChartGeometry]]:
        """Extract chart geometries for a path or workbook handle."""

    def extract_draw_page_shapes(
        self, workbook: Path | LibreOfficeWorkbookHandle
    ) -> dict[str, list[LibreOfficeDrawPageShape]]:
        """Extract draw-page shapes for a path or workbook handle."""


_LibreOfficeRichSession = (
    _LibreOfficePathSessionProtocol | _LibreOfficeWorkbookLifecycleSessionProtocol
)


def _supports_workbook_lifecycle(
    session: _LibreOfficeRichSession,
) -> TypeGuard[_LibreOfficeWorkbookLifecycleSessionProtocol]:
    """Return whether a session exposes the typed workbook lifecycle hooks."""

    return hasattr(session, "load_workbook") and hasattr(session, "close_workbook")


class LibreOfficeRichBackend(RichBackend):
    """Best-effort rich extraction backend gated by LibreOffice runtime availability."""

    def __init__(
        self,
        file_path: Path,
        *,
        session_factory: Callable[
            [], AbstractContextManager[_LibreOfficeRichSession]
        ] = (LibreOfficeSession.from_env),
    ) -> None:
        """Store the workbook path and session factory used for lazy LibreOffice extraction."""

        self.file_path = file_path
        self._session_factory = session_factory
        self._chart_geometries: dict[str, list[LibreOfficeChartGeometry]] | None = None
        self._draw_page_shapes: dict[str, list[LibreOfficeDrawPageShape]] | None = None

    def extract_shapes(self, *, mode: str) -> dict[str, list[Shape | Arrow | SmartArt]]:
        """Extract LibreOffice-mode shapes and connectors for each worksheet.

        Args:
            mode: Requested extraction mode. Only ``"libreoffice"`` is supported.

        Returns:
            Mapping of sheet names to emitted shape, arrow, and SmartArt models.

        Raises:
            ValueError: If a non-LibreOffice mode is requested.
        """

        if mode != "libreoffice":
            raise ValueError("LibreOfficeRichBackend only supports libreoffice mode.")
        drawings = read_sheet_drawings(self.file_path)
        draw_page_shapes = self._read_draw_page_shapes()
        shape_data: ShapeData = {}
        sheet_names = list(dict.fromkeys([*drawings.keys(), *draw_page_shapes.keys()]))
        for sheet_name in sheet_names:
            drawing = drawings.get(sheet_name)
            snapshots = draw_page_shapes.get(sheet_name, [])
            if snapshots:
                _log_unmatched_ooxml_candidates(
                    sheet_name=sheet_name,
                    snapshots=snapshots,
                    drawing_shapes=drawing.shapes if drawing is not None else [],
                    drawing_connectors=drawing.connectors
                    if drawing is not None
                    else [],
                )
                # UNO draw-page snapshots define the canonical emitted order for v1.
                # Unmatched OOXML-only shapes/connectors remain supplemental metadata
                # and are intentionally not appended to the emitted list.
                shape_data[sheet_name] = _build_shapes_from_draw_page(
                    snapshots,
                    drawing_shapes=drawing.shapes if drawing is not None else [],
                    drawing_connectors=drawing.connectors
                    if drawing is not None
                    else [],
                )
                continue
            if drawing is not None:
                shape_data[sheet_name] = _build_shapes_from_ooxml(
                    drawing.shapes,
                    drawing.connectors,
                )
        return shape_data

    def extract_charts(self, *, mode: str) -> dict[str, list[Chart]]:
        """Extract LibreOffice-mode charts for each worksheet.

        Args:
            mode: Requested extraction mode. Only ``"libreoffice"`` is supported.

        Returns:
            Mapping of sheet names to emitted chart models.

        Raises:
            ValueError: If a non-LibreOffice mode is requested.
        """

        if mode != "libreoffice":
            raise ValueError("LibreOfficeRichBackend only supports libreoffice mode.")
        drawings = read_sheet_drawings(self.file_path)
        chart_geometries = self._read_chart_geometries()
        chart_data: ChartData = {}
        for sheet_name, drawing in drawings.items():
            charts: list[Chart] = []
            geometry_matches = _match_chart_geometries(
                drawing.charts,
                chart_geometries.get(sheet_name, []),
            )
            for chart_info, geometry in zip(
                drawing.charts, geometry_matches, strict=False
            ):
                left = chart_info.anchor_left or 0
                top = chart_info.anchor_top or 0
                width = chart_info.anchor_width
                height = chart_info.anchor_height
                confidence = 0.5
                if geometry is not None:
                    left = geometry.left if geometry.left is not None else left
                    top = geometry.top if geometry.top is not None else top
                    width = geometry.width if geometry.width is not None else width
                    height = geometry.height if geometry.height is not None else height
                    confidence = 0.8
                charts.append(
                    Chart(
                        name=chart_info.name,
                        chart_type=chart_info.chart_type,
                        title=chart_info.title,
                        y_axis_title=chart_info.y_axis_title,
                        y_axis_range=chart_info.y_axis_range,
                        w=width,
                        h=height,
                        series=chart_info.series,
                        l=left,
                        t=top,
                        provenance="libreoffice_uno",
                        approximation_level="partial",
                        confidence=confidence,
                    )
                )
            chart_data[sheet_name] = charts
        return chart_data

    def _read_chart_geometries(self) -> dict[str, list[LibreOfficeChartGeometry]]:
        """Load and cache chart geometry snapshots from the LibreOffice session."""

        if self._chart_geometries is not None:
            return self._chart_geometries
        with self._session_factory() as session:
            if hasattr(session, "extract_chart_geometries"):
                self._chart_geometries = (
                    self._extract_chart_geometries_with_optional_workbook_lifecycle(
                        session=session,
                    )
                )
            else:
                self._chart_geometries = {}
        return self._chart_geometries

    def _read_draw_page_shapes(self) -> dict[str, list[LibreOfficeDrawPageShape]]:
        """Load and cache draw-page shape snapshots from the LibreOffice session."""

        if self._draw_page_shapes is not None:
            return self._draw_page_shapes
        with self._session_factory() as session:
            if hasattr(session, "extract_draw_page_shapes"):
                self._draw_page_shapes = (
                    self._extract_draw_page_shapes_with_optional_workbook_lifecycle(
                        session=session,
                    )
                )
            else:
                self._draw_page_shapes = {}
        return self._draw_page_shapes

    def _extract_chart_geometries_with_optional_workbook_lifecycle(
        self,
        *,
        session: _LibreOfficeRichSession,
    ) -> dict[str, list[LibreOfficeChartGeometry]]:
        """Call chart extraction while preserving legacy path-only sessions."""

        if not _supports_workbook_lifecycle(session):
            return session.extract_chart_geometries(self.file_path)
        workbook = session.load_workbook(self.file_path)
        try:
            return session.extract_chart_geometries(workbook)
        finally:
            session.close_workbook(workbook)

    def _extract_draw_page_shapes_with_optional_workbook_lifecycle(
        self,
        *,
        session: _LibreOfficeRichSession,
    ) -> dict[str, list[LibreOfficeDrawPageShape]]:
        """Call draw-page extraction while preserving legacy path-only sessions."""

        if not _supports_workbook_lifecycle(session):
            return session.extract_draw_page_shapes(self.file_path)
        workbook = session.load_workbook(self.file_path)
        try:
            return session.extract_draw_page_shapes(workbook)
        finally:
            session.close_workbook(workbook)


def _build_shapes_from_draw_page(
    snapshots: Sequence[LibreOfficeDrawPageShape],
    *,
    drawing_shapes: Sequence[OoxmlShapeInfo],
    drawing_connectors: Sequence[OoxmlConnectorInfo],
) -> list[Shape | Arrow | SmartArt]:
    """Merge UNO draw-page snapshots with OOXML drawing metadata.

    Args:
        snapshots: LibreOffice draw-page snapshots for one worksheet.
        drawing_shapes: OOXML shape candidates for the worksheet.
        drawing_connectors: OOXML connector candidates for the worksheet.

    Returns:
        Emitted shape and arrow models built from the combined metadata.
    """

    emitted: list[Shape | Arrow | SmartArt] = []
    snapshot_shapes = [snapshot for snapshot in snapshots if not snapshot.is_connector]
    snapshot_connectors = [snapshot for snapshot in snapshots if snapshot.is_connector]
    matched_shapes = _match_shape_infos(snapshot_shapes, drawing_shapes)
    matched_connectors = _match_connector_infos(snapshot_connectors, drawing_connectors)
    drawing_to_shape_id: dict[int, int] = {}
    shape_name_to_id: dict[str, int] = {}
    shape_boxes: dict[int, _ShapeBox] = {}
    next_shape_id = 0
    assigned_shapes: list[
        tuple[LibreOfficeDrawPageShape, OoxmlShapeInfo | None, int]
    ] = []

    for snapshot, shape_info in zip(snapshot_shapes, matched_shapes, strict=False):
        next_shape_id += 1
        shape_id = next_shape_id
        assigned_shapes.append((snapshot, shape_info, shape_id))
        if shape_info is not None:
            drawing_to_shape_id[shape_info.ref.drawing_id] = shape_id
        shape_name_to_id[snapshot.name] = shape_id
        box = _to_shape_box(
            shape_id=shape_id,
            left=snapshot.left,
            top=snapshot.top,
            width=snapshot.width,
            height=snapshot.height,
        )
        if box is not None:
            shape_boxes[shape_id] = box

    shape_index = 0
    connector_index = 0

    for snapshot in snapshots:
        if snapshot.is_connector:
            connector_info = matched_connectors[connector_index]
            connector_index += 1
            begin_id, end_id, approximation_level, confidence = _resolve_connector(
                connector_info,
                uno_connector=snapshot,
                drawing_to_shape_id=drawing_to_shape_id,
                shape_name_to_id=shape_name_to_id,
                shape_boxes=shape_boxes,
            )
            emitted.append(
                Arrow(
                    id=None,
                    text=snapshot.text
                    or (connector_info.text if connector_info else ""),
                    l=_first_int(
                        snapshot.left,
                        connector_info.ref.left if connector_info else None,
                        default=0,
                    ),
                    t=_first_int(
                        snapshot.top,
                        connector_info.ref.top if connector_info else None,
                        default=0,
                    ),
                    w=_first_optional_int(
                        snapshot.width,
                        connector_info.ref.width if connector_info else None,
                    ),
                    h=_first_optional_int(
                        snapshot.height,
                        connector_info.ref.height if connector_info else None,
                    ),
                    rotation=_first_optional_float(
                        snapshot.rotation,
                        connector_info.rotation if connector_info else None,
                    ),
                    begin_arrow_style=connector_info.begin_arrow_style
                    if connector_info is not None
                    else None,
                    end_arrow_style=connector_info.end_arrow_style
                    if connector_info is not None
                    else None,
                    begin_id=begin_id,
                    end_id=end_id,
                    direction=_resolve_direction(
                        connector_info=connector_info,
                        uno_connector=snapshot,
                        begin_id=begin_id,
                        end_id=end_id,
                        shape_boxes=shape_boxes,
                    ),
                    provenance="libreoffice_uno",
                    approximation_level=approximation_level,
                    confidence=confidence,
                )
            )
            continue

        shape_snapshot, shape_info, shape_id = assigned_shapes[shape_index]
        shape_index += 1
        emitted.append(
            Shape(
                id=shape_id,
                text=shape_snapshot.text or (shape_info.text if shape_info else ""),
                l=_first_int(
                    shape_snapshot.left,
                    shape_info.ref.left if shape_info else None,
                    default=0,
                ),
                t=_first_int(
                    shape_snapshot.top,
                    shape_info.ref.top if shape_info else None,
                    default=0,
                ),
                w=_first_optional_int(
                    shape_snapshot.width,
                    shape_info.ref.width if shape_info else None,
                ),
                h=_first_optional_int(
                    shape_snapshot.height,
                    shape_info.ref.height if shape_info else None,
                ),
                rotation=_first_optional_float(
                    shape_snapshot.rotation,
                    shape_info.rotation if shape_info else None,
                ),
                type=shape_info.shape_type
                if shape_info is not None and shape_info.shape_type
                else _shape_type_from_uno(shape_snapshot.shape_type),
                provenance="libreoffice_uno",
                approximation_level="partial",
                confidence=0.75,
            )
        )
    return emitted


def _log_unmatched_ooxml_candidates(
    *,
    sheet_name: str,
    snapshots: Sequence[LibreOfficeDrawPageShape],
    drawing_shapes: Sequence[OoxmlShapeInfo],
    drawing_connectors: Sequence[OoxmlConnectorInfo],
) -> None:
    """Emit a debug log when snapshot-backed sheets drop OOXML-only candidates."""

    snapshot_shapes = [snapshot for snapshot in snapshots if not snapshot.is_connector]
    snapshot_connectors = [snapshot for snapshot in snapshots if snapshot.is_connector]
    unmatched_shape_count = len(drawing_shapes) - sum(
        shape_info is not None
        for shape_info in _match_shape_infos(snapshot_shapes, drawing_shapes)
    )
    unmatched_connector_count = len(drawing_connectors) - sum(
        connector_info is not None
        for connector_info in _match_connector_infos(
            snapshot_connectors, drawing_connectors
        )
    )
    if unmatched_shape_count <= 0 and unmatched_connector_count <= 0:
        return
    logger.debug(
        "Skipping %d OOXML-only shapes and %d OOXML-only connectors on sheet %s "
        "because UNO draw-page snapshots define the canonical emitted order.",
        unmatched_shape_count,
        unmatched_connector_count,
        sheet_name,
    )


def _match_shape_infos(
    snapshots: Sequence[LibreOfficeDrawPageShape],
    candidates: Sequence[OoxmlShapeInfo],
) -> list[OoxmlShapeInfo | None]:
    """Match UNO shape snapshots to OOXML shape metadata."""

    return [
        candidates[index] if index is not None else None
        for index in _match_by_name_then_order(
            [snapshot.name for snapshot in snapshots],
            [candidate.ref.name for candidate in candidates],
        )
    ]


def _match_connector_infos(
    snapshots: Sequence[LibreOfficeDrawPageShape],
    candidates: Sequence[OoxmlConnectorInfo],
) -> list[OoxmlConnectorInfo | None]:
    """Match UNO connector snapshots to OOXML connector metadata."""

    return [
        candidates[index] if index is not None else None
        for index in _match_by_name_then_order(
            [snapshot.name for snapshot in snapshots],
            [candidate.ref.name for candidate in candidates],
        )
    ]


def _match_by_name_then_order(
    snapshot_names: Sequence[str],
    candidate_names: Sequence[str],
) -> list[int | None]:
    """Match snapshot names to candidate names by name first, then by order."""

    matches: list[int | None] = [None] * len(snapshot_names)
    unused = list(range(len(candidate_names)))

    for index, snapshot_name in enumerate(snapshot_names):
        matched_index = next(
            (
                candidate_index
                for candidate_index in unused
                if candidate_names[candidate_index] == snapshot_name
            ),
            None,
        )
        if matched_index is None:
            continue
        matches[index] = matched_index
        unused.remove(matched_index)

    remaining_snapshot_indexes = [
        index for index, match in enumerate(matches) if match is None
    ]
    if remaining_snapshot_indexes:
        for snapshot_index, candidate_index in zip(
            remaining_snapshot_indexes, unused, strict=False
        ):
            matches[snapshot_index] = candidate_index

    return matches


def _shape_type_from_uno(shape_type: str | None) -> str | None:
    """Collapse a fully qualified UNO shape type to its leaf name."""

    if not shape_type:
        return None
    return shape_type.rsplit(".", 1)[-1]


def _match_chart_geometries(
    charts: Sequence[object],
    candidates: Sequence[LibreOfficeChartGeometry],
) -> list[LibreOfficeChartGeometry | None]:
    """Match OOXML chart entries to LibreOffice chart geometry candidates."""

    matches: list[LibreOfficeChartGeometry | None] = [None] * len(charts)
    unused = list(range(len(candidates)))

    for index, chart in enumerate(charts):
        chart_name = getattr(chart, "name", None)
        if not isinstance(chart_name, str):
            continue
        matched_index = next(
            (
                candidate_index
                for candidate_index in unused
                if candidates[candidate_index].name == chart_name
                or candidates[candidate_index].persist_name == chart_name
            ),
            None,
        )
        if matched_index is None:
            continue
        matches[index] = candidates[matched_index]
        unused.remove(matched_index)

    remaining_chart_indexes = [
        index for index, match in enumerate(matches) if match is None
    ]
    if remaining_chart_indexes and len(remaining_chart_indexes) == len(unused):
        for chart_index, candidate_index in zip(
            remaining_chart_indexes, unused, strict=False
        ):
            matches[chart_index] = candidates[candidate_index]

    return matches


def _first_int(*values: int | None, default: int) -> int:
    """Return the first non-``None`` integer or the provided default."""

    for value in values:
        if value is not None:
            return value
    return default


def _first_optional_int(*values: int | None) -> int | None:
    """Return the first non-``None`` integer, if any."""

    for value in values:
        if value is not None:
            return value
    return None


def _first_optional_float(*values: float | None) -> float | None:
    """Return the first non-``None`` float, if any."""

    for value in values:
        if value is not None:
            return value
    return None
