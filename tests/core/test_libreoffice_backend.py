"""Tests for the LibreOffice rich extraction backend."""

from __future__ import annotations

from collections.abc import Callable
import json
from pathlib import Path
import subprocess
from typing import cast
from xml.etree import ElementTree
from zipfile import ZipFile

from _pytest.monkeypatch import MonkeyPatch
import pytest

from exstruct.core.backends.libreoffice_backend import (
    LibreOfficeRichBackend,
    _resolve_direction,
)
from exstruct.core.libreoffice import (
    LibreOfficeChartGeometry,
    LibreOfficeDrawPageShape,
    LibreOfficeSession,
    LibreOfficeSessionConfig,
    LibreOfficeUnavailableError,
    _parse_chart_payload,
    _parse_draw_page_payload,
    _probe_uno_bridge_handshake,
    _python_supports_libreoffice_bridge,
    _resolve_python_path,
)
from exstruct.core.ooxml_drawing import (
    DrawingConnectorRef,
    DrawingShapeRef,
    OoxmlConnectorInfo,
    OoxmlShapeInfo,
    SheetDrawingData,
    _extract_chart_series,
    _merge_anchor_geometry,
    _parse_connector_node,
    _read_relationships,
    read_sheet_drawings,
)
from exstruct.models import Arrow, Shape


class _DummySession:
    """Dummy LibreOffice session used in tests."""

    def __enter__(self) -> _DummySession:
        """Return the test double as the context-manager result."""

        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        """Accept context-manager exit arguments without suppressing errors."""

        _ = exc_type
        _ = exc
        _ = tb

    def load_workbook(self, file_path: Path) -> object:
        """Return a lightweight workbook token for tests."""

        return {"file_path": str(file_path)}

    def close_workbook(self, workbook: object) -> None:
        """Accept a workbook token without additional cleanup."""

        _ = workbook

    def extract_chart_geometries(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeChartGeometry]]:
        """Provide chart geometry data for this test double."""

        _ = file_path
        return {}

    def extract_draw_page_shapes(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeDrawPageShape]]:
        """Provide draw-page shape data for this test double."""

        _ = file_path
        return {}


def _dummy_session_factory() -> LibreOfficeSession:
    """Return a dummy LibreOffice session test double."""

    return cast(LibreOfficeSession, _DummySession())


class _ChartGeometrySession(_DummySession):
    """Session double that returns fixed chart geometry data."""

    def extract_chart_geometries(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeChartGeometry]]:
        """Provide chart geometry data for this test double."""

        _ = file_path
        return {
            "Sheet1": [
                LibreOfficeChartGeometry(
                    name="グラフ 1",
                    persist_name="Object 1",
                    left=370,
                    top=26,
                    width=351,
                    height=214,
                )
            ]
        }


def _chart_geometry_session_factory() -> LibreOfficeSession:
    """Return a LibreOffice session double with chart geometries."""

    return cast(LibreOfficeSession, _ChartGeometrySession())


class _DrawPageSession(_DummySession):
    """Session double that returns fixed draw-page shapes."""

    def __init__(self, payload: dict[str, list[LibreOfficeDrawPageShape]]) -> None:
        """Initialize the test double."""

        self._payload = payload

    def extract_draw_page_shapes(
        self, file_path: Path
    ) -> dict[str, list[LibreOfficeDrawPageShape]]:
        """Provide draw-page shape data for this test double."""

        _ = file_path
        return self._payload


class _FakeLibreOfficeProcess:
    """Simple soffice process double used for startup and shutdown tests."""

    def __init__(
        self,
        *,
        stderr: str = "",
        returncode: int | None = None,
        wait_timeouts: int = 0,
    ) -> None:
        """Initialize the fake process with optional stderr and wait behavior."""

        self.stderr = stderr
        self.returncode = returncode
        self.wait_timeouts = wait_timeouts
        self.terminate_called = False
        self.kill_called = False
        self.wait_calls = 0
        self.args: list[str] = []

    def poll(self) -> int | None:
        """Return the configured process status."""

        return self.returncode

    def terminate(self) -> None:
        """Record a termination request."""

        self.terminate_called = True
        if self.returncode is None:
            self.returncode = 0

    def wait(self, *, timeout: float) -> int:
        """Return the configured exit code or raise a timeout when requested."""

        _ = timeout
        self.wait_calls += 1
        if self.wait_timeouts > 0:
            self.wait_timeouts -= 1
            raise subprocess.TimeoutExpired(cmd="soffice", timeout=timeout)
        if self.returncode is None:
            self.returncode = 0
        return self.returncode

    def kill(self) -> None:
        """Record a forced kill request."""

        self.kill_called = True
        self.returncode = -9

    def communicate(self, timeout: float | None = None) -> tuple[str, str]:
        """Return the configured stderr payload."""

        _ = timeout
        return ("", self.stderr)


def test_libreoffice_backend_extracts_connector_graph_from_sample() -> None:
    """Verify that LibreOffice backend extracts connector graph from sample."""

    backend = LibreOfficeRichBackend(
        Path("sample/flowchart/sample-shape-connector.xlsx"),
        session_factory=_dummy_session_factory,
    )
    shape_data = backend.extract_shapes(mode="libreoffice")
    assert shape_data
    first_sheet = next(iter(shape_data.values()))
    connectors = [shape for shape in first_sheet if isinstance(shape, Arrow)]
    resolved = [
        connector
        for connector in connectors
        if connector.begin_id is not None and connector.end_id is not None
    ]
    assert len(connectors) >= 10
    assert len(resolved) >= 10
    assert all(connector.provenance == "libreoffice_uno" for connector in resolved)


def test_libreoffice_backend_extracts_chart_from_sample() -> None:
    """Verify that LibreOffice backend extracts chart from sample."""

    backend = LibreOfficeRichBackend(
        Path("sample/basic/sample.xlsx"),
        session_factory=_chart_geometry_session_factory,
    )
    chart_data = backend.extract_charts(mode="libreoffice")
    charts = chart_data["Sheet1"]
    assert len(charts) >= 1
    chart = charts[0]
    assert chart.title == "売上データ"
    assert len(chart.series) == 3
    assert (chart.l, chart.t, chart.w, chart.h) == (370, 26, 351, 214)
    assert chart.provenance == "libreoffice_uno"
    assert chart.approximation_level == "partial"
    assert chart.confidence == 0.8


def test_libreoffice_backend_uses_draw_page_shapes_without_ooxml(
    monkeypatch: MonkeyPatch,
) -> None:
    """Verify that LibreOffice backend uses draw page shapes without OOXML."""

    payload = {
        "Sheet1": [
            LibreOfficeDrawPageShape(
                name="Start",
                shape_type="com.sun.star.drawing.CustomShape",
                text="S",
                left=10,
                top=20,
                width=30,
                height=40,
            ),
            LibreOfficeDrawPageShape(
                name="Connector",
                shape_type="com.sun.star.drawing.ConnectorShape",
                left=40,
                top=30,
                width=50,
                height=60,
                is_connector=True,
                start_shape_name="Start",
                end_shape_name="End",
            ),
            LibreOfficeDrawPageShape(
                name="End",
                shape_type="com.sun.star.drawing.CustomShape",
                text="E",
                left=100,
                top=120,
                width=30,
                height=40,
            ),
        ]
    }
    monkeypatch.setattr(
        "exstruct.core.backends.libreoffice_backend.read_sheet_drawings",
        lambda _path: {"Sheet1": SheetDrawingData()},
    )
    backend = LibreOfficeRichBackend(
        Path("sample/flowchart/sample-shape-connector.xlsx"),
        session_factory=lambda: cast(LibreOfficeSession, _DrawPageSession(payload)),
    )

    shape_data = backend.extract_shapes(mode="libreoffice")

    items = shape_data["Sheet1"]
    assert [item.kind for item in items] == ["shape", "arrow", "shape"]
    first_shape = cast(Shape, items[0])
    connector = cast(Arrow, items[1])
    second_shape = cast(Shape, items[2])
    assert first_shape.id == 1
    assert first_shape.type == "CustomShape"
    assert second_shape.id == 2
    assert connector.begin_id == 1
    assert connector.end_id == 2
    assert connector.approximation_level == "direct"
    assert connector.confidence == 0.9


def test_libreoffice_backend_prefers_ooxml_refs_over_uno_direct_refs(
    monkeypatch: MonkeyPatch,
) -> None:
    """Verify that LibreOffice backend prefers OOXML refs over uno direct refs."""

    payload = {
        "Sheet1": [
            LibreOfficeDrawPageShape(
                name="Start",
                shape_type="com.sun.star.drawing.CustomShape",
                left=10,
                top=20,
                width=30,
                height=40,
            ),
            LibreOfficeDrawPageShape(
                name="Connector",
                shape_type="com.sun.star.drawing.ConnectorShape",
                left=40,
                top=30,
                width=50,
                height=60,
                is_connector=True,
                start_shape_name="Start",
                end_shape_name="End",
            ),
            LibreOfficeDrawPageShape(
                name="End",
                shape_type="com.sun.star.drawing.CustomShape",
                left=100,
                top=120,
                width=30,
                height=40,
            ),
        ]
    }
    monkeypatch.setattr(
        "exstruct.core.backends.libreoffice_backend.read_sheet_drawings",
        lambda _path: {
            "Sheet1": SheetDrawingData(
                shapes=[
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=10,
                            name="Start",
                            kind="shape",
                            left=10,
                            top=20,
                            width=30,
                            height=40,
                        )
                    ),
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=20,
                            name="End",
                            kind="shape",
                            left=100,
                            top=120,
                            width=30,
                            height=40,
                        )
                    ),
                ],
                connectors=[
                    OoxmlConnectorInfo(
                        ref=DrawingShapeRef(
                            drawing_id=30,
                            name="Connector",
                            kind="connector",
                            left=40,
                            top=30,
                            width=50,
                            height=60,
                        ),
                        connection=DrawingConnectorRef(
                            drawing_id=30,
                            start_drawing_id=20,
                            end_drawing_id=10,
                        ),
                    )
                ],
            )
        },
    )
    backend = LibreOfficeRichBackend(
        Path("sample/flowchart/sample-shape-connector.xlsx"),
        session_factory=lambda: cast(LibreOfficeSession, _DrawPageSession(payload)),
    )

    shape_data = backend.extract_shapes(mode="libreoffice")

    connector = next(
        shape for shape in shape_data["Sheet1"] if isinstance(shape, Arrow)
    )
    assert connector.begin_id == 2
    assert connector.end_id == 1
    assert connector.approximation_level == "direct"
    assert connector.confidence == 1.0


def test_libreoffice_backend_ignores_unmatched_ooxml_when_draw_page_exists(
    monkeypatch: MonkeyPatch,
) -> None:
    """Verify that snapshot-backed sheets do not append OOXML-only items in v1."""

    payload = {
        "Sheet1": [
            LibreOfficeDrawPageShape(
                name="Start",
                shape_type="com.sun.star.drawing.CustomShape",
                text="S",
                left=10,
                top=20,
                width=30,
                height=40,
            )
        ]
    }
    monkeypatch.setattr(
        "exstruct.core.backends.libreoffice_backend.read_sheet_drawings",
        lambda _path: {
            "Sheet1": SheetDrawingData(
                shapes=[
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=10,
                            name="Start",
                            kind="shape",
                            left=10,
                            top=20,
                            width=30,
                            height=40,
                        ),
                        text="S",
                        shape_type="rect",
                    ),
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=20,
                            name="OOXML-only",
                            kind="shape",
                            left=100,
                            top=120,
                            width=30,
                            height=40,
                        ),
                        text="extra",
                        shape_type="ellipse",
                    ),
                ],
                connectors=[
                    OoxmlConnectorInfo(
                        ref=DrawingShapeRef(
                            drawing_id=30,
                            name="Connector",
                            kind="connector",
                            left=40,
                            top=30,
                            width=50,
                            height=60,
                        ),
                        connection=DrawingConnectorRef(
                            drawing_id=30,
                            start_drawing_id=10,
                            end_drawing_id=20,
                        ),
                    )
                ],
            )
        },
    )
    backend = LibreOfficeRichBackend(
        Path("sample/flowchart/sample-shape-connector.xlsx"),
        session_factory=lambda: cast(LibreOfficeSession, _DrawPageSession(payload)),
    )

    items = backend.extract_shapes(mode="libreoffice")["Sheet1"]

    assert len(items) == 1
    shape = cast(Shape, items[0])
    assert shape.kind == "shape"
    assert shape.id == 1
    assert shape.text == "S"
    assert shape.type == "rect"


def test_libreoffice_backend_logs_unmatched_ooxml_when_draw_page_exists(
    monkeypatch: MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    """Verify that snapshot-backed sheets log dropped OOXML-only candidate counts."""

    payload = {
        "Sheet1": [
            LibreOfficeDrawPageShape(
                name="Start",
                shape_type="com.sun.star.drawing.CustomShape",
                text="S",
                left=10,
                top=20,
                width=30,
                height=40,
            )
        ]
    }
    monkeypatch.setattr(
        "exstruct.core.backends.libreoffice_backend.read_sheet_drawings",
        lambda _path: {
            "Sheet1": SheetDrawingData(
                shapes=[
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=10,
                            name="Start",
                            kind="shape",
                            left=10,
                            top=20,
                            width=30,
                            height=40,
                        )
                    ),
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=20,
                            name="OOXML-only",
                            kind="shape",
                            left=100,
                            top=120,
                            width=30,
                            height=40,
                        )
                    ),
                ],
                connectors=[
                    OoxmlConnectorInfo(
                        ref=DrawingShapeRef(
                            drawing_id=30,
                            name="Connector",
                            kind="connector",
                            left=40,
                            top=30,
                            width=50,
                            height=60,
                        ),
                        connection=DrawingConnectorRef(
                            drawing_id=30,
                            start_drawing_id=10,
                            end_drawing_id=20,
                        ),
                    )
                ],
            )
        },
    )
    backend = LibreOfficeRichBackend(
        Path("sample/flowchart/sample-shape-connector.xlsx"),
        session_factory=lambda: cast(LibreOfficeSession, _DrawPageSession(payload)),
    )

    with caplog.at_level("DEBUG", logger="exstruct.core.backends.libreoffice_backend"):
        backend.extract_shapes(mode="libreoffice")

    assert (
        "Skipping 1 OOXML-only shapes and 1 OOXML-only connectors on sheet Sheet1"
        in caplog.text
    )


def test_ooxml_connector_tail_end_maps_to_begin_arrow_style() -> None:
    """Verify that OOXML connector tail end maps to begin arrow style."""

    node = ElementTree.fromstring(
        """
        <xdr:cxnSp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xdr:nvCxnSpPr>
            <xdr:cNvPr id="7" name="Connector 1" />
            <xdr:cNvCxnSpPr />
          </xdr:nvCxnSpPr>
          <xdr:spPr>
            <a:xfrm>
              <a:off x="0" y="0" />
              <a:ext cx="12700" cy="12700" />
            </a:xfrm>
            <a:ln>
              <a:tailEnd type="triangle" />
            </a:ln>
          </xdr:spPr>
        </xdr:cxnSp>
        """
    )
    anchor = ElementTree.fromstring(
        """
        <xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
          <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        </xdr:twoCellAnchor>
        """
    )
    connector = _parse_connector_node(anchor, node)
    assert connector is not None
    assert connector.begin_arrow_style == 2
    assert connector.end_arrow_style is None


def test_ooxml_connector_head_end_maps_to_end_arrow_style() -> None:
    """Verify that OOXML connector head end maps to end arrow style."""

    node = ElementTree.fromstring(
        """
        <xdr:cxnSp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xdr:nvCxnSpPr>
            <xdr:cNvPr id="7" name="Connector 1" />
            <xdr:cNvCxnSpPr />
          </xdr:nvCxnSpPr>
          <xdr:spPr>
            <a:xfrm>
              <a:off x="0" y="0" />
              <a:ext cx="12700" cy="12700" />
            </a:xfrm>
            <a:ln>
              <a:headEnd type="triangle" />
            </a:ln>
          </xdr:spPr>
        </xdr:cxnSp>
        """
    )
    anchor = ElementTree.fromstring(
        """
        <xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
          <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        </xdr:twoCellAnchor>
        """
    )
    connector = _parse_connector_node(anchor, node)
    assert connector is not None
    assert connector.begin_arrow_style is None
    assert connector.end_arrow_style == 2


def test_libreoffice_backend_combines_ooxml_and_uno_connector_endpoints(
    monkeypatch: MonkeyPatch,
) -> None:
    """Verify that missing OOXML endpoints are completed from UNO direct refs."""

    payload = {
        "Sheet1": [
            LibreOfficeDrawPageShape(
                name="Start",
                shape_type="com.sun.star.drawing.CustomShape",
                left=10,
                top=20,
                width=30,
                height=40,
            ),
            LibreOfficeDrawPageShape(
                name="Connector",
                shape_type="com.sun.star.drawing.ConnectorShape",
                left=40,
                top=30,
                width=0,
                height=60,
                is_connector=True,
                start_shape_name="Start",
                end_shape_name="End",
            ),
            LibreOfficeDrawPageShape(
                name="End",
                shape_type="com.sun.star.drawing.CustomShape",
                left=100,
                top=120,
                width=30,
                height=40,
            ),
        ]
    }
    monkeypatch.setattr(
        "exstruct.core.backends.libreoffice_backend.read_sheet_drawings",
        lambda _path: {
            "Sheet1": SheetDrawingData(
                shapes=[
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=10,
                            name="Start",
                            kind="shape",
                            left=10,
                            top=20,
                            width=30,
                            height=40,
                        )
                    ),
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=20,
                            name="End",
                            kind="shape",
                            left=100,
                            top=120,
                            width=30,
                            height=40,
                        )
                    ),
                ],
                connectors=[
                    OoxmlConnectorInfo(
                        ref=DrawingShapeRef(
                            drawing_id=30,
                            name="Connector",
                            kind="connector",
                            left=40,
                            top=30,
                            width=0,
                            height=60,
                        ),
                        connection=DrawingConnectorRef(
                            drawing_id=30,
                            start_drawing_id=10,
                            end_drawing_id=None,
                        ),
                        direction_dx=0,
                        direction_dy=0,
                    )
                ],
            )
        },
    )
    backend = LibreOfficeRichBackend(
        Path("sample/flowchart/sample-shape-connector.xlsx"),
        session_factory=lambda: cast(LibreOfficeSession, _DrawPageSession(payload)),
    )

    connector = next(
        shape
        for shape in backend.extract_shapes(mode="libreoffice")["Sheet1"]
        if isinstance(shape, Arrow)
    )

    assert connector.begin_id == 1
    assert connector.end_id == 2
    assert connector.approximation_level == "partial"
    assert connector.confidence == 0.9
    assert connector.direction == "N"


def test_ooxml_zero_delta_direction_falls_back_to_uno_geometry() -> None:
    """Verify that a zero-length OOXML delta does not force an eastward direction."""

    connector = OoxmlConnectorInfo(
        ref=DrawingShapeRef(
            drawing_id=1,
            name="Connector",
            kind="connector",
            left=0,
            top=0,
            width=0,
            height=0,
        ),
        connection=DrawingConnectorRef(
            drawing_id=1,
            start_drawing_id=None,
            end_drawing_id=None,
        ),
        direction_dx=0,
        direction_dy=0,
    )
    uno_connector = LibreOfficeDrawPageShape(
        name="Connector",
        is_connector=True,
        width=0,
        height=25,
    )

    assert (
        _resolve_direction(connector_info=connector, uno_connector=uno_connector) == "N"
    )


def test_read_sheet_drawings_uses_anchor_fallback_for_chart_geometry() -> None:
    """Verify that zero-xfrm charts still get positive anchor geometry from OOXML."""

    drawings = read_sheet_drawings(Path("sample/basic/sample.xlsx"))
    chart = drawings["Sheet1"].charts[0]

    assert chart.anchor_left is not None and chart.anchor_left > 0
    assert chart.anchor_top is not None and chart.anchor_top > 0
    assert chart.anchor_width is not None and chart.anchor_width > 0
    assert chart.anchor_height is not None and chart.anchor_height > 0


def test_read_relationships_keeps_type_metadata() -> None:
    """Verify that OOXML relationship parsing preserves relationship types."""

    with ZipFile(Path("sample/basic/sample.xlsx")) as archive:
        relationships = _read_relationships(
            archive,
            "xl/drawings/_rels/drawing1.xml.rels",
        )

    assert any(
        relationship.relationship_type.endswith("/chart")
        and relationship.target.startswith("xl/charts/")
        for relationship in relationships.values()
    )


def test_merge_anchor_geometry_prefers_parent_anchor_for_left_top() -> None:
    """Verify that anchor placement wins over child-transform offsets."""

    anchor = ElementTree.fromstring(
        """
        <xdr:oneCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <xdr:from>
            <xdr:col>2</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>3</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
          </xdr:from>
          <xdr:ext cx="127000" cy="254000" />
        </xdr:oneCellAnchor>
        """
    )

    left, top, width, height = _merge_anchor_geometry(
        anchor,
        left=999,
        top=777,
        width=10,
        height=20,
    )

    assert left == 96
    assert top == 45
    assert width == 10
    assert height == 20


def test_extract_chart_series_supports_scatter_xval_yval() -> None:
    """Verify that scatter/bubble series use xVal/yVal references when present."""

    chart_root = ElementTree.fromstring(
        """
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:scatterChart>
                <c:ser>
                  <c:tx><c:v>Series 1</c:v></c:tx>
                  <c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$5</c:f></c:numRef></c:xVal>
                  <c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f></c:numRef></c:yVal>
                </c:ser>
              </c:scatterChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
        """
    )

    series = _extract_chart_series(chart_root)

    assert len(series) == 1
    assert series[0].x_range == "Sheet1!$A$2:$A$5"
    assert series[0].y_range == "Sheet1!$B$2:$B$5"


def test_libreoffice_session_skips_temp_profile_when_version_probe_fails(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that version probe failure happens before temp-profile creation."""

    mkdtemp_called = False
    removed_paths: list[Path] = []
    created_dir = tmp_path / "lo-profile"
    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    soffice_path.write_text("", encoding="utf-8")
    python_path.write_text("", encoding="utf-8")

    def _fake_mkdtemp(*, prefix: str, **kwargs: object) -> str:
        """Provide a fake mkdtemp implementation for this test."""

        nonlocal mkdtemp_called
        _ = prefix
        _ = kwargs
        mkdtemp_called = True
        created_dir.mkdir(parents=True, exist_ok=True)
        return str(created_dir)

    def _fake_rmtree(path: Path | str, *, ignore_errors: bool) -> None:
        """Provide a fake rmtree implementation for this test."""

        _ = ignore_errors
        removed_paths.append(Path(path))

    def _fake_run(*_args: object, **_kwargs: object) -> object:
        """Provide a fake run implementation for this test."""

        raise subprocess.TimeoutExpired(cmd="soffice --version", timeout=1.0)

    monkeypatch.setattr(
        "exstruct.core.libreoffice.mkdtemp",
        cast(Callable[..., str], _fake_mkdtemp),
    )
    monkeypatch.setattr("exstruct.core.libreoffice.shutil.rmtree", _fake_rmtree)
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)
    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )

    with pytest.raises(LibreOfficeUnavailableError):
        session.__enter__()

    assert mkdtemp_called is False
    assert removed_paths == []
    assert session._temp_profile_dir is None


def test_libreoffice_session_enters_with_isolated_profile(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the primary startup path uses an isolated temp profile."""

    cleaned_paths: list[Path] = []
    created_dir = tmp_path / "lo-profile"
    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    process = _FakeLibreOfficeProcess()
    soffice_path.write_text("", encoding="utf-8")
    python_path.write_text("", encoding="utf-8")

    def _fake_mkdtemp(*, prefix: str, **kwargs: object) -> str:
        _ = prefix
        _ = kwargs
        created_dir.mkdir(parents=True, exist_ok=True)
        return str(created_dir)

    def _fake_run(
        args: list[str], **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    def _fake_popen(args: list[str], **_kwargs: object) -> _FakeLibreOfficeProcess:
        process.args = list(args)
        return process

    monkeypatch.setattr(
        "exstruct.core.libreoffice.mkdtemp",
        cast(Callable[..., str], _fake_mkdtemp),
    )
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.Popen", _fake_popen)
    monkeypatch.setattr(
        "exstruct.core.libreoffice._wait_for_socket",
        lambda **_kwargs: None,
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._cleanup_profile_dir",
        lambda path: cleaned_paths.append(path),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )

    session.__enter__()

    assert session._temp_profile_dir == created_dir
    assert cast(object, session._soffice_process) is process
    assert any(
        arg == f"-env:UserInstallation={created_dir.as_uri()}" for arg in process.args
    )

    session.__exit__(None, None, None)

    assert process.terminate_called is True
    assert cleaned_paths == [created_dir]


def test_probe_uno_bridge_handshake_uses_bridge_script(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that startup handshake runs the bundled bridge with fixed args/env."""

    python_path = tmp_path / "python.exe"
    python_path.write_text("", encoding="utf-8")
    process = _FakeLibreOfficeProcess()
    calls: dict[str, object] = {}

    def _fake_run(
        args: list[str], **kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        calls["args"] = list(args)
        calls["env"] = kwargs["env"]
        calls["timeout"] = kwargs["timeout"]
        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)

    _probe_uno_bridge_handshake(
        python_path=python_path,
        host="127.0.0.1",
        port=42001,
        timeout_sec=1.5,
        process=cast(subprocess.Popen[str], process),
    )

    args = cast(list[str], calls["args"])
    assert args[0] == str(python_path.resolve())
    assert args[1].endswith("_libreoffice_bridge.py")
    assert "--handshake" in args
    assert args[args.index("--host") + 1] == "127.0.0.1"
    assert args[args.index("--port") + 1] == "42001"
    assert args[args.index("--connect-timeout") + 1] == "1.5"
    env = cast(dict[str, str], calls["env"])
    assert env["PYTHONIOENCODING"] == "utf-8"
    assert calls["timeout"] == 1.5


def test_probe_uno_bridge_handshake_reports_bridge_failures(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that startup handshake surfaces actionable bridge failures."""

    python_path = tmp_path / "python.exe"
    python_path.write_text("", encoding="utf-8")

    def _fake_run(
        args: list[str], **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        raise subprocess.CalledProcessError(
            returncode=1,
            cmd=args,
            stderr="Connector could not be established",
        )

    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)

    with pytest.raises(
        LibreOfficeUnavailableError,
        match="UNO bridge handshake failed. \\(Connector could not be established\\)",
    ):
        _probe_uno_bridge_handshake(
            python_path=python_path,
            host="127.0.0.1",
            port=42001,
            timeout_sec=1.0,
            process=cast(subprocess.Popen[str], _FakeLibreOfficeProcess()),
        )


def test_libreoffice_session_retries_port_within_startup_attempt(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that one startup strategy retries with a fresh port before fallback."""

    cleaned_paths: list[Path] = []
    created_dir = tmp_path / "lo-profile"
    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    first_process = _FakeLibreOfficeProcess(stderr="Address already in use")
    second_process = _FakeLibreOfficeProcess()
    spawned: list[_FakeLibreOfficeProcess] = []
    reserved_ports = [42001, 42002]
    soffice_path.write_text("", encoding="utf-8")
    python_path.write_text("", encoding="utf-8")

    def _fake_mkdtemp(*, prefix: str, **kwargs: object) -> str:
        _ = prefix
        _ = kwargs
        created_dir.mkdir(parents=True, exist_ok=True)
        return str(created_dir)

    def _fake_run(
        args: list[str], **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    def _fake_popen(args: list[str], **_kwargs: object) -> _FakeLibreOfficeProcess:
        process = (first_process, second_process)[len(spawned)]
        process.args = list(args)
        spawned.append(process)
        return process

    def _fake_wait_for_socket(
        *,
        host: str,
        port: int,
        timeout_sec: float,
        process: object,
    ) -> None:
        _ = host
        _ = port
        _ = timeout_sec
        if process is first_process:
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: soffice exited during startup."
            )

    monkeypatch.setattr(
        "exstruct.core.libreoffice.mkdtemp",
        cast(Callable[..., str], _fake_mkdtemp),
    )
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.Popen", _fake_popen)
    monkeypatch.setattr(
        "exstruct.core.libreoffice._wait_for_socket",
        _fake_wait_for_socket,
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._reserve_tcp_port",
        lambda: reserved_ports.pop(0),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._cleanup_profile_dir",
        lambda path: cleaned_paths.append(path),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )

    session.__enter__()

    assert spawned == [first_process, second_process]
    assert first_process.terminate_called is True
    assert all(
        arg.startswith("--accept=socket,host=127.0.0.1,port=42001")
        for arg in first_process.args
        if arg.startswith("--accept=")
    )
    assert all(
        arg.startswith("--accept=socket,host=127.0.0.1,port=42002")
        for arg in second_process.args
        if arg.startswith("--accept=")
    )
    assert any(arg.startswith("-env:UserInstallation=") for arg in first_process.args)
    assert any(arg.startswith("-env:UserInstallation=") for arg in second_process.args)
    assert session._temp_profile_dir == created_dir
    assert cast(object, session._soffice_process) is second_process

    session.__exit__(None, None, None)

    assert second_process.terminate_called is True
    assert cleaned_paths == [created_dir]


def test_libreoffice_session_retries_when_uno_handshake_finds_wrong_listener(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that startup retries when socket accept succeeds but UNO handshake fails."""

    cleaned_paths: list[Path] = []
    created_dir = tmp_path / "lo-profile"
    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    first_process = _FakeLibreOfficeProcess()
    second_process = _FakeLibreOfficeProcess()
    spawned: list[_FakeLibreOfficeProcess] = []
    reserved_ports = [42001, 42002]
    handshake_ports: list[int] = []
    soffice_path.write_text("", encoding="utf-8")
    python_path.write_text("", encoding="utf-8")

    def _fake_mkdtemp(*, prefix: str, **kwargs: object) -> str:
        _ = prefix
        _ = kwargs
        created_dir.mkdir(parents=True, exist_ok=True)
        return str(created_dir)

    def _fake_run(
        args: list[str], **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    def _fake_popen(args: list[str], **_kwargs: object) -> _FakeLibreOfficeProcess:
        process = (first_process, second_process)[len(spawned)]
        process.args = list(args)
        spawned.append(process)
        return process

    def _fake_probe_uno_bridge_handshake(
        *,
        python_path: Path,
        host: str,
        port: int,
        timeout_sec: float,
        process: object,
    ) -> None:
        _ = python_path
        _ = host
        _ = timeout_sec
        handshake_ports.append(port)
        if process is first_process:
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: UNO bridge handshake failed. (connection refused)"
            )

    monkeypatch.setattr(
        "exstruct.core.libreoffice.mkdtemp",
        cast(Callable[..., str], _fake_mkdtemp),
    )
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.Popen", _fake_popen)
    monkeypatch.setattr(
        "exstruct.core.libreoffice._wait_for_socket",
        lambda **_kwargs: None,
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._probe_uno_bridge_handshake",
        _fake_probe_uno_bridge_handshake,
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._reserve_tcp_port",
        lambda: reserved_ports.pop(0),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._cleanup_profile_dir",
        lambda path: cleaned_paths.append(path),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )

    session.__enter__()

    assert spawned == [first_process, second_process]
    assert handshake_ports == [42001, 42002]
    assert first_process.terminate_called is True
    assert session._temp_profile_dir == created_dir
    assert cast(object, session._soffice_process) is second_process

    session.__exit__(None, None, None)

    assert second_process.terminate_called is True
    assert cleaned_paths == [created_dir]


def test_libreoffice_session_retries_without_temp_profile_after_startup_failure(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that shared-profile retry runs after isolated startup failure."""

    cleaned_paths: list[Path] = []
    created_dir = tmp_path / "lo-profile"
    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    first_process = _FakeLibreOfficeProcess(stderr="javaldx failed!")
    second_process = _FakeLibreOfficeProcess(stderr="javaldx failed!")
    third_process = _FakeLibreOfficeProcess(stderr="javaldx failed!")
    fourth_process = _FakeLibreOfficeProcess()
    spawned: list[_FakeLibreOfficeProcess] = []
    soffice_path.write_text("", encoding="utf-8")
    python_path.write_text("", encoding="utf-8")

    def _fake_mkdtemp(*, prefix: str, **kwargs: object) -> str:
        _ = prefix
        _ = kwargs
        created_dir.mkdir(parents=True, exist_ok=True)
        return str(created_dir)

    def _fake_run(
        args: list[str], **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    def _fake_popen(args: list[str], **_kwargs: object) -> _FakeLibreOfficeProcess:
        process = (
            first_process,
            second_process,
            third_process,
            fourth_process,
        )[len(spawned)]
        process.args = list(args)
        spawned.append(process)
        return process

    def _fake_wait_for_socket(
        *,
        host: str,
        port: int,
        timeout_sec: float,
        process: object,
    ) -> None:
        _ = host
        _ = port
        _ = timeout_sec
        if process in (first_process, second_process, third_process):
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: soffice socket startup timed out."
            )

    monkeypatch.setattr(
        "exstruct.core.libreoffice.mkdtemp",
        cast(Callable[..., str], _fake_mkdtemp),
    )
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.Popen", _fake_popen)
    monkeypatch.setattr(
        "exstruct.core.libreoffice._wait_for_socket",
        _fake_wait_for_socket,
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._cleanup_profile_dir",
        lambda path: cleaned_paths.append(path),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )

    session.__enter__()

    assert spawned == [first_process, second_process, third_process, fourth_process]
    assert first_process.terminate_called is True
    assert second_process.terminate_called is True
    assert third_process.terminate_called is True
    assert cleaned_paths == [created_dir]
    assert any(arg.startswith("-env:UserInstallation=") for arg in first_process.args)
    assert any(arg.startswith("-env:UserInstallation=") for arg in second_process.args)
    assert any(arg.startswith("-env:UserInstallation=") for arg in third_process.args)
    assert all(
        not arg.startswith("-env:UserInstallation=") for arg in fourth_process.args
    )
    assert session._temp_profile_dir is None
    assert cast(object, session._soffice_process) is fourth_process

    session.__exit__(None, None, None)

    assert fourth_process.terminate_called is True


def test_libreoffice_session_reports_both_startup_attempt_failures(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that final startup errors include isolated and shared attempt detail."""

    cleaned_paths: list[Path] = []
    created_dir = tmp_path / "lo-profile"
    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    first_process = _FakeLibreOfficeProcess(stderr="javaldx failed!")
    second_process = _FakeLibreOfficeProcess(stderr="javaldx failed!")
    third_process = _FakeLibreOfficeProcess(stderr="javaldx failed!")
    fourth_process = _FakeLibreOfficeProcess(
        stderr="User installation could not be completed.",
        returncode=1,
    )
    fifth_process = _FakeLibreOfficeProcess(
        stderr="User installation could not be completed.",
        returncode=1,
    )
    sixth_process = _FakeLibreOfficeProcess(
        stderr="User installation could not be completed.",
        returncode=1,
    )
    spawned: list[_FakeLibreOfficeProcess] = []
    soffice_path.write_text("", encoding="utf-8")
    python_path.write_text("", encoding="utf-8")

    def _fake_mkdtemp(*, prefix: str, **kwargs: object) -> str:
        _ = prefix
        _ = kwargs
        created_dir.mkdir(parents=True, exist_ok=True)
        return str(created_dir)

    def _fake_run(
        args: list[str], **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    def _fake_popen(args: list[str], **_kwargs: object) -> _FakeLibreOfficeProcess:
        process = (
            first_process,
            second_process,
            third_process,
            fourth_process,
            fifth_process,
            sixth_process,
        )[len(spawned)]
        process.args = list(args)
        spawned.append(process)
        return process

    def _fake_wait_for_socket(
        *,
        host: str,
        port: int,
        timeout_sec: float,
        process: object,
    ) -> None:
        _ = host
        _ = port
        _ = timeout_sec
        if process in (first_process, second_process, third_process):
            raise LibreOfficeUnavailableError(
                "LibreOffice runtime is unavailable: soffice socket startup timed out."
            )
        raise LibreOfficeUnavailableError(
            "LibreOffice runtime is unavailable: soffice exited during startup."
        )

    monkeypatch.setattr(
        "exstruct.core.libreoffice.mkdtemp",
        cast(Callable[..., str], _fake_mkdtemp),
    )
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.Popen", _fake_popen)
    monkeypatch.setattr(
        "exstruct.core.libreoffice._wait_for_socket",
        _fake_wait_for_socket,
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._cleanup_profile_dir",
        lambda path: cleaned_paths.append(path),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )

    with pytest.raises(LibreOfficeUnavailableError) as excinfo:
        session.__enter__()

    message = str(excinfo.value)
    assert "isolated-profile: attempt 1/3: soffice socket startup timed out." in message
    assert "attempt 3/3: soffice socket startup timed out." in message
    assert "stderr=javaldx failed!" in message
    assert "shared-profile: attempt 1/3: soffice exited during startup." in message
    assert "attempt 3/3: soffice exited during startup." in message
    assert "stderr=User installation could not be completed." in message
    assert cleaned_paths == [created_dir]
    assert session._soffice_process is None
    assert session._temp_profile_dir is None


def test_resolve_python_path_prefers_override(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that an explicit Python override wins over auto-detection."""

    override_path = tmp_path / "custom-python"
    override_path.write_text("", encoding="utf-8")
    monkeypatch.setenv("EXSTRUCT_LIBREOFFICE_PYTHON_PATH", str(override_path))

    def _fake_run(
        args: list[str], **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        """Allow the override probe to succeed."""

        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)

    assert _resolve_python_path(tmp_path / "soffice") == override_path


def test_resolve_python_path_checks_resolved_soffice_dir(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that bundled Python lookup follows the real soffice program directory."""

    real_program_dir = tmp_path / "real-program"
    real_program_dir.mkdir()
    real_soffice = real_program_dir / "soffice"
    real_soffice.write_text("", encoding="utf-8")
    real_python = real_program_dir / "python.bin"
    real_python.write_text("", encoding="utf-8")

    soffice_path = tmp_path / "bin" / "soffice"
    soffice_path.parent.mkdir(parents=True)
    soffice_path.write_text("", encoding="utf-8")

    original_resolve = Path.resolve

    def _fake_resolve(self: Path, *, strict: bool = False) -> Path:
        if self == soffice_path:
            return real_soffice
        return original_resolve(self, strict=strict)

    monkeypatch.setattr(Path, "resolve", _fake_resolve)
    monkeypatch.setattr(
        "exstruct.core.libreoffice._python_supports_libreoffice_bridge",
        lambda path: path == real_python,
    )

    assert _resolve_python_path(soffice_path) == real_python


def test_resolve_python_path_falls_back_to_system_python(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that Debian/Ubuntu-style system Python fallback is supported."""

    soffice_path = tmp_path / "soffice"
    soffice_path.write_text("", encoding="utf-8")
    venv_python = tmp_path / "venv-python"
    venv_python.write_text("", encoding="utf-8")
    system_python = tmp_path / "system-python"
    system_python.write_text("", encoding="utf-8")

    monkeypatch.setattr(
        "exstruct.core.libreoffice._system_python_candidates",
        lambda: (venv_python, system_python),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._python_supports_libreoffice_bridge",
        lambda path: path == system_python,
    )

    assert _resolve_python_path(soffice_path) == system_python


def test_resolve_python_path_returns_none_without_compatible_python(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that auto-detection rejects Python candidates without UNO support."""

    soffice_path = tmp_path / "soffice"
    soffice_path.write_text("", encoding="utf-8")
    candidate = tmp_path / "python3"
    candidate.write_text("", encoding="utf-8")

    monkeypatch.setattr(
        "exstruct.core.libreoffice._system_python_candidates",
        lambda: (candidate,),
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._python_supports_libreoffice_bridge",
        lambda _path: False,
    )

    assert _resolve_python_path(soffice_path) is None


def test_python_supports_libreoffice_bridge_uses_probe_command(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that bridge compatibility is checked via the bundled probe script."""

    python_path = tmp_path / "python3"
    python_path.write_text("", encoding="utf-8")
    calls: list[list[str]] = []

    def _fake_run(
        args: list[str], **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        """Capture subprocess arguments for the bridge probe."""

        calls.append(args)
        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)

    assert _python_supports_libreoffice_bridge(python_path) is True
    assert len(calls) == 1
    assert calls[0][0] == str(python_path)
    assert calls[0][1].endswith("_libreoffice_bridge.py")
    assert calls[0][2] == "--probe"


def test_python_supports_libreoffice_bridge_filters_env(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that bridge probes pass a bounded inherited environment."""

    python_path = tmp_path / "python3"
    python_path.write_text("", encoding="utf-8")
    captured_env: dict[str, str] = {}

    def _fake_run(
        args: list[str], **kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        """Capture subprocess environment for the bridge probe."""

        _ = args
        env = kwargs.get("env")
        if isinstance(env, dict):
            captured_env.update(env)
        return subprocess.CompletedProcess(
            args=args, returncode=0, stdout="", stderr=""
        )

    monkeypatch.setenv("UNSAFE_TEST_ENV", "secret")
    monkeypatch.setenv("PATH", "test-path")
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)

    assert _python_supports_libreoffice_bridge(python_path) is True
    assert captured_env["PYTHONIOENCODING"] == "utf-8"
    assert captured_env["PATH"] == "test-path"
    assert "UNSAFE_TEST_ENV" not in captured_env


def test_resolve_python_path_rejects_system_python_when_bridge_probe_fails(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that system Python fallback rejects bridge-incompatible runtimes."""

    soffice_path = tmp_path / "soffice"
    soffice_path.write_text("", encoding="utf-8")
    system_python = tmp_path / "system-python"
    system_python.write_text("", encoding="utf-8")

    def _fake_run(
        *_args: object, **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        """Simulate a Python that imports UNO but cannot parse the bundled bridge."""

        raise subprocess.CalledProcessError(
            returncode=1,
            cmd=str(system_python),
            stderr="SyntaxError: invalid syntax",
        )

    monkeypatch.setattr(
        "exstruct.core.libreoffice._system_python_candidates",
        lambda: (system_python,),
    )
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)

    assert _resolve_python_path(soffice_path) is None


def test_resolve_python_path_rejects_incompatible_override_early(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that an explicit override fails fast when the bridge probe fails."""

    soffice_path = tmp_path / "soffice"
    soffice_path.write_text("", encoding="utf-8")
    override_path = tmp_path / "custom-python"
    override_path.write_text("", encoding="utf-8")
    monkeypatch.setenv("EXSTRUCT_LIBREOFFICE_PYTHON_PATH", str(override_path))

    def _fake_run(
        *_args: object, **_kwargs: object
    ) -> subprocess.CompletedProcess[str]:
        """Simulate a bridge parse failure for the explicit override."""

        raise subprocess.CalledProcessError(
            returncode=1,
            cmd=str(override_path),
            stderr="SyntaxError: invalid syntax",
        )

    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)

    with pytest.raises(
        LibreOfficeUnavailableError,
        match="EXSTRUCT_LIBREOFFICE_PYTHON_PATH.*incompatible",
    ):
        _resolve_python_path(soffice_path)


def test_libreoffice_session_exit_cleans_profile_even_if_kill_wait_times_out(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that LibreOffice session exit cleans profile even if kill wait times out."""

    cleaned_paths: list[Path] = []

    def _fake_cleanup(path: Path) -> None:
        """Provide a fake cleanup implementation for this test."""

        cleaned_paths.append(path)

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=tmp_path / "soffice.exe",
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )
    profile_dir = tmp_path / "lo-profile"
    process = _FakeLibreOfficeProcess(wait_timeouts=2)
    session._temp_profile_dir = profile_dir
    session._soffice_process = cast(subprocess.Popen[str], process)
    session._accept_port = 12345
    monkeypatch.setattr("exstruct.core.libreoffice._cleanup_profile_dir", _fake_cleanup)

    session.__exit__(None, None, None)

    assert process.terminate_called is True
    assert process.kill_called is True
    assert process.wait_calls == 2
    assert cleaned_paths == [profile_dir]
    assert session._soffice_process is None
    assert session._temp_profile_dir is None
    assert session._accept_port is None


def test_libreoffice_session_extractors_cache_bridge_payloads_and_parse_results(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify bridge-backed extractors cache per kind and coerce payload details."""

    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    workbook_path = tmp_path / "book.xlsx"
    for path in (soffice_path, python_path, workbook_path):
        path.write_text("", encoding="utf-8")
    bridge_calls: list[list[str]] = []

    def _fake_run_trusted_subprocess(
        args: list[str], *, timeout_sec: float, env: dict[str, str] | None = None
    ) -> subprocess.CompletedProcess[str]:
        bridge_calls.append(list(args))
        assert timeout_sec == 2.0
        assert env is not None
        assert env["PYTHONIOENCODING"] == "utf-8"
        kind = args[args.index("--kind") + 1]
        draw_items: list[object] = [
            {
                "name": "Flow",
                "shape_type": "com.sun.star.drawing.ConnectorShape",
                "text": "step",
                "left": 10.2,
                "top": 20,
                "width": 30.7,
                "height": 40.1,
                "rotation": 45,
                "is_connector": 1,
                "start_shape_name": "Start",
                "end_shape_name": "End",
            },
            "skip me",
            {
                "name": "",
                "shape_type": 123,
                "text": None,
                "left": 2.4,
                "top": 3.6,
                "width": 4.4,
                "height": 5.6,
                "rotation": 90,
                "is_connector": True,
                "start_shape_name": "Origin",
                "end_shape_name": 999,
            },
        ]
        chart_items: list[object] = [
            {
                "name": "Chart 1",
                "persist_name": "Object 1",
                "left": 100.2,
                "top": 200.2,
                "width": 300.8,
                "height": 400.1,
            },
            {"name": "", "persist_name": "ignored"},
            123,
            {"name": "Chart 2", "persist_name": 1234},
        ]
        payload: object = (
            {"sheets": {"Sheet1": draw_items, "Broken": "not-a-list"}}
            if kind == "draw-page"
            else {"sheets": {"Sheet1": chart_items, "Broken": "not-a-list"}}
        )
        return subprocess.CompletedProcess(
            args=args,
            returncode=0,
            stdout=json.dumps(payload),
            stderr="",
        )

    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._run_trusted_subprocess",
        _fake_run_trusted_subprocess,
    )

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=2.0,
            profile_root=None,
        )
    )
    session._accept_port = 42001

    workbook_token = session.load_workbook(workbook_path)
    assert workbook_token == {"file_path": str(workbook_path.resolve())}
    session.close_workbook(workbook_token)

    draw_page_shapes = session.extract_draw_page_shapes(workbook_path)
    chart_geometries = session.extract_chart_geometries(workbook_path)

    assert session.extract_draw_page_shapes(workbook_path) == draw_page_shapes
    assert session.extract_chart_geometries(workbook_path) == chart_geometries
    assert [call[call.index("--kind") + 1] for call in bridge_calls] == [
        "draw-page",
        "charts",
    ]

    assert draw_page_shapes == {
        "Sheet1": [
            LibreOfficeDrawPageShape(
                name="Flow",
                shape_type="com.sun.star.drawing.ConnectorShape",
                text="step",
                left=10,
                top=20,
                width=31,
                height=40,
                rotation=45.0,
                is_connector=True,
                start_shape_name="Start",
                end_shape_name="End",
            ),
            LibreOfficeDrawPageShape(
                name="Shape 3",
                shape_type=None,
                text="",
                left=2,
                top=4,
                width=4,
                height=6,
                rotation=90.0,
                is_connector=True,
                start_shape_name="Origin",
                end_shape_name=None,
            ),
        ]
    }
    assert chart_geometries == {
        "Sheet1": [
            LibreOfficeChartGeometry(
                name="Chart 1",
                persist_name="Object 1",
                left=100,
                top=200,
                width=301,
                height=400,
            ),
            LibreOfficeChartGeometry(
                name="Chart 2",
                persist_name=None,
                left=None,
                top=None,
                width=None,
                height=None,
            ),
        ]
    }


def test_libreoffice_session_run_bridge_requires_entered_session(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify bridge extraction fails before a session has been entered."""

    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    workbook_path = tmp_path / "book.xlsx"
    for path in (soffice_path, python_path, workbook_path):
        path.write_text("", encoding="utf-8")
    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )
    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )

    with pytest.raises(RuntimeError, match="must be entered before extraction"):
        session._run_bridge(workbook_path, kind="charts")


def test_libreoffice_session_run_bridge_surfaces_subprocess_failures(
    tmp_path: Path,
    monkeypatch: MonkeyPatch,
) -> None:
    """Verify bridge extraction turns subprocess failures into actionable errors."""

    cases: tuple[tuple[Exception, type[BaseException], str], ...] = (
        (
            FileNotFoundError("missing python"),
            LibreOfficeUnavailableError,
            "compatible Python runtime could not be executed",
        ),
        (
            subprocess.TimeoutExpired(cmd="bridge", timeout=2.0),
            LibreOfficeUnavailableError,
            "UNO bridge extraction timed out",
        ),
        (
            subprocess.CalledProcessError(
                returncode=1,
                cmd="bridge",
                stderr="bridge crashed",
            ),
            RuntimeError,
            "LibreOffice UNO bridge extraction failed: bridge crashed",
        ),
    )
    for raised, expected_exception, expected_message in cases:
        soffice_path = tmp_path / "soffice.exe"
        python_path = tmp_path / "python.exe"
        workbook_path = tmp_path / "book.xlsx"
        for path in (soffice_path, python_path, workbook_path):
            path.write_text("", encoding="utf-8")

        def _fake_run_trusted_subprocess(
            _args: list[str],
            *,
            timeout_sec: float,
            env: dict[str, str] | None = None,
            raised: Exception = raised,
        ) -> subprocess.CompletedProcess[str]:
            _ = timeout_sec
            _ = env
            raise raised

        monkeypatch.setattr(
            "exstruct.core.libreoffice._resolve_python_path",
            lambda _path, python_path=python_path: python_path,
        )
        monkeypatch.setattr(
            "exstruct.core.libreoffice._run_trusted_subprocess",
            _fake_run_trusted_subprocess,
        )

        session = LibreOfficeSession(
            LibreOfficeSessionConfig(
                soffice_path=soffice_path,
                startup_timeout_sec=1.0,
                exec_timeout_sec=2.0,
                profile_root=None,
            )
        )
        session._accept_port = 42001

        with pytest.raises(expected_exception, match=expected_message):
            session._run_bridge(workbook_path, kind="charts")


def test_libreoffice_session_run_bridge_rejects_invalid_json(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify bridge extraction rejects malformed JSON payloads."""

    soffice_path = tmp_path / "soffice.exe"
    python_path = tmp_path / "python.exe"
    workbook_path = tmp_path / "book.xlsx"
    for path in (soffice_path, python_path, workbook_path):
        path.write_text("", encoding="utf-8")

    def _fake_run_trusted_subprocess(
        args: list[str], *, timeout_sec: float, env: dict[str, str] | None = None
    ) -> subprocess.CompletedProcess[str]:
        _ = timeout_sec
        _ = env
        return subprocess.CompletedProcess(
            args=args,
            returncode=0,
            stdout="{not-json",
            stderr="",
        )

    monkeypatch.setattr(
        "exstruct.core.libreoffice._resolve_python_path",
        lambda _path: python_path,
    )
    monkeypatch.setattr(
        "exstruct.core.libreoffice._run_trusted_subprocess",
        _fake_run_trusted_subprocess,
    )

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=2.0,
            profile_root=None,
        )
    )
    session._accept_port = 42001

    with pytest.raises(
        RuntimeError,
        match="LibreOffice UNO bridge extraction returned invalid JSON",
    ):
        session._run_bridge(workbook_path, kind="draw-page")


def test_libreoffice_payload_parsers_reject_invalid_top_level_data() -> None:
    """Verify payload parsers reject malformed top-level bridge payloads."""

    cases: tuple[tuple[Callable[[object], object], object, str], ...] = (
        (
            _parse_chart_payload,
            [],
            "LibreOffice UNO chart extraction returned a non-object payload",
        ),
        (
            _parse_chart_payload,
            {"sheets": []},
            "LibreOffice UNO chart extraction payload is missing sheets",
        ),
        (
            _parse_draw_page_payload,
            [],
            "LibreOffice UNO draw-page extraction returned a non-object payload",
        ),
        (
            _parse_draw_page_payload,
            {"sheets": []},
            "LibreOffice UNO draw-page extraction payload is missing sheets",
        ),
    )

    for parser, payload, expected_message in cases:
        with pytest.raises(RuntimeError, match=expected_message):
            parser(payload)
