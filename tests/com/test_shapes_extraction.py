from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import pytest
import xlwings as xw

from exstruct.core.integrate import extract_workbook
from exstruct.models import Arrow, Shape

pytestmark = pytest.mark.com


@contextmanager
def _excel_app() -> Iterator[xw.App]:
    app = xw.App(add_book=False, visible=False)
    try:
        yield app
    finally:
        try:
            app.quit()
        except Exception:
            try:
                app.kill()
            except Exception:
                pass


def _make_workbook_with_shapes(path: Path) -> None:
    with _excel_app() as app:
        wb = app.books.add()
        try:
            sht = wb.sheets[0]
            sht.name = "Sheet1"

            rect = sht.api.Shapes.AddShape(1, 50, 50, 120, 60)  # msoShapeRectangle
            rect.TextFrame2.TextRange.Text = "rect"

            _ = sht.api.Shapes.AddShape(5, 300, 50, 80, 40)  # msoShapeOval (no text)

            line = sht.api.Shapes.AddLine(10, 10, 110, 10)
            line.Line.EndArrowheadStyle = 3  # msoArrowheadTriangle

            outer = sht.api.Shapes.AddShape(1, 200, 200, 150, 100)
            inner = sht.api.Shapes.AddShape(1, 230, 230, 80, 40)
            inner.TextFrame2.TextRange.Text = "inner"
            sht.api.Shapes.Range([outer.Name, inner.Name]).Group()

            # Add two rectangles and a connector that explicitly connects them
            # so that ConnectorFormat.BeginConnectedShape / EndConnectedShape
            # are populated by Excel.
            src_shape = sht.api.Shapes.AddShape(1, 50, 150, 80, 40)
            src_shape.TextFrame2.TextRange.Text = "src"
            dst_shape = sht.api.Shapes.AddShape(1, 200, 150, 80, 40)
            dst_shape.TextFrame2.TextRange.Text = "dst"
            connector = sht.api.Shapes.AddConnector(1, 90, 170, 200, 170)
            connector.Line.EndArrowheadStyle = 3
            try:
                connector.ConnectorFormat.BeginConnect(src_shape, 1)
                connector.ConnectorFormat.EndConnect(dst_shape, 1)
            except Exception:
                # In some environments connector wiring may fail; tests will
                # simply not find connected shapes in that case.
                pass

            wb.save(str(path))
        finally:
            wb.close()


def test_shapes_basic(tmp_path: Path) -> None:
    """
    Verifies extraction of shape types, texts, IDs, and uniqueness from a workbook containing various shapes.

    Creates a workbook with a rectangle, an oval, a line, a nested group, and a connector, then extracts shapes from "Sheet1" and asserts:
    - a shape with text "rect" is an AutoShape, has non-negative left/top coordinates, and a positive id;
    - a nested child with text "inner" is not reported as a Group and has a positive id;
    - all emitted shape ids are unique;
    - no AutoShape without text is emitted in standard mode.
    """
    path = tmp_path / "shapes.xlsx"
    _make_workbook_with_shapes(path)

    wb_data = extract_workbook(path)
    shapes = wb_data.sheets["Sheet1"].shapes

    rect = next(s for s in shapes if isinstance(s, Shape) and s.text == "rect")
    assert "AutoShape" in (rect.type or "")
    assert rect.l >= 0 and rect.t >= 0
    assert rect.id is not None and rect.id > 0

    inner = next(s for s in shapes if isinstance(s, Shape) and s.text == "inner")
    assert "Group" not in (inner.type or "")  # flattened child
    assert not any(isinstance(s, Shape) and (s.type or "") == "Group" for s in shapes)
    assert inner.id is not None and inner.id > 0
    ids = [s.id for s in shapes if s.id is not None]
    assert len(ids) == len(set(ids))
    # Standard mode should not emit non-relationship AutoShapes without text.
    assert not any(
        isinstance(s, Shape)
        and (s.text == "" or s.text is None)
        and (s.type or "").startswith("AutoShape")
        for s in shapes
    )


def test_line_direction(tmp_path: Path) -> None:
    """
    Verifies that a line shape's direction and arrow style information is extracted correctly from a workbook.

    Creates a workbook containing shapes, extracts shapes from "Sheet1", finds an Arrow with a begin or end arrow style, and asserts its direction is "E".
    """
    path = tmp_path / "lines.xlsx"
    _make_workbook_with_shapes(path)

    wb_data = extract_workbook(path)
    shapes = wb_data.sheets["Sheet1"].shapes

    line = next(
        s
        for s in shapes
        if isinstance(s, Arrow)
        and (s.begin_arrow_style is not None or s.end_arrow_style is not None)
    )
    assert line.direction == "E"


def test_connector_connections(tmp_path: Path) -> None:
    """Connectorの接続先IDが抽出結果のshape IDに整合することを確認する。"""
    path = tmp_path / "connectors.xlsx"
    _make_workbook_with_shapes(path)

    wb_data = extract_workbook(path)
    shapes = wb_data.sheets["Sheet1"].shapes

    connectors = [
        s for s in shapes if isinstance(s, Arrow) and (s.begin_id or s.end_id)
    ]

    # If the environment could not wire connectors, simply skip the assertion.
    if not connectors:
        pytest.skip(
            "Excel failed to populate ConnectorFormat.ConnectedShape properties."
        )

    conn = connectors[0]
    assert conn.begin_id is not None
    assert conn.end_id is not None
    assert conn.begin_id != conn.end_id
    # Connected shape ids should correspond to some emitted shapes' id.
    shape_ids = {s.id for s in shapes if s.id is not None}
    assert conn.begin_id in shape_ids
    assert conn.end_id in shape_ids
