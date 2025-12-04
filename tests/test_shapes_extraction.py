from pathlib import Path

import pytest
import xlwings as xw

from exstruct.core.integrate import extract_workbook


def _ensure_excel() -> None:
    try:
        app = xw.App(add_book=False, visible=False)
        app.quit()
    except Exception:
        pytest.skip("Excel COM is unavailable; skipping shapes extraction tests.")


def _make_workbook_with_shapes(path: Path) -> None:
    app = xw.App(add_book=False, visible=False)
    try:
        wb = app.books.add()
        sht = wb.sheets[0]
        sht.name = "Sheet1"

        rect = sht.api.Shapes.AddShape(1, 50, 50, 120, 60)  # msoShapeRectangle
        rect.TextFrame2.TextRange.Text = "rect"

        silent_rect = sht.api.Shapes.AddShape(5, 300, 50, 80, 40)  # msoShapeOval (no text)

        line = sht.api.Shapes.AddLine(10, 10, 110, 10)
        line.Line.EndArrowheadStyle = 3  # msoArrowheadTriangle

        outer = sht.api.Shapes.AddShape(1, 200, 200, 150, 100)
        inner = sht.api.Shapes.AddShape(1, 230, 230, 80, 40)
        inner.TextFrame2.TextRange.Text = "inner"
        sht.api.Shapes.Range([outer.Name, inner.Name]).Group()

        wb.save(str(path))
        wb.close()
    finally:
        try:
            app.quit()
        except Exception:
            pass


def test_図形の種別とテキストが抽出される(tmp_path: Path) -> None:
    _ensure_excel()
    path = tmp_path / "shapes.xlsx"
    _make_workbook_with_shapes(path)

    wb_data = extract_workbook(path)
    shapes = wb_data.sheets["Sheet1"].shapes

    rect = next(s for s in shapes if s.text == "rect")
    assert "AutoShape" in (rect.type or "")
    assert rect.l >= 0 and rect.t >= 0

    inner = next(s for s in shapes if s.text == "inner")
    assert "Group" not in (inner.type or "")  # flattened child
    assert not any((s.type or "") == "Group" for s in shapes)
    assert not any(
        (s.text == "" or s.text is None)
        and (s.type or "").startswith("AutoShape")
        for s in shapes
    )


def test_線図形の方向と矢印情報が抽出される(tmp_path: Path) -> None:
    _ensure_excel()
    path = tmp_path / "lines.xlsx"
    _make_workbook_with_shapes(path)

    wb_data = extract_workbook(path)
    shapes = wb_data.sheets["Sheet1"].shapes

    line = next(s for s in shapes if s.begin_arrow_style is not None or s.end_arrow_style is not None)
    assert line.direction == "E"
