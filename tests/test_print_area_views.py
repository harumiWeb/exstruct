import json
from pathlib import Path

from exstruct.io import save_print_area_views
from exstruct.models import CellRow, Chart, PrintArea, Shape, SheetData, WorkbookData


def _workbook_with_print_area() -> WorkbookData:
    shape_inside = Shape(id=1, text="inside", l=10, t=5, w=20, h=10, type="Rect")
    shape_outside = Shape(id=2, text="outside", l=200, t=200, w=30, h=30, type="Rect")
    chart_inside = Chart(
        name="c1",
        chart_type="Line",
        title=None,
        y_axis_title="",
        y_axis_range=[],
        w=50,
        h=30,
        series=[],
        l=0,
        t=0,
        error=None,
    )
    chart_outside = Chart(
        name="c2",
        chart_type="Line",
        title=None,
        y_axis_title="",
        y_axis_range=[],
        w=None,
        h=None,
        series=[],
        l=300,
        t=300,
        error=None,
    )
    sheet = SheetData(
        rows=[
            CellRow(r=1, c={"0": "A", "2": "skip"}, links={"2": "http://example.com"}),
            CellRow(r=2, c={"1": "B"}),
            CellRow(r=3, c={"1": "C"}),
        ],
        shapes=[shape_inside, shape_outside],
        charts=[chart_inside, chart_outside],
        table_candidates=["A1:B2", "C1:C1"],
        print_areas=[PrintArea(r1=1, c1=0, r2=2, c2=1)],
    )
    return WorkbookData(book_name="book.xlsx", sheets={"Sheet1": sheet})


def test_save_print_area_views_filters_rows_and_tables(tmp_path: Path) -> None:
    wb = _workbook_with_print_area()
    written = save_print_area_views(wb, tmp_path, fmt="json", pretty=True)
    assert len(written) == 1
    path = next(iter(written.values()))
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["sheet_name"] == "Sheet1"
    assert data["area"] == {"r1": 1, "c1": 0, "r2": 2, "c2": 1}
    # Only cells within the area remain; out-of-range cells and links are dropped.
    assert data["rows"] == [{"r": 1, "c": {"0": "A"}}, {"r": 2, "c": {"1": "B"}}]
    # Only table candidates fully contained in the print area remain.
    assert data["table_candidates"] == ["A1:B2"]
    # Shapes/Charts filtered by overlap; outside or size-less charts are dropped.
    assert len(data["shapes"]) == 1 and data["shapes"][0]["text"] == "inside"
    assert len(data["charts"]) == 1 and data["charts"][0]["name"] == "c1"


def test_save_print_area_views_normalizes_when_requested(tmp_path: Path) -> None:
    wb = _workbook_with_print_area()
    # Shift the print area to demonstrate normalization.
    wb.sheets["Sheet1"].print_areas[0] = PrintArea(r1=2, c1=1, r2=3, c2=2)
    written = save_print_area_views(
        wb, tmp_path, fmt="json", pretty=False, normalize=True
    )
    path = next(iter(written.values()))
    data = json.loads(path.read_text(encoding="utf-8"))
    # Rows inside the area are re-based to start at 0, columns start at 0.
    assert data["rows"] == [{"r": 0, "c": {"0": "B"}}, {"r": 1, "c": {"0": "C"}}]
    # Shapes/charts may be dropped when they don't overlap; ensure no unexpected entries.
    assert "shapes" not in data or len(data["shapes"]) <= 1


def test_save_print_area_views_no_print_areas_returns_empty(tmp_path: Path) -> None:
    wb = _workbook_with_print_area()
    wb.sheets["Sheet1"].print_areas = []
    written = save_print_area_views(wb, tmp_path, fmt="json")
    assert written == {}
