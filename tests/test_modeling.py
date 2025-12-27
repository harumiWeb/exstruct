from exstruct.core.modeling import SheetRawData, WorkbookRawData, build_workbook_data
from exstruct.models import CellRow, Chart, ChartSeries, PrintArea, Shape


def test_build_workbook_data_from_raw() -> None:
    """Build WorkbookData from raw containers."""
    raw_sheet = SheetRawData(
        rows=[CellRow(r=1, c={"0": "A"}, links=None)],
        shapes=[Shape(text="S", l=0, t=0)],
        charts=[
            Chart(
                name="C1",
                chart_type="Column",
                title=None,
                y_axis_title="Y",
                y_axis_range=[],
                w=None,
                h=None,
                series=[ChartSeries(name="S1")],
                l=0,
                t=0,
            )
        ],
        table_candidates=["A1:A1"],
        print_areas=[PrintArea(r1=1, c1=0, r2=1, c2=0)],
        auto_print_areas=[],
        colors_map={"#FFFFFF": [(1, 0)]},
    )
    raw_workbook = WorkbookRawData(book_name="book.xlsx", sheets={"Sheet1": raw_sheet})

    wb = build_workbook_data(raw_workbook)

    assert wb.book_name == "book.xlsx"
    assert "Sheet1" in wb.sheets
    sheet = wb.sheets["Sheet1"]
    assert sheet.rows
    assert sheet.shapes
    assert sheet.charts
    assert sheet.print_areas
