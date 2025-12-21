from pydantic import ValidationError
import pytest

from exstruct.models import (
    CellRow,
    Chart,
    ChartSeries,
    Shape,
    SheetData,
    WorkbookData,
)


def test_モデルのデフォルトとオプション値() -> None:
    shape = Shape(id=1, text="t", l=1, t=2, w=None, h=None)
    assert shape.rotation is None
    assert shape.direction is None

    cell = CellRow(r=0, c={"0": "v"})
    assert cell.c["0"] == "v"

    chart_series = ChartSeries(name="s")
    assert chart_series.name_range is None
    assert chart_series.x_range is None
    assert chart_series.y_range is None

    chart = Chart(
        name="c",
        chart_type="Line",
        title=None,
        y_axis_title="",
        y_axis_range=[],
        series=[],
        l=0,
        t=0,
    )
    assert chart.error is None

    sheet = SheetData()
    assert sheet.rows == []
    assert sheet.shapes == []
    assert sheet.charts == []
    assert sheet.table_candidates == []

    wb = WorkbookData(book_name="b.xlsx", sheets={"S": sheet})
    assert wb.sheets["S"].rows == []


def test_directionのリテラル検証() -> None:
    with pytest.raises(ValidationError):
        Shape(id=1, text="bad", l=0, t=0, w=None, h=None, direction="X")


def test_cellrowの数値正規化() -> None:
    cell = CellRow(r=1, c={"0": 123, "1": 1.5, "2": "text"})
    assert isinstance(cell.c["0"], int)
    assert isinstance(cell.c["1"], float)
    assert cell.c["2"] == "text"
