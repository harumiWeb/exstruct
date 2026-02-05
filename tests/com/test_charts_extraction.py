from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

import pytest
import xlwings as xw

from exstruct.core.integrate import extract_workbook

pytestmark = pytest.mark.com


@contextmanager
def _excel_app() -> xw.App:
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


def _make_workbook_with_chart(path: Path) -> None:
    with _excel_app() as app:
        wb = app.books.add()
        try:
            sht = wb.sheets[0]
            sht.name = "Sheet1"
            sht.range("A1").value = ["Month", "Sales"]
            sht.range("A2").value = [
                ["Jan", 100],
                ["Feb", 120],
                ["Mar", 150],
                ["Apr", 180],
            ]

            chart = sht.charts.add(left=200, top=50, width=300, height=200)
            chart.chart_type = "column_clustered"
            chart.set_source_data(sht.range("A1:B5"))
            chart_name = chart.name
            chart_com = sht.api.ChartObjects(chart_name).Chart
            chart_com.HasTitle = True
            chart_com.ChartTitle.Text = "Sales Chart"
            y_axis = chart_com.Axes(2, 1)
            y_axis.HasTitle = True
            y_axis.AxisTitle.Text = "Amount"
            y_axis.MinimumScale = 0
            y_axis.MaximumScale = 200

            wb.save(str(path))
        finally:
            wb.close()


def test_chart_basic_metadata(tmp_path: Path) -> None:
    path = tmp_path / "chart.xlsx"
    _make_workbook_with_chart(path)

    wb_data = extract_workbook(path)
    charts = wb_data.sheets["Sheet1"].charts
    assert len(charts) == 1

    ch = charts[0]
    assert ch.chart_type.lower().startswith("column")
    assert ch.title == "Sales Chart"
    assert ch.y_axis_title == "Amount"
    assert ch.y_axis_range == [0, 200]
    assert ch.error is None


def test_chart_series_ranges(tmp_path: Path) -> None:
    path = tmp_path / "chart_series.xlsx"
    _make_workbook_with_chart(path)

    wb_data = extract_workbook(path)
    ch = wb_data.sheets["Sheet1"].charts[0]
    assert len(ch.series) == 1
    s = ch.series[0]
    assert s.name_range is None or s.name_range.endswith("Sheet1!$B$1")
    assert s.x_range is None or s.x_range.endswith("Sheet1!$A$2:$A$5")
    assert s.y_range is None or s.y_range.endswith("Sheet1!$B$2:$B$5")


def test_chart_verbose_size(tmp_path: Path) -> None:
    path = tmp_path / "chart_verbose.xlsx"
    _make_workbook_with_chart(path)

    wb_data = extract_workbook(path, mode="verbose")
    ch = wb_data.sheets["Sheet1"].charts[0]
    assert ch.w is not None and ch.w > 0
    assert ch.h is not None and ch.h > 0
