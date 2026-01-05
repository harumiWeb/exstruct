from __future__ import annotations

from types import SimpleNamespace

from _pytest.monkeypatch import MonkeyPatch

from exstruct.core.cells import SheetColorsMap, WorkbookColorsMap
from exstruct.core.pipeline import collect_sheet_raw_data
from exstruct.models import CellRow, Chart, ChartSeries, PrintArea, Shape


def _make_chart() -> Chart:
    """Build a minimal chart instance for tests.

    Returns:
        Chart instance with the minimum required fields.
    """
    return Chart(
        name="Chart1",
        chart_type="Column",
        title=None,
        y_axis_title="Y",
        y_axis_range=[],
        w=None,
        h=None,
        series=[ChartSeries(name="Series1")],
        l=0,
        t=0,
    )


def test_collect_sheet_raw_data_includes_extracted_fields(
    monkeypatch: MonkeyPatch,
) -> None:
    """Collect raw sheet data with shapes/charts/tables and colors_map.

    Args:
        monkeypatch: Pytest monkeypatch fixture.
    """
    sheet = SimpleNamespace(name="Sheet1")
    workbook = SimpleNamespace(sheets={"Sheet1": sheet})

    monkeypatch.setattr(
        "exstruct.core.pipeline.detect_tables",
        lambda _sheet: ["A1:B2"],
    )

    colors_map = WorkbookColorsMap(
        sheets={
            "Sheet1": SheetColorsMap(
                sheet_name="Sheet1", colors_map={"#FFFFFF": [(1, 0)]}
            )
        }
    )
    result = collect_sheet_raw_data(
        cell_data={"Sheet1": [CellRow(r=1, c={"0": "A"}, links=None)]},
        shape_data={"Sheet1": [Shape(text="S", l=0, t=0)]},
        chart_data={"Sheet1": [_make_chart()]},
        merged_cell_data={"Sheet1": []},
        workbook=workbook,
        mode="standard",
        print_area_data={"Sheet1": [PrintArea(r1=1, c1=0, r2=1, c2=0)]},
        auto_page_break_data={"Sheet1": [PrintArea(r1=1, c1=0, r2=1, c2=0)]},
        colors_map_data=colors_map,
    )

    raw = result["Sheet1"]
    assert raw.rows
    assert raw.shapes
    assert raw.charts
    assert raw.table_candidates == ["A1:B2"]
    assert raw.print_areas
    assert raw.auto_print_areas
    assert raw.colors_map == {"#FFFFFF": [(1, 0)]}


def test_collect_sheet_raw_data_skips_charts_in_light_mode(
    monkeypatch: MonkeyPatch,
) -> None:
    """Skip chart extraction in light mode.

    Args:
        monkeypatch: Pytest monkeypatch fixture.
    """
    sheet = SimpleNamespace(name="Sheet1")
    workbook = SimpleNamespace(sheets={"Sheet1": sheet})

    monkeypatch.setattr("exstruct.core.pipeline.detect_tables", lambda _sheet: [])

    result = collect_sheet_raw_data(
        cell_data={"Sheet1": []},
        shape_data={"Sheet1": []},
        chart_data={"Sheet1": []},
        merged_cell_data={"Sheet1": []},
        workbook=workbook,
        mode="light",
        print_area_data=None,
        auto_page_break_data=None,
        colors_map_data=None,
    )

    assert result["Sheet1"].charts == []
