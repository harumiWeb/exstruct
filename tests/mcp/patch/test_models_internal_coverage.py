from __future__ import annotations

from collections.abc import Callable
from typing import Any, cast
from unittest.mock import MagicMock

from pydantic import ValidationError
import pytest

from exstruct.mcp.patch import internal, models

PatchOpFactory = Callable[..., object]


@pytest.mark.parametrize(
    ("op_factory", "request_factory", "make_factory"),
    [
        (models.PatchOp, models.PatchRequest, models.MakeRequest),
        (internal.PatchOp, internal.PatchRequest, internal.MakeRequest),
    ],
    ids=["models", "internal"],
)  # type: ignore[misc]
@pytest.mark.parametrize(
    ("payload", "message"),
    [
        (
            {"op": "set_value", "sheet": "Sheet1", "cell": "1A", "value": "x"},
            "Invalid cell reference",
        ),
        (
            {"op": "set_dimensions", "sheet": "Sheet1"},
            "set_dimensions requires rows and/or columns",
        ),
        (
            {"op": "set_dimensions", "sheet": "Sheet1", "rows": [1]},
            "set_dimensions requires row_height when rows is provided",
        ),
        (
            {
                "op": "set_alignment",
                "sheet": "Sheet1",
                "cell": "A1",
            },
            "set_alignment requires at least one of",
        ),
        (
            {
                "op": "set_style",
                "sheet": "Sheet1",
                "cell": "A1",
                "font_size": 0,
            },
            "set_style font_size must be > 0",
        ),
        (
            {
                "op": "draw_grid_border",
                "sheet": "Sheet1",
                "base_cell": "A1",
                "row_count": 0,
                "col_count": 1,
            },
            "draw_grid_border requires row_count >= 1 and col_count >= 1",
        ),
        (
            {
                "op": "set_formula",
                "sheet": "Sheet1",
                "cell": "A1",
                "formula": "SUM(1,1)",
            },
            "set_formula requires formula starting with '='",
        ),
        (
            {
                "op": "set_fill_color",
                "sheet": "Sheet1",
                "cell": "A1",
                "fill_color": "red",
            },
            "Invalid fill_color format",
        ),
        (
            {
                "op": "set_font_color",
                "sheet": "Sheet1",
                "cell": "A1",
                "color": "#112233",
                "fill_color": "#FFFFFF",
            },
            "set_font_color does not accept fill_color",
        ),
        (
            {
                "op": "auto_fit_columns",
                "sheet": "Sheet1",
                "min_width": 10,
                "max_width": 5,
            },
            "auto_fit_columns requires min_width <= max_width",
        ),
        (
            {
                "op": "set_dimensions",
                "sheet": "Sheet1",
                "columns": [0],
                "column_width": 18,
            },
            "columns numeric values must be positive",
        ),
    ],
)  # type: ignore[misc]
def test_patch_op_validation_errors(
    op_factory: PatchOpFactory,
    request_factory: Callable[..., object],
    make_factory: Callable[..., object],
    payload: dict[str, Any],
    message: str,
) -> None:
    with pytest.raises(ValidationError, match=message):
        op_factory(**payload)

    # Keep fixtures referenced so parametrization stays aligned across modules.
    assert request_factory is not None
    assert make_factory is not None


@pytest.mark.parametrize(
    ("op_factory", "request_factory", "make_factory"),
    [
        (models.PatchOp, models.PatchRequest, models.MakeRequest),
        (internal.PatchOp, internal.PatchRequest, internal.MakeRequest),
    ],
    ids=["models", "internal"],
)  # type: ignore[misc]
def test_backend_com_rejects_dry_run_and_restore_design_snapshot(
    op_factory: PatchOpFactory,
    request_factory: Callable[..., object],
    make_factory: Callable[..., object],
) -> None:
    with pytest.raises(ValidationError, match="backend='com' does not support"):
        request_factory(
            xlsx_path="book.xlsx",
            ops=[op_factory(op="add_sheet", sheet="Data")],
            dry_run=True,
            backend="com",
        )

    with pytest.raises(ValidationError, match="backend='com' does not support"):
        make_factory(
            out_path="book.xlsx",
            ops=[op_factory(op="add_sheet", sheet="Data")],
            dry_run=True,
            backend="com",
        )

    design_snapshot: dict[str, list[object]] = {
        "borders": [],
        "fonts": [],
        "fills": [],
        "alignments": [],
        "row_dimensions": [],
        "column_dimensions": [],
    }
    with pytest.raises(
        ValidationError,
        match="backend='com' does not support restore_design_snapshot operation",
    ):
        request_factory(
            xlsx_path="book.xlsx",
            ops=[
                op_factory(
                    op="restore_design_snapshot",
                    sheet="Sheet1",
                    design_snapshot=design_snapshot,
                )
            ],
            backend="com",
        )


def test_internal_xlwings_helpers_error_and_success_paths() -> None:
    class _FakeFont:
        Bold: bool = False
        Size: float = 0.0
        Color: int = 0

    class _FakeInterior:
        Color: int = 0

    class _FakeRangeApi:
        Font: _FakeFont
        Interior: _FakeInterior

        def __init__(self) -> None:
            self.Font = _FakeFont()
            self.Interior = _FakeInterior()

    class _FakeRange:
        value: object | None = None
        formula: str | None = None
        api: _FakeRangeApi

        def __init__(self) -> None:
            self.api = _FakeRangeApi()

    class _FakeSheet:
        name = "Sheet1"
        api = object()

        def __init__(self) -> None:
            self.ranges: dict[str, _FakeRange] = {}

        def range(self, ref: str) -> _FakeRange:
            self.ranges.setdefault(ref, _FakeRange())
            return self.ranges[ref]

    class _FakeSheets:
        def __init__(self, initial: list[_FakeSheet]) -> None:
            self._items = initial

        def __getitem__(self, index: int) -> _FakeSheet:
            return self._items[index]

        def __len__(self) -> int:
            return len(self._items)

        def add(self, name: str, after: _FakeSheet | None = None) -> _FakeSheet:
            del after
            sheet = _FakeSheet()
            sheet.name = name
            self._items.append(sheet)
            return sheet

    class _FakeWorkbook:
        def __init__(self) -> None:
            self.sheets = _FakeSheets([_FakeSheet()])

    workbook = _FakeWorkbook()
    known_sheet = workbook.sheets[0]
    sheet_map = {"Sheet1": known_sheet}

    add_sheet_op = internal.PatchOp(op="add_sheet", sheet="NewSheet")
    diff = internal._apply_xlwings_op(
        cast(internal.XlwingsWorkbookProtocol, workbook),
        cast(dict[str, internal.XlwingsSheetProtocol], sheet_map),
        add_sheet_op,
        0,
        False,
    )
    assert diff.after is not None
    assert diff.after.value == "NewSheet"
    assert "NewSheet" in sheet_map

    missing_sheet_op = internal.PatchOp(
        op="set_value", sheet="Missing", cell="A1", value="x"
    )
    with pytest.raises(ValueError, match="Sheet not found: Missing"):
        internal._apply_xlwings_op(
            cast(internal.XlwingsWorkbookProtocol, workbook),
            cast(dict[str, internal.XlwingsSheetProtocol], sheet_map),
            missing_sheet_op,
            1,
            False,
        )

    bad_values_op = internal.PatchOp.model_construct(
        op="set_range_values",
        sheet="Sheet1",
        range="A1:B2",
        values=[[1], [2]],
    )
    with pytest.raises(ValueError, match="values width does not match range"):
        internal._apply_xlwings_set_range_values(
            cast(internal.XlwingsSheetProtocol, known_sheet),
            bad_values_op,
            index=2,
        )

    with pytest.raises(
        ValueError, match="apply_table_style is supported only on openpyxl backend"
    ):
        internal._apply_xlwings_apply_table_style(
            internal.PatchOp(
                op="apply_table_style",
                sheet="Sheet1",
                range="A1:B2",
                style="TableStyleMedium2",
            )
        )

    cell = known_sheet.range("A1")
    with pytest.raises(ValueError, match="set_value rejects values starting with '='"):
        internal._set_xlwings_cell_value(
            cast(internal.XlwingsRangeProtocol, cell),
            "=1+1",
            auto_formula=False,
            op_name="set_value",
        )
    converted = internal._set_xlwings_cell_value(
        cast(internal.XlwingsRangeProtocol, cell),
        "=1+1",
        auto_formula=True,
        op_name="set_value",
    )
    assert converted.kind == "formula"
    assert cell.formula == "=1+1"


def test_internal_auto_fit_column_resolution_defaults() -> None:
    class _SheetWithoutUsedRange:
        pass

    assert internal._resolve_auto_fit_columns_xlwings(
        cast(internal.XlwingsSheetProtocol, _SheetWithoutUsedRange()), None
    ) == ["A"]

    class _LastCell:
        column = 3

    class _UsedRange:
        last_cell = _LastCell()

    class _SheetWithUsedRange:
        used_range = _UsedRange()

    assert internal._resolve_auto_fit_columns_xlwings(
        cast(internal.XlwingsSheetProtocol, _SheetWithUsedRange()), None
    ) == ["A", "B", "C"]


def test_internal_create_chart_honors_titles_from_data_false() -> None:
    series_1 = MagicMock()
    series_1.Name = "Header-A"
    series_2 = MagicMock()
    series_2.Name = "Header-B"
    series_items = {1: series_1, 2: series_2}
    series_collection = MagicMock()
    series_collection.side_effect = (
        lambda index=None: MagicMock(Count=2) if index is None else series_items[index]
    )

    chart = MagicMock()
    chart.SeriesCollection = series_collection
    chart_object = MagicMock()
    chart_object.Chart = chart
    chart_object.Name = "Chart 1"

    chart_collection = MagicMock()
    chart_collection.Count = 0
    chart_collection.Add.return_value = chart_object
    chart_objects = MagicMock()
    chart_objects.side_effect = lambda index=None: (
        chart_collection if index is None else chart_object
    )

    anchor_range = MagicMock()
    anchor_range.api = MagicMock(Left=10.0, Top=20.0)
    data_range = MagicMock()
    data_range.api = "DATA_API"
    sheet = MagicMock()
    sheet.api = MagicMock(ChartObjects=chart_objects)
    sheet.range.side_effect = lambda ref: {
        "E2": anchor_range,
        "A1:C3": data_range,
    }[ref]

    op = internal.PatchOp(
        op="create_chart",
        sheet="Sheet1",
        chart_type="line",
        data_range="A1:C3",
        anchor_cell="E2",
        titles_from_data=False,
    )

    diff = internal._apply_xlwings_create_chart(
        cast(internal.XlwingsSheetProtocol, sheet), op, index=0
    )

    chart.SetSourceData.assert_called_once_with("DATA_API")
    assert series_1.Name == "Series 1"
    assert series_2.Name == "Series 2"
    assert diff.after is not None
    assert diff.after.kind == "chart"


def test_internal_create_chart_allows_name_matching_new_default_name() -> None:
    chart = MagicMock()
    chart.SeriesCollection = MagicMock(
        side_effect=lambda index=None: MagicMock(Count=0)
    )
    chart_object = MagicMock()
    chart_object.Chart = chart
    chart_object.Name = "Chart 1"

    chart_collection = MagicMock()
    chart_collection.Count = 0
    chart_collection.Add.return_value = chart_object
    chart_objects = MagicMock(side_effect=lambda index=None: chart_collection)

    anchor_range = MagicMock()
    anchor_range.api = MagicMock(Left=15.0, Top=25.0)
    data_range = MagicMock()
    data_range.api = "DATA_API"
    sheet = MagicMock()
    sheet.api = MagicMock(ChartObjects=chart_objects)
    sheet.range.side_effect = lambda ref: {
        "D2": anchor_range,
        "A1:B3": data_range,
    }[ref]

    op = internal.PatchOp(
        op="create_chart",
        sheet="Sheet1",
        chart_type="line",
        data_range="A1:B3",
        anchor_cell="D2",
        chart_name="Chart 1",
    )

    internal._apply_xlwings_create_chart(
        cast(internal.XlwingsSheetProtocol, sheet), op, index=0
    )

    assert chart_object.Name == "Chart 1"
