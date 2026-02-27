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
        ValueError, match="apply_table_style requires sheet ListObjects COM API"
    ):
        internal._apply_xlwings_apply_table_style(
            cast(internal.XlwingsSheetProtocol, known_sheet),
            internal.PatchOp(
                op="apply_table_style",
                sheet="Sheet1",
                range="A1:B2",
                style="TableStyleMedium2",
            ),
            index=3,
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


def test_internal_xlwings_add_list_object_retries_with_headers_keyword() -> None:
    class _ListObjects:
        def __init__(self) -> None:
            self.calls: list[tuple[tuple[object, ...], dict[str, object]]] = []

        def Add(self, *args: object, **kwargs: object) -> object:  # noqa: N802
            self.calls.append((args, kwargs))
            if kwargs.get("XlListObjectHasHeaders") == 1:
                return {"status": "ok"}
            raise RuntimeError("header mode required")

    list_objects = _ListObjects()
    source_range_api = MagicMock()
    source_range_api.Address.return_value = "$A$1:$B$3"

    created = internal._xlwings_add_list_object(
        list_objects=list_objects,
        source_range_api=source_range_api,
    )

    assert created == {"status": "ok"}
    assert any("XlListObjectHasHeaders" in kwargs for _, kwargs in list_objects.calls)


def test_internal_resolve_xlwings_list_objects_uses_collection_like_accessor() -> None:
    class _ListObjectsAccessor:
        Count = 0

        @staticmethod
        def Add(*args: object, **kwargs: object) -> object:  # noqa: N802
            return {"status": "unused"}

        def __call__(self) -> object:
            raise RuntimeError("must not call accessor when collection-like")

    class _SheetApi:
        ListObjects = _ListObjectsAccessor()

    resolved = internal._resolve_xlwings_list_objects(
        cast(internal.XlwingsSheetApiProtocol, _SheetApi())
    )
    assert resolved is _SheetApi.ListObjects


def test_internal_apply_table_style_accepts_property_list_objects() -> None:
    class _FakeTable:
        Name = ""
        TableStyle = ""

    class _FakeListObjects:
        Count = 0

        def __init__(self) -> None:
            self.add_calls: list[tuple[tuple[object, ...], dict[str, object]]] = []
            self.created_table: _FakeTable | None = None

        def Add(self, *args: object, **kwargs: object) -> object:  # noqa: N802
            self.add_calls.append((args, kwargs))
            table = _FakeTable()
            self.created_table = table
            return table

    class _FakeRangeApi:
        @staticmethod
        def Address(*args: object) -> str:  # noqa: ARG004
            return "$A$1:$B$3"

    class _FakeRange:
        api = _FakeRangeApi()

    class _FakeSheetApi:
        ListObjects = _FakeListObjects()

    class _FakeSheet:
        name = "Sheet1"
        api = _FakeSheetApi()

        def range(self, ref: str) -> _FakeRange:
            assert ref == "A1:B3"
            return _FakeRange()

    diff = internal._apply_xlwings_apply_table_style(
        cast(internal.XlwingsSheetProtocol, _FakeSheet()),
        internal.PatchOp(
            op="apply_table_style",
            sheet="Sheet1",
            range="A1:B3",
            style="TableStyleMedium2",
        ),
        index=0,
    )

    assert diff.after is not None
    assert diff.after.value == "table=Table1;table_style=TableStyleMedium2"
    assert _FakeSheetApi.ListObjects.add_calls
    assert _FakeSheetApi.ListObjects.created_table is not None
    assert _FakeSheetApi.ListObjects.created_table.Name == "Table1"
    assert _FakeSheetApi.ListObjects.created_table.TableStyle == "TableStyleMedium2"


def test_internal_xlwings_add_list_object_falls_back_to_address_source() -> None:
    class _ListObjects:
        def __init__(self) -> None:
            self.calls: list[tuple[tuple[object, ...], dict[str, object]]] = []

        def Add(self, *args: object, **kwargs: object) -> object:  # noqa: N802
            self.calls.append((args, kwargs))
            source = kwargs.get("Source")
            if source is None and len(args) >= 2:
                source = args[1]
            if isinstance(source, str):
                return {"status": "address-ok"}
            raise RuntimeError("source must be A1 string")

    list_objects = _ListObjects()
    source_range_api = MagicMock()
    source_range_api.Address.return_value = "$C$2:$E$9"

    created = internal._xlwings_add_list_object(
        list_objects=list_objects,
        source_range_api=source_range_api,
    )

    assert created == {"status": "address-ok"}
    assert any(
        isinstance(kwargs.get("Source"), str)
        or (len(args) >= 2 and isinstance(args[1], str))
        for args, kwargs in list_objects.calls
    )


def test_internal_normalize_table_range_address_handles_external_notation() -> None:
    normalized = internal._normalize_table_range_address(
        "='[Book1.xlsx]Sales Data'!$B$2:$D$11"
    )
    assert normalized == "B2:D11"


def test_internal_create_chart_honors_titles_from_data_false() -> None:
    series_1 = MagicMock()
    series_1.Name = "Header-A"
    series_2 = MagicMock()
    series_2.Name = "Header-B"
    series_items = {1: series_1, 2: series_2}

    class _SeriesCollection:
        Count = 2

        def Item(self, index: int) -> MagicMock:
            return series_items[index]

    chart = MagicMock()
    chart.SeriesCollection = MagicMock(return_value=_SeriesCollection())
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
    chart.SeriesCollection = MagicMock(return_value=MagicMock(Count=0))
    chart_object = MagicMock()
    chart_object.Chart = chart
    chart_object.Name = "Existing"

    chart_collection = MagicMock()
    chart_collection.Count = 0
    chart_collection.Add.return_value = chart_object
    chart_objects = MagicMock(
        side_effect=lambda index=None: chart_collection
        if index is None
        else chart_object
    )

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

    diff = internal._apply_xlwings_create_chart(
        cast(internal.XlwingsSheetProtocol, sheet), op, index=0
    )

    chart.SetSourceData.assert_called_once_with("DATA_API")
    chart_collection.Add.assert_called_once()
    assert diff.after is not None
    assert diff.after.kind == "chart"
    assert chart_object.Name == "Chart 1"


def test_internal_get_com_collection_item_uses_item_fallback() -> None:
    class _Collection:
        def __call__(self, index: int) -> object:
            raise TypeError("not callable in this dispatch mode")

        def Item(self, index: int) -> str:  # noqa: N802
            return f"item-{index}"

    item = internal._get_com_collection_item(_Collection(), 2)
    assert item == "item-2"


def test_internal_get_com_collection_item_raises_on_both_paths_failure() -> None:
    class _Collection:
        def __call__(self, index: int) -> object:
            raise TypeError("call failed")

        def Item(self, index: int) -> object:  # noqa: N802
            raise RuntimeError("item failed")

    with pytest.raises(ValueError, match="COM collection item access failed"):
        internal._get_com_collection_item(_Collection(), 1)


@pytest.mark.parametrize(
    ("chart_type", "expected_chart_type_id"),
    [
        ("line", 4),
        ("column", 51),
        ("bar", 57),
        ("area", 1),
        ("pie", 5),
        ("doughnut", -4120),
        ("scatter", -4169),
        ("radar", -4151),
    ],
)  # type: ignore[misc]
def test_internal_resolve_chart_type_id_supports_phase1_major_types(
    chart_type: str, expected_chart_type_id: int
) -> None:
    assert internal._resolve_chart_type_id(chart_type) == expected_chart_type_id


def test_internal_resolve_chart_type_id_supports_aliases() -> None:
    assert internal._resolve_chart_type_id("column_clustered") == 51
    assert internal._resolve_chart_type_id("bar_clustered") == 57
    assert internal._resolve_chart_type_id("xy_scatter") == -4169
    assert internal._resolve_chart_type_id("donut") == -4120


def test_internal_create_chart_supports_multi_ranges_and_sheet_qualified_refs() -> None:
    first_series = MagicMock()

    class _SeriesCollection:
        def __init__(self) -> None:
            self._items: list[MagicMock] = [first_series]
            self.Count = 1

        def Item(self, index: int) -> MagicMock:  # noqa: N802
            return self._items[index - 1]

        def NewSeries(self) -> MagicMock:  # noqa: N802
            item = MagicMock()
            self._items.append(item)
            self.Count = len(self._items)
            return item

    series_collection = _SeriesCollection()
    chart = MagicMock()
    chart.SeriesCollection = MagicMock(return_value=series_collection)
    chart.ChartTitle = MagicMock()
    axis_x = MagicMock()
    axis_x.AxisTitle = MagicMock()
    axis_y = MagicMock()
    axis_y.AxisTitle = MagicMock()
    chart.Axes = MagicMock(
        side_effect=lambda axis_type: {1: axis_x, 2: axis_y}[axis_type]
    )
    chart_object = MagicMock()
    chart_object.Chart = chart
    chart_object.Name = "Chart 1"

    chart_collection = MagicMock()
    chart_collection.Count = 0
    chart_collection.Add.return_value = chart_object
    chart_objects = MagicMock(side_effect=lambda index=None: chart_collection)

    anchor_range = MagicMock()
    anchor_range.api = MagicMock(Left=30.0, Top=40.0)
    chart_sheet = MagicMock()
    chart_sheet.api = MagicMock(ChartObjects=chart_objects)
    chart_sheet.range.side_effect = lambda ref: {"E2": anchor_range}[ref]

    range_b = MagicMock()
    range_b.api = "B_API"
    range_c = MagicMock()
    range_c.api = "C_API"
    range_a = MagicMock()
    range_a.api = "A_API"
    data_sheet = MagicMock()
    data_sheet.name = "Data"
    data_sheet.range.side_effect = lambda ref: {
        "B2:B10": range_b,
        "C2:C10": range_c,
        "A2:A10": range_a,
    }[ref]
    chart_sheet.book = MagicMock(sheets={"Data": data_sheet})

    op = internal.PatchOp(
        op="create_chart",
        sheet="Chart",
        chart_type="line",
        data_range=["'Data'!B2:B10", "'Data'!C2:C10"],
        category_range="'Data'!A2:A10",
        anchor_cell="E2",
        chart_title="Revenue",
        x_axis_title="Month",
        y_axis_title="Amount",
    )

    diff = internal._apply_xlwings_create_chart(
        cast(internal.XlwingsSheetProtocol, chart_sheet), op, index=0
    )

    chart.SetSourceData.assert_called_once_with("B_API")
    assert first_series.Values == "B_API"
    assert first_series.XValues == "A_API"
    second_series = series_collection.Item(2)
    assert second_series.Values == "C_API"
    assert second_series.XValues == "A_API"
    assert axis_x.AxisTitle.Text == "Month"
    assert axis_y.AxisTitle.Text == "Amount"
    assert chart.ChartTitle.Text == "Revenue"
    assert diff.after is not None
    assert diff.after.kind == "chart"


def test_internal_patch_op_error_adds_error_code_and_failed_field() -> None:
    op = internal.PatchOp(
        op="create_chart",
        sheet="Sheet1",
        chart_type="line",
        data_range="A1:B2",
        anchor_cell="D2",
    )
    err = internal.PatchOpError.from_op(
        2, op, ValueError("Invalid chart range reference: bad")
    )
    assert err.detail.error_code == "invalid_range"
    assert err.detail.failed_field == "data_range"


def test_internal_classify_sheet_not_found_uses_category_failed_field() -> None:
    classified = internal._classify_known_patch_error(
        "create_chart sheet not found for category range reference: missing_sheet"
    )
    assert classified == ("sheet_not_found", "category_range")


def test_internal_patch_op_error_classifies_table_style_and_add_failures() -> None:
    op = internal.PatchOp(
        op="apply_table_style",
        sheet="Sheet1",
        range="A1:D10",
        style="BadStyle",
    )
    style_error = internal.PatchOpError.from_op(
        1, op, ValueError("apply_table_style invalid table style: 'BadStyle'")
    )
    assert style_error.detail.error_code == "table_style_invalid"
    assert style_error.detail.failed_field == "style"

    add_error = internal.PatchOpError.from_op(
        1,
        op,
        ValueError(
            "apply_table_style failed to add table after COM Add signature retries."
        ),
    )
    assert add_error.detail.error_code == "list_object_add_failed"
    assert add_error.detail.failed_field == "range"
