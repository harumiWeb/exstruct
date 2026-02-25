from __future__ import annotations

from pydantic import ValidationError
import pytest

from exstruct.mcp.tools import (
    DescribeOpToolInput,
    DescribeOpToolOutput,
    ExtractToolInput,
    ListOpsToolOutput,
    MakeToolInput,
    MakeToolOutput,
    PatchToolInput,
    PatchToolOutput,
    ReadCellsToolInput,
    ReadFormulasToolInput,
    ReadJsonChunkToolInput,
    ReadRangeToolInput,
    RuntimeInfoToolOutput,
)


def test_extract_tool_input_defaults() -> None:
    payload = ExtractToolInput(xlsx_path="input.xlsx")
    assert payload.mode == "standard"
    assert payload.format == "json"
    assert payload.out_dir is None
    assert payload.out_name is None


def test_read_json_chunk_rejects_invalid_max_bytes() -> None:
    with pytest.raises(ValidationError):
        ReadJsonChunkToolInput(out_path="out.json", max_bytes=0)


def test_patch_tool_input_defaults() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[{"op": "add_sheet", "sheet": "New"}],
    )
    assert payload.out_dir is None
    assert payload.out_name is None
    assert payload.on_conflict is None
    assert payload.dry_run is False
    assert payload.return_inverse_ops is False
    assert payload.preflight_formula_check is False
    assert payload.backend == "auto"
    assert payload.mirror_artifact is False
    assert payload.sheet is None


def test_make_tool_input_defaults() -> None:
    payload = MakeToolInput(out_path="output.xlsx")
    assert payload.ops == []
    assert payload.on_conflict is None
    assert payload.dry_run is False
    assert payload.return_inverse_ops is False
    assert payload.preflight_formula_check is False
    assert payload.backend == "auto"
    assert payload.mirror_artifact is False
    assert payload.sheet is None


def test_patch_tool_input_applies_top_level_sheet_fallback() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        sheet="Sheet1",
        ops=[{"op": "set_value", "cell": "A1", "value": "x"}],
    )
    assert payload.ops[0].sheet == "Sheet1"


def test_patch_tool_input_prioritizes_op_sheet_over_top_level() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        sheet="Sheet1",
        ops=[{"op": "set_value", "sheet": "Data", "cell": "A1", "value": "x"}],
    )
    assert payload.ops[0].sheet == "Data"


def test_patch_tool_input_rejects_add_sheet_without_explicit_sheet() -> None:
    with pytest.raises(ValidationError, match="add_sheet\\) is missing sheet"):
        PatchToolInput(
            xlsx_path="input.xlsx",
            sheet="Sheet1",
            ops=[{"op": "add_sheet"}],
        )


def test_patch_tool_input_accepts_add_sheet_name_alias() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[{"op": "add_sheet", "name": "Data"}],
    )
    assert payload.ops[0].sheet == "Data"


def test_patch_tool_input_rejects_unresolved_sheet_for_non_add_sheet() -> None:
    with pytest.raises(ValidationError, match="missing sheet"):
        PatchToolInput(
            xlsx_path="input.xlsx",
            ops=[{"op": "set_value", "cell": "A1", "value": "x"}],
        )


def test_make_tool_input_applies_top_level_sheet_fallback() -> None:
    payload = MakeToolInput(
        out_path="output.xlsx",
        sheet="Sheet1",
        ops=[{"op": "set_value", "cell": "A1", "value": "x"}],
    )
    assert payload.ops[0].sheet == "Sheet1"


def test_make_tool_input_accepts_add_sheet_name_alias() -> None:
    payload = MakeToolInput(
        out_path="output.xlsx",
        ops=[{"op": "add_sheet", "name": "Data"}],
    )
    assert payload.ops[0].sheet == "Data"


def test_patch_and_make_tool_output_defaults() -> None:
    patch_output = PatchToolOutput(
        out_path="out.xlsx", patch_diff=[], engine="openpyxl"
    )
    make_output = MakeToolOutput(out_path="out.xlsx", patch_diff=[], engine="openpyxl")
    assert patch_output.mirrored_out_path is None
    assert make_output.mirrored_out_path is None


def test_patch_tool_input_accepts_design_ops() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[
            {
                "op": "set_dimensions",
                "sheet": "Sheet1",
                "rows": [1, 2],
                "row_height": 20,
                "columns": ["A", 2],
                "column_width": 18,
            }
        ],
    )
    assert payload.ops[0].op == "set_dimensions"


def test_patch_tool_input_accepts_merge_and_alignment_ops() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[
            {"op": "merge_cells", "sheet": "Sheet1", "range": "A1:B1"},
            {
                "op": "set_alignment",
                "sheet": "Sheet1",
                "range": "A1:B1",
                "horizontal_align": "center",
                "vertical_align": "center",
                "wrap_text": True,
            },
        ],
    )
    assert payload.ops[0].op == "merge_cells"
    assert payload.ops[1].op == "set_alignment"


def test_patch_tool_input_accepts_set_style_op() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[
            {
                "op": "set_style",
                "sheet": "Sheet1",
                "range": "A1:B1",
                "bold": True,
                "fill_color": "d9e1f2",
                "horizontal_align": "center",
            }
        ],
    )
    assert payload.ops[0].op == "set_style"
    assert payload.ops[0].fill_color == "#D9E1F2"


def test_patch_tool_input_accepts_apply_table_style_op() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[
            {
                "op": "apply_table_style",
                "sheet": "Sheet1",
                "range": "A1:B3",
                "style": "TableStyleMedium2",
                "table_name": "SalesTable",
            }
        ],
    )
    assert payload.ops[0].op == "apply_table_style"
    assert payload.ops[0].style == "TableStyleMedium2"
    assert payload.ops[0].table_name == "SalesTable"


def test_patch_tool_input_accepts_auto_fit_columns_op() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[
            {
                "op": "auto_fit_columns",
                "sheet": "Sheet1",
                "columns": ["A", 2],
                "min_width": 8,
                "max_width": 40,
            }
        ],
    )
    assert payload.ops[0].op == "auto_fit_columns"
    assert payload.ops[0].columns == ["A", 2]
    assert payload.ops[0].min_width == 8
    assert payload.ops[0].max_width == 40


def test_patch_tool_input_accepts_set_font_size_op() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[
            {
                "op": "set_font_size",
                "sheet": "Sheet1",
                "cell": "A1",
                "font_size": 14,
            }
        ],
    )
    assert payload.ops[0].op == "set_font_size"


def test_patch_tool_input_accepts_set_font_color_op() -> None:
    payload = PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[
            {
                "op": "set_font_color",
                "sheet": "Sheet1",
                "range": "A1:B1",
                "color": "1f4e79",
            }
        ],
    )
    assert payload.ops[0].op == "set_font_color"
    assert payload.ops[0].color == "#1F4E79"


def test_patch_tool_input_rejects_invalid_horizontal_align() -> None:
    with pytest.raises(ValidationError):
        PatchToolInput(
            xlsx_path="input.xlsx",
            ops=[
                {
                    "op": "set_alignment",
                    "sheet": "Sheet1",
                    "cell": "A1",
                    "horizontal_align": "middle",
                }
            ],
        )


def test_patch_tool_input_rejects_alignment_without_target_fields() -> None:
    with pytest.raises(ValidationError):
        PatchToolInput(
            xlsx_path="input.xlsx",
            ops=[{"op": "set_alignment", "sheet": "Sheet1", "cell": "A1"}],
        )


def test_read_range_tool_input_defaults() -> None:
    payload = ReadRangeToolInput(out_path="out.json", range="A1:B2")
    assert payload.sheet is None
    assert payload.include_formulas is False
    assert payload.include_empty is True
    assert payload.max_cells == 10_000


def test_read_range_tool_input_rejects_invalid_max_cells() -> None:
    with pytest.raises(ValidationError):
        ReadRangeToolInput(out_path="out.json", range="A1:B2", max_cells=0)


def test_read_cells_tool_input_rejects_empty_addresses() -> None:
    with pytest.raises(ValidationError):
        ReadCellsToolInput(out_path="out.json", addresses=[])


def test_read_formulas_tool_input_defaults() -> None:
    payload = ReadFormulasToolInput(out_path="out.json")
    assert payload.sheet is None
    assert payload.range is None
    assert payload.include_values is False


def test_runtime_info_tool_output_model() -> None:
    payload = RuntimeInfoToolOutput(
        root="C:\\data",
        cwd="C:\\workspace",
        platform="win32",
        path_examples={
            "relative": "outputs/book.xlsx",
            "absolute": "C:\\data\\outputs\\book.xlsx",
        },
    )
    assert payload.path_examples.relative == "outputs/book.xlsx"


def test_describe_op_tool_input_accepts_op_name() -> None:
    payload = DescribeOpToolInput(op="set_fill_color")
    assert payload.op == "set_fill_color"


def test_list_ops_tool_output_defaults() -> None:
    payload = ListOpsToolOutput()
    assert payload.ops == []


def test_describe_op_tool_output_defaults() -> None:
    payload = DescribeOpToolOutput(op="set_value")
    assert payload.required == []
    assert payload.optional == []
    assert payload.constraints == []
    assert payload.example == {}
    assert payload.aliases == {}
