from __future__ import annotations

from pydantic import ValidationError
import pytest

from exstruct.mcp.tools import (
    ExtractToolInput,
    PatchToolInput,
    ReadCellsToolInput,
    ReadFormulasToolInput,
    ReadJsonChunkToolInput,
    ReadRangeToolInput,
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
