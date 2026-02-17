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
