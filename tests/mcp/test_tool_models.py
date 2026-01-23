from __future__ import annotations

from pydantic import ValidationError
import pytest

from exstruct.mcp.tools import ExtractToolInput, ReadJsonChunkToolInput


def test_extract_tool_input_defaults() -> None:
    payload = ExtractToolInput(xlsx_path="input.xlsx")
    assert payload.mode == "standard"
    assert payload.format == "json"
    assert payload.out_dir is None
    assert payload.out_name is None


def test_read_json_chunk_rejects_invalid_max_bytes() -> None:
    with pytest.raises(ValidationError):
        ReadJsonChunkToolInput(out_path="out.json", max_bytes=0)
