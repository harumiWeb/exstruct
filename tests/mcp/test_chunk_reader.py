from __future__ import annotations

from collections.abc import Mapping
import json
from pathlib import Path

import pytest

from exstruct.mcp import chunk_reader
from exstruct.mcp.chunk_reader import (
    ReadJsonChunkFilter,
    ReadJsonChunkRequest,
    read_json_chunk,
)


def _write_json(path: Path, data: Mapping[str, object]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def test_read_json_chunk_raw(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(out_path=out, max_bytes=10_000)
    result = read_json_chunk(request)
    assert json.loads(result.chunk) == data
    assert result.next_cursor is None


def test_read_json_chunk_raw_too_large(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(out_path=out, max_bytes=10)
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_raw_too_large_has_retry_hint(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(out_path=out, max_bytes=10)
    with pytest.raises(ValueError, match=r"Retry with `sheet` .* or `filter`"):
        read_json_chunk(request)


def test_read_json_chunk_with_filters(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {
            "Sheet1": {
                "rows": [
                    {"r": 1, "c": {"0": "A", "1": "B"}},
                    {"r": 2, "c": {"0": "C", "1": "D"}},
                ]
            }
        },
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(rows=(1, 1), cols=(1, 1)),
    )
    result = read_json_chunk(request)
    payload = json.loads(result.chunk)
    rows = payload["sheet"]["rows"]
    assert len(rows) == 1
    assert rows[0]["r"] == 1
    assert rows[0]["c"] == {"0": "A"}


def test_read_json_chunk_requires_sheet(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"A": {"rows": []}, "B": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(rows=(1, 1)),
    )
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_requires_sheet_lists_available_names(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"A": {"rows": []}, "B": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(rows=(1, 1)),
    )
    with pytest.raises(ValueError, match=r"Available sheets: A, B"):
        read_json_chunk(request)


def test_read_json_chunk_invalid_cursor(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        cursor="bad",
        max_bytes=10_000,
    )
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_cursor_beyond_rows(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {"Sheet1": {"rows": [{"r": 1, "c": {"0": "A"}}]}},
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        cursor="2",
        max_bytes=10_000,
    )
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_rejects_directory(tmp_path: Path) -> None:
    out_dir = tmp_path / "out"
    out_dir.mkdir()
    request = ReadJsonChunkRequest(out_path=out_dir, sheet="Sheet1")
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_rejects_non_object_root(tmp_path: Path) -> None:
    out = tmp_path / "out.json"
    out.write_text(json.dumps([1, 2, 3]), encoding="utf-8")
    request = ReadJsonChunkRequest(out_path=out, sheet="Sheet1")
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_rejects_invalid_sheets_mapping(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": ["Sheet1"]}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(out_path=out, sheet="Sheet1")
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_rejects_missing_sheet(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(out_path=out, sheet="Missing")
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_rejects_negative_cursor(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out, sheet="Sheet1", cursor="-1", max_bytes=10_000
    )
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_warns_on_row_filter_inversion(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {"Sheet1": {"rows": [{"r": 1, "c": {"0": "A"}}]}},
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(rows=(2, 1)),
    )
    result = read_json_chunk(request)
    assert any("Row filter ignored" in warning for warning in result.warnings)


def test_read_json_chunk_warns_on_col_filter_inversion(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {"Sheet1": {"rows": [{"r": 1, "c": {"0": "A"}}]}},
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(cols=(2, 1)),
    )
    result = read_json_chunk(request)
    assert any("Column filter ignored" in warning for warning in result.warnings)


def test_read_json_chunk_with_alpha_col_keys(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {
            "Sheet1": {
                "rows": [
                    {"r": 1, "c": {"A": "A1", "B": "B1", "AA": "AA1", "ab": "AB1"}},
                ]
            }
        },
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(cols=(1, 2)),
    )
    result = read_json_chunk(request)
    payload = json.loads(result.chunk)
    assert payload["sheet"]["rows"][0]["c"] == {"A": "A1", "B": "B1"}

    request_aa = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(cols=(27, 28)),
    )
    result_aa = read_json_chunk(request_aa)
    payload_aa = json.loads(result_aa.chunk)
    assert payload_aa["sheet"]["rows"][0]["c"] == {"AA": "AA1", "ab": "AB1"}


def test_read_json_chunk_warns_on_base_payload_exceeds_max_bytes(
    tmp_path: Path,
) -> None:
    data = {
        "book_name": "book",
        "sheets": {"Sheet1": {"rows": [{"r": 1, "c": {"0": "A"}}]}},
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(out_path=out, sheet="Sheet1", max_bytes=1)
    result = read_json_chunk(request)
    assert any("Base payload exceeds" in warning for warning in result.warnings)


def test_read_json_chunk_warns_on_too_small_max_bytes(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {"Sheet1": {"rows": [{"r": 1, "c": {"0": "x" * 200}}]}},
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    base_payload = {
        "book_name": "book",
        "sheet_name": "Sheet1",
        "sheet": {"rows": []},
    }
    base_json = chunk_reader._serialize_json(base_payload)
    max_bytes = len(base_json.encode("utf-8")) + 1
    request = ReadJsonChunkRequest(out_path=out, sheet="Sheet1", max_bytes=max_bytes)
    result = read_json_chunk(request)
    assert any("max_bytes too small" in warning for warning in result.warnings)


def test_read_json_chunk_missing_output_file_raises(tmp_path: Path) -> None:
    request = ReadJsonChunkRequest(out_path=tmp_path / "missing.json", sheet="Sheet1")
    with pytest.raises(FileNotFoundError, match="Output file not found"):
        read_json_chunk(request)


def test_read_json_chunk_single_sheet_without_explicit_sheet(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {"Only": {"rows": [{"r": 1, "c": {"0": "A"}}]}},
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet=None,
        filter=ReadJsonChunkFilter(rows=(1, 1)),
        max_bytes=10_000,
    )
    result = read_json_chunk(request)
    payload = json.loads(result.chunk)
    assert payload["sheet_name"] == "Only"


def test_chunk_reader_private_helpers_cover_boundary_paths() -> None:
    rows = [{"r": 1, "c": {"0": "A"}}, {"r": 2, "c": "not-dict"}]
    filtered = chunk_reader._filter_cols(rows, (1, 1), warnings=[])
    assert filtered[1]["c"] == "not-dict"
    assert chunk_reader._extract_rows({"rows": "invalid"}) == []
    assert chunk_reader._col_in_range("bad-key", 0, 3) is False
    assert chunk_reader._row_index({"r": "x"}) == -1
    assert chunk_reader._parse_col_index("A1") is None


def test_build_sheet_chunk_returns_next_cursor_when_second_row_overflows() -> None:
    data = {"book_name": "book"}
    sheet_name = "Sheet1"
    rows = [
        {"r": 1, "c": {"0": "x" * 80}},
        {"r": 2, "c": {"0": "y" * 80}},
    ]
    sheet_data = {"rows": rows}

    first_payload = {
        "book_name": "book",
        "sheet_name": sheet_name,
        "sheet": {"rows": [rows[0]]},
    }
    second_payload = {
        "book_name": "book",
        "sheet_name": sheet_name,
        "sheet": {"rows": rows},
    }
    first_size = len(chunk_reader._serialize_json(first_payload).encode("utf-8"))
    second_size = len(chunk_reader._serialize_json(second_payload).encode("utf-8"))
    max_bytes = first_size + 1
    assert second_size > max_bytes

    chunk, next_cursor, warnings = chunk_reader._build_sheet_chunk(
        data,
        sheet_name,
        sheet_data,
        rows,
        cursor=None,
        max_bytes=max_bytes,
    )
    payload = json.loads(chunk)
    assert len(payload["sheet"]["rows"]) == 1
    assert next_cursor == "1"
    assert warnings == []
