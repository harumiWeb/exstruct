from __future__ import annotations

import json
from pathlib import Path
from typing import Any, cast

from pydantic import BaseModel, Field

from .io import PathPolicy


class ReadJsonChunkFilter(BaseModel):
    """Filter options for JSON chunk extraction."""

    rows: tuple[int, int] | None = Field(
        default=None, description="Row range (1-based, inclusive)."
    )
    cols: tuple[int, int] | None = Field(
        default=None, description="Column range (1-based, inclusive)."
    )


class ReadJsonChunkRequest(BaseModel):
    """Input model for JSON chunk extraction."""

    out_path: Path
    sheet: str | None = None
    max_bytes: int = Field(default=50_000, ge=1)
    filter: ReadJsonChunkFilter | None = Field(default=None)  # noqa: A003
    cursor: str | None = None


class ReadJsonChunkResult(BaseModel):
    """Output model for JSON chunk extraction."""

    chunk: str
    next_cursor: str | None = None
    warnings: list[str] = Field(default_factory=list)


def read_json_chunk(
    request: ReadJsonChunkRequest, *, policy: PathPolicy | None = None
) -> ReadJsonChunkResult:
    """Read a JSON chunk from an ExStruct output file.

    Args:
        request: Chunk request payload.
        policy: Optional path policy for access control.

    Returns:
        Chunk result with JSON string payload and cursor.

    Raises:
        FileNotFoundError: If the output file does not exist.
        ValueError: If the request is invalid or violates policy.
    """
    resolved = _resolve_output_path(request.out_path, policy=policy)
    raw_text = _read_text(resolved)

    if request.sheet is None and request.filter is None and request.cursor is None:
        return _chunk_raw_text(raw_text, request.max_bytes)

    data = _parse_json(raw_text)
    sheet_name, sheet_data = _select_sheet(data, request.sheet)
    rows = _extract_rows(sheet_data)
    filtered_rows, warnings = _apply_filters(rows, request.filter)
    chunk, next_cursor, more_warnings = _build_sheet_chunk(
        data,
        sheet_name,
        sheet_data,
        filtered_rows,
        request.cursor,
        request.max_bytes,
    )
    warnings.extend(more_warnings)
    return ReadJsonChunkResult(
        chunk=chunk,
        next_cursor=next_cursor,
        warnings=warnings,
    )


def _resolve_output_path(path: Path, *, policy: PathPolicy | None) -> Path:
    """Resolve and validate the output path.

    Args:
        path: Output file path.
        policy: Optional path policy.

    Returns:
        Resolved path.

    Raises:
        FileNotFoundError: If the output file does not exist.
        ValueError: If the path violates the policy or is not a file.
    """
    resolved = policy.ensure_allowed(path) if policy else path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"Output file not found: {resolved}")
    if not resolved.is_file():
        raise ValueError(f"Output path is not a file: {resolved}")
    return resolved


def _read_text(path: Path) -> str:
    """Read UTF-8 text from disk.

    Args:
        path: File path.

    Returns:
        File contents as text.
    """
    return path.read_text(encoding="utf-8")


def _parse_json(text: str) -> dict[str, Any]:
    """Parse JSON into a dictionary.

    Args:
        text: JSON string.

    Returns:
        Parsed JSON object.
    """
    parsed = json.loads(text)
    if not isinstance(parsed, dict):
        raise ValueError("Invalid workbook JSON: expected object at root.")
    return cast(dict[str, Any], parsed)


def _chunk_raw_text(text: str, max_bytes: int) -> ReadJsonChunkResult:
    """Return a raw JSON chunk without parsing.

    Args:
        text: JSON text.
        max_bytes: Maximum bytes to return.

    Returns:
        Chunk result with optional cursor.

    Raises:
        ValueError: If the text exceeds max_bytes.
    """
    payload_bytes = text.encode("utf-8")
    if len(payload_bytes) <= max_bytes:
        return ReadJsonChunkResult(chunk=text, next_cursor=None, warnings=[])
    raise ValueError("Output is too large. Specify sheet or filter to chunk.")


def _select_sheet(
    data: dict[str, Any], sheet: str | None
) -> tuple[str, dict[str, Any]]:
    """Select a sheet from the workbook payload.

    Args:
        data: Parsed workbook JSON.
        sheet: Optional sheet name.

    Returns:
        Sheet name and sheet data.

    Raises:
        ValueError: If sheet selection is ambiguous or missing.
    """
    sheets = data.get("sheets", {})
    if not isinstance(sheets, dict):
        raise ValueError("Invalid workbook JSON: sheets is not a mapping.")
    if sheet is not None:
        if sheet not in sheets:
            raise ValueError(f"Sheet not found: {sheet}")
        return sheet, sheets[sheet]
    if len(sheets) == 1:
        only_name = next(iter(sheets.keys()))
        return only_name, sheets[only_name]
    raise ValueError("Sheet is required when multiple sheets exist.")


def _extract_rows(sheet_data: dict[str, Any]) -> list[dict[str, Any]]:
    """Extract rows from sheet data.

    Args:
        sheet_data: Sheet JSON data.

    Returns:
        List of row dictionaries.
    """
    rows = sheet_data.get("rows", [])
    if not isinstance(rows, list):
        return []
    return rows


def _apply_filters(
    rows: list[dict[str, Any]], filter_data: ReadJsonChunkFilter | None
) -> tuple[list[dict[str, Any]], list[str]]:
    """Apply row/column filters to rows.

    Args:
        rows: Row dictionaries.
        filter_data: Optional filter data.

    Returns:
        Filtered rows and warnings.
    """
    if filter_data is None:
        return rows, []
    warnings: list[str] = []
    filtered_rows = rows
    if filter_data.rows is not None:
        filtered_rows = _filter_rows(filtered_rows, filter_data.rows, warnings)
    if filter_data.cols is not None:
        filtered_rows = _filter_cols(filtered_rows, filter_data.cols, warnings)
    return filtered_rows, warnings


def _filter_rows(
    rows: list[dict[str, Any]], row_range: tuple[int, int], warnings: list[str]
) -> list[dict[str, Any]]:
    """Filter rows by row range.

    Args:
        rows: Row dictionaries.
        row_range: Row range tuple (1-based, inclusive).
        warnings: Warning collector.

    Returns:
        Filtered rows.
    """
    start, end = row_range
    if start > end:
        warnings.append("Row filter ignored because start > end.")
        return rows
    return [row for row in rows if start <= _row_index(row) <= end]


def _filter_cols(
    rows: list[dict[str, Any]], col_range: tuple[int, int], warnings: list[str]
) -> list[dict[str, Any]]:
    """Filter columns within each row by column range.

    Args:
        rows: Row dictionaries.
        col_range: Column range tuple (1-based, inclusive).
        warnings: Warning collector.

    Returns:
        Rows with filtered column maps.
    """
    start, end = col_range
    if start > end:
        warnings.append("Column filter ignored because start > end.")
        return rows
    start_index = start - 1
    end_index = end - 1
    filtered: list[dict[str, Any]] = []
    for row in rows:
        cols = row.get("c")
        if not isinstance(cols, dict):
            filtered.append(row)
            continue
        new_cols = {
            key: value
            for key, value in cols.items()
            if _col_in_range(key, start_index, end_index)
        }
        new_row = dict(row)
        new_row["c"] = new_cols
        filtered.append(new_row)
    return filtered


def _col_in_range(key: str, start: int, end: int) -> bool:
    """Check if a column key is within range.

    Args:
        key: Column key string.
        start: Start index (0-based).
        end: End index (0-based).

    Returns:
        True if the column index is within range.
    """
    try:
        index = int(key)
    except (TypeError, ValueError):
        return False
    return start <= index <= end


def _row_index(row: dict[str, Any]) -> int:
    """Extract row index from a row dictionary.

    Args:
        row: Row dictionary.

    Returns:
        Row index or -1 if unavailable.
    """
    value = row.get("r")
    if isinstance(value, int):
        return value
    return -1


def _build_sheet_chunk(
    data: dict[str, Any],
    sheet_name: str,
    sheet_data: dict[str, Any],
    rows: list[dict[str, Any]],
    cursor: str | None,
    max_bytes: int,
) -> tuple[str, str | None, list[str]]:
    """Build a JSON chunk for a sheet.

    Args:
        data: Workbook JSON data.
        sheet_name: Target sheet name.
        sheet_data: Sheet JSON data.
        rows: Filtered rows.
        cursor: Optional cursor (row index in filtered list).
        max_bytes: Maximum payload size in bytes.

    Returns:
        Tuple of JSON chunk, next cursor, and warnings.
    """
    warnings: list[str] = []
    start_index = _parse_cursor(cursor)
    if start_index > len(rows):
        raise ValueError("Cursor is beyond the filtered row count.")
    remaining_rows = rows[start_index:]
    sheet_payload = dict(sheet_data)
    sheet_payload["rows"] = []
    payload = {
        "book_name": data.get("book_name"),
        "sheet_name": sheet_name,
        "sheet": sheet_payload,
    }
    base_json = _serialize_json(payload)
    if _json_size(base_json) > max_bytes:
        warnings.append("Base payload exceeds max_bytes; returning without rows.")
        return base_json, None, warnings

    selected: list[dict[str, Any]] = []
    next_cursor = None
    for offset, row in enumerate(remaining_rows):
        selected.append(row)
        sheet_payload["rows"] = selected
        candidate = _serialize_json(payload)
        if _json_size(candidate) > max_bytes:
            if len(selected) == 1:
                warnings.append("max_bytes too small; returning a single row chunk.")
                next_cursor = (
                    str(start_index + 1) if (start_index + 1) < len(rows) else None
                )
                return candidate, next_cursor, warnings
            selected.pop()
            sheet_payload["rows"] = selected
            next_cursor_index = start_index + offset
            next_cursor = (
                str(next_cursor_index) if next_cursor_index < len(rows) else None
            )
            return _serialize_json(payload), next_cursor, warnings

    sheet_payload["rows"] = selected
    return _serialize_json(payload), None, warnings


def _parse_cursor(cursor: str | None) -> int:
    """Parse cursor into a start index.

    Args:
        cursor: Cursor string.

    Returns:
        Parsed start index.

    Raises:
        ValueError: If cursor is invalid.
    """
    if cursor is None:
        return 0
    try:
        value = int(cursor)
    except (TypeError, ValueError) as exc:
        raise ValueError("Cursor must be an integer string.") from exc
    if value < 0:
        raise ValueError("Cursor must be non-negative.")
    return value


def _serialize_json(payload: dict[str, Any]) -> str:
    """Serialize a payload to JSON string.

    Args:
        payload: JSON-serializable payload.

    Returns:
        Serialized JSON text.
    """
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def _json_size(payload: str) -> int:
    """Return payload size in bytes.

    Args:
        payload: JSON text.

    Returns:
        Payload size in bytes.
    """
    return len(payload.encode("utf-8"))
