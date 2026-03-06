from __future__ import annotations

import re

from pydantic import BaseModel, ConfigDict

_A1_PATTERN = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*$")
_A1_RANGE_PATTERN = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*:[A-Za-z]{1,3}[1-9][0-9]*$")
_COLUMN_LABEL_PATTERN = re.compile(r"^[A-Za-z]{1,3}$")
_QUALIFIED_A1_RANGE_PATTERN = re.compile(
    r"^(?:(?P<sheet>'(?:[^']|'')+'|[^'!]+)!)?"
    r"(?P<range>[A-Za-z]{1,3}[1-9][0-9]*:[A-Za-z]{1,3}[1-9][0-9]*)$"
)


class QualifiedA1Range(BaseModel):
    """Parsed A1 range with optional sheet qualifier."""

    model_config = ConfigDict(frozen=True)

    sheet: str | None
    range_ref: str


class SheetRangeSelection(BaseModel):
    """Normalized sheet and range selection for render-style calls."""

    model_config = ConfigDict(frozen=True)

    sheet: str | None
    range_ref: str | None


def split_a1(value: str) -> tuple[str, int]:
    """Split A1 notation into normalized (column_label, row_index)."""
    if not _A1_PATTERN.match(value):
        raise ValueError(f"Invalid cell reference: {value}")
    idx = 0
    for index, char in enumerate(value):
        if char.isdigit():
            idx = index
            break
    column = value[:idx].upper()
    row = int(value[idx:])
    return column, row


def column_label_to_index(label: str) -> int:
    """Convert Excel-style column label (A/AA) to 1-based index."""
    normalized = label.strip().upper()
    if not _COLUMN_LABEL_PATTERN.match(normalized):
        raise ValueError(f"Invalid column label: {label}")
    index = 0
    for char in normalized:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def column_index_to_label(index: int) -> str:
    """Convert 1-based column index to Excel-style column label."""
    if index < 1:
        raise ValueError("Column index must be positive.")
    chunks: list[str] = []
    current = index
    while current > 0:
        current -= 1
        chunks.append(chr(ord("A") + (current % 26)))
        current //= 26
    return "".join(reversed(chunks))


def range_cell_count(range_ref: str) -> int:
    """Return the number of cells represented by an A1 range."""
    start, end = normalize_range(range_ref).split(":", maxsplit=1)
    start_col, start_row = split_a1(start)
    end_col, end_row = split_a1(end)
    min_col = min(column_label_to_index(start_col), column_label_to_index(end_col))
    max_col = max(column_label_to_index(start_col), column_label_to_index(end_col))
    min_row = min(start_row, end_row)
    max_row = max(start_row, end_row)
    return (max_col - min_col + 1) * (max_row - min_row + 1)


def normalize_range(value: str) -> str:
    """Validate and normalize an A1 range string."""
    candidate = value.strip()
    if not _A1_RANGE_PATTERN.match(candidate):
        raise ValueError(f"Invalid range reference: {value}")
    start, end = candidate.split(":", maxsplit=1)
    return f"{start.upper()}:{end.upper()}"


def parse_qualified_a1_range(value: str) -> QualifiedA1Range:
    """Parse `A1:B2` with optional `Sheet!` qualifier.

    Supported forms:
    - `A1:B2`
    - `Sheet1!A1:B2`
    - `'Sheet 1'!A1:B2`
    """
    candidate = value.strip()
    match = _QUALIFIED_A1_RANGE_PATTERN.fullmatch(candidate)
    if match is None:
        raise ValueError(f"Invalid range reference: {value}")
    sheet_token = match.group("sheet")
    parsed_sheet = _parse_sheet_token(sheet_token) if sheet_token is not None else None
    parsed_range = normalize_range(match.group("range"))
    return QualifiedA1Range(sheet=parsed_sheet, range_ref=parsed_range)


def resolve_sheet_and_range(
    sheet: str | None, range_ref: str | None
) -> SheetRangeSelection:
    """Normalize `sheet`/`range` and validate cross-field consistency."""
    normalized_sheet = _normalize_optional_sheet(sheet)
    if range_ref is None:
        return SheetRangeSelection(sheet=normalized_sheet, range_ref=None)
    parsed = parse_qualified_a1_range(range_ref)
    if normalized_sheet is None:
        if parsed.sheet is None:
            raise ValueError(
                "sheet is required when range is specified and range is unqualified."
            )
        return SheetRangeSelection(sheet=parsed.sheet, range_ref=parsed.range_ref)
    if parsed.sheet is not None and parsed.sheet != normalized_sheet:
        raise ValueError("sheet and range sheet qualifier must match.")
    return SheetRangeSelection(sheet=normalized_sheet, range_ref=parsed.range_ref)


def _normalize_optional_sheet(value: str | None) -> str | None:
    """Normalize optional sheet name and reject blank text."""
    if value is None:
        return None
    candidate = value.strip()
    if not candidate:
        raise ValueError("sheet must not be empty when provided.")
    return candidate


def _parse_sheet_token(token: str) -> str:
    """Parse sheet token from range qualifier."""
    candidate = token.strip()
    if not candidate:
        raise ValueError("Invalid sheet qualifier in range.")
    if candidate.startswith("'") and candidate.endswith("'"):
        inner = candidate[1:-1]
        sheet = inner.replace("''", "'")
        if not sheet:
            raise ValueError("Invalid sheet qualifier in range.")
        return sheet
    return candidate


def parse_range_geometry(range_ref: str) -> tuple[str, int, int]:
    """Parse A1 range and return top-left cell + (rows, cols)."""
    start_ref, end_ref = normalize_range(range_ref).split(":", maxsplit=1)
    start_col, start_row = split_a1(start_ref)
    end_col, end_row = split_a1(end_ref)
    min_col = min(column_label_to_index(start_col), column_label_to_index(end_col))
    max_col = max(column_label_to_index(start_col), column_label_to_index(end_col))
    min_row = min(start_row, end_row)
    max_row = max(start_row, end_row)
    return (
        f"{column_index_to_label(min_col)}{min_row}",
        max_row - min_row + 1,
        max_col - min_col + 1,
    )
