from __future__ import annotations

import json
from pathlib import Path
import re
from typing import Any, TypeAlias, cast

from pydantic import BaseModel, Field

from .io import PathPolicy

JsonScalar: TypeAlias = str | int | float | bool | None
CellCoordinate: TypeAlias = tuple[int, int]
RangeCoordinate: TypeAlias = tuple[int, int, int, int]

_CELL_RE = re.compile(r"^([A-Za-z]+)([1-9][0-9]*)$")


class CellReadItem(BaseModel):
    """Cell read result item."""

    cell: str
    value: JsonScalar = None
    formula: str | None = None


class FormulaReadItem(BaseModel):
    """Formula read result item."""

    cell: str
    formula: str
    value: JsonScalar = None


class ReadRangeRequest(BaseModel):
    """Input model for range reading."""

    out_path: Path
    sheet: str | None = None
    range: str
    include_formulas: bool = False
    include_empty: bool = True
    max_cells: int = Field(default=10_000, ge=1)


class ReadRangeResult(BaseModel):
    """Output model for range reading."""

    book_name: str | None = None
    sheet_name: str
    range: str
    cells: list[CellReadItem] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ReadCellsRequest(BaseModel):
    """Input model for cell list reading."""

    out_path: Path
    sheet: str | None = None
    addresses: list[str] = Field(min_length=1)
    include_formulas: bool = True


class ReadCellsResult(BaseModel):
    """Output model for cell list reading."""

    book_name: str | None = None
    sheet_name: str
    cells: list[CellReadItem] = Field(default_factory=list)
    missing_cells: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ReadFormulasRequest(BaseModel):
    """Input model for formula reading."""

    out_path: Path
    sheet: str | None = None
    range: str | None = None
    include_values: bool = False


class ReadFormulasResult(BaseModel):
    """Output model for formula reading."""

    book_name: str | None = None
    sheet_name: str
    range: str | None = None
    formulas: list[FormulaReadItem] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


def read_range(
    request: ReadRangeRequest, *, policy: PathPolicy | None = None
) -> ReadRangeResult:
    """Read a rectangular range from extracted JSON.

    Args:
        request: Range read request.
        policy: Optional path policy for access control.

    Returns:
        Range read result.
    """
    resolved = _resolve_output_path(request.out_path, policy=policy)
    data = _parse_json(_read_text(resolved))
    sheet_name, sheet_data = _select_sheet(data, request.sheet)
    start_row, start_col, end_row, end_col = _parse_range(request.range)
    normalized_range = _format_range(start_row, start_col, end_row, end_col)
    cell_count = (end_row - start_row + 1) * (end_col - start_col + 1)
    if cell_count > request.max_cells:
        raise ValueError(
            "Requested range exceeds max_cells. "
            f"range={normalized_range}, cells={cell_count}, max_cells={request.max_cells}"
        )

    value_map = _build_value_map(sheet_data)
    formula_map, has_formula_map = _build_formula_map(sheet_data)
    warnings: list[str] = []
    if request.include_formulas and not has_formula_map:
        warnings.append(
            "formulas_map is not available. Re-run exstruct_extract with mode='verbose' "
            "to inspect formulas."
        )

    items: list[CellReadItem] = []
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            coord = (row, col)
            value = value_map.get(coord)
            formula = formula_map.get(coord) if request.include_formulas else None
            if not request.include_empty and value is None and formula is None:
                continue
            items.append(
                CellReadItem(
                    cell=_format_cell(row, col),
                    value=value,
                    formula=formula,
                )
            )

    return ReadRangeResult(
        book_name=_as_optional_str(data.get("book_name")),
        sheet_name=sheet_name,
        range=normalized_range,
        cells=items,
        warnings=warnings,
    )


def read_cells(
    request: ReadCellsRequest, *, policy: PathPolicy | None = None
) -> ReadCellsResult:
    """Read specific cells from extracted JSON.

    Args:
        request: Cell list read request.
        policy: Optional path policy for access control.

    Returns:
        Cell list read result.
    """
    resolved = _resolve_output_path(request.out_path, policy=policy)
    data = _parse_json(_read_text(resolved))
    sheet_name, sheet_data = _select_sheet(data, request.sheet)
    value_map = _build_value_map(sheet_data)
    formula_map, has_formula_map = _build_formula_map(sheet_data)

    warnings: list[str] = []
    if request.include_formulas and not has_formula_map:
        warnings.append(
            "formulas_map is not available. Re-run exstruct_extract with mode='verbose' "
            "to inspect formulas."
        )

    items: list[CellReadItem] = []
    missing_cells: list[str] = []
    for address in request.addresses:
        row, col = _parse_cell(address)
        normalized = _format_cell(row, col)
        coord = (row, col)
        value = value_map.get(coord)
        formula = formula_map.get(coord) if request.include_formulas else None
        if value is None and formula is None:
            missing_cells.append(normalized)
        items.append(CellReadItem(cell=normalized, value=value, formula=formula))

    return ReadCellsResult(
        book_name=_as_optional_str(data.get("book_name")),
        sheet_name=sheet_name,
        cells=items,
        missing_cells=missing_cells,
        warnings=warnings,
    )


def read_formulas(
    request: ReadFormulasRequest, *, policy: PathPolicy | None = None
) -> ReadFormulasResult:
    """Read formulas from extracted JSON.

    Args:
        request: Formula read request.
        policy: Optional path policy for access control.

    Returns:
        Formula read result.
    """
    resolved = _resolve_output_path(request.out_path, policy=policy)
    data = _parse_json(_read_text(resolved))
    sheet_name, sheet_data = _select_sheet(data, request.sheet)
    value_map = _build_value_map(sheet_data)
    formula_map, has_formula_map = _build_formula_map(sheet_data)

    range_filter: RangeCoordinate | None = None
    normalized_range: str | None = None
    if request.range is not None:
        range_filter = _parse_range(request.range)
        normalized_range = _format_range(*range_filter)

    warnings: list[str] = []
    if not has_formula_map:
        warnings.append(
            "formulas_map is not available. Re-run exstruct_extract with mode='verbose' "
            "to inspect formulas."
        )
        return ReadFormulasResult(
            book_name=_as_optional_str(data.get("book_name")),
            sheet_name=sheet_name,
            range=normalized_range,
            formulas=[],
            warnings=warnings,
        )

    items: list[FormulaReadItem] = []
    for row, col in sorted(formula_map):
        if range_filter is not None and not _in_range((row, col), range_filter):
            continue
        coord = (row, col)
        items.append(
            FormulaReadItem(
                cell=_format_cell(row, col),
                formula=formula_map[coord],
                value=value_map.get(coord) if request.include_values else None,
            )
        )

    return ReadFormulasResult(
        book_name=_as_optional_str(data.get("book_name")),
        sheet_name=sheet_name,
        range=normalized_range,
        formulas=items,
        warnings=warnings,
    )


def _resolve_output_path(path: Path, *, policy: PathPolicy | None) -> Path:
    """Resolve and validate output path."""
    resolved = policy.ensure_allowed(path) if policy else path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"Output file not found: {resolved}")
    if not resolved.is_file():
        raise ValueError(f"Output path is not a file: {resolved}")
    return resolved


def _read_text(path: Path) -> str:
    """Read UTF-8 text from file."""
    return path.read_text(encoding="utf-8")


def _parse_json(text: str) -> dict[str, Any]:
    """Parse JSON text into object root."""
    parsed = json.loads(text)
    if not isinstance(parsed, dict):
        raise ValueError("Invalid workbook JSON: expected object at root.")
    return cast(dict[str, Any], parsed)


def _select_sheet(
    data: dict[str, Any], sheet: str | None
) -> tuple[str, dict[str, Any]]:
    """Select sheet payload from workbook object."""
    sheets = data.get("sheets")
    if not isinstance(sheets, dict):
        raise ValueError("Invalid workbook JSON: sheets is not a mapping.")
    if sheet is not None:
        selected = sheets.get(sheet)
        if not isinstance(selected, dict):
            raise ValueError(f"Sheet not found: {sheet}")
        return sheet, selected
    if len(sheets) == 1:
        only_name = next(iter(sheets.keys()))
        only_payload = sheets[only_name]
        if not isinstance(only_payload, dict):
            raise ValueError("Invalid workbook JSON: sheet payload is not an object.")
        return str(only_name), only_payload
    names = ", ".join(str(name) for name in sheets.keys())
    raise ValueError(
        "Sheet is required when multiple sheets exist. "
        f"Available sheets: {names}. "
        "Retry with `sheet='<name>'`."
    )


def _parse_cell(address: str) -> CellCoordinate:
    """Parse A1 cell address into row/column tuple (1-based)."""
    match = _CELL_RE.fullmatch(address.strip())
    if match is None:
        raise ValueError(f"Invalid A1 cell address: {address}")
    col_label, row_text = match.groups()
    row = int(row_text)
    col = _alpha_to_col(col_label)
    return row, col


def _parse_range(range_text: str) -> RangeCoordinate:
    """Parse A1 range text into start/end tuple."""
    cleaned = range_text.strip()
    if ":" in cleaned:
        left, right = cleaned.split(":", maxsplit=1)
        start_row, start_col = _parse_cell(left)
        end_row, end_col = _parse_cell(right)
    else:
        start_row, start_col = _parse_cell(cleaned)
        end_row, end_col = start_row, start_col
    if start_row > end_row or start_col > end_col:
        raise ValueError(f"Invalid A1 range: {range_text}")
    return start_row, start_col, end_row, end_col


def _build_value_map(sheet_data: dict[str, Any]) -> dict[CellCoordinate, JsonScalar]:
    """Build coordinate->value map from sheet rows."""
    rows = sheet_data.get("rows")
    if not isinstance(rows, list):
        return {}
    value_map: dict[CellCoordinate, JsonScalar] = {}
    for row_payload in rows:
        if not isinstance(row_payload, dict):
            continue
        row_index = row_payload.get("r")
        if not isinstance(row_index, int) or row_index <= 0:
            continue
        cols = row_payload.get("c")
        if not isinstance(cols, dict):
            continue
        for key, raw_value in cols.items():
            if not isinstance(key, str):
                continue
            col_index = _parse_col_key(key)
            if col_index is None:
                continue
            value = _normalize_scalar(raw_value)
            value_map[(row_index, col_index)] = value
    return value_map


def _build_formula_map(
    sheet_data: dict[str, Any],
) -> tuple[dict[CellCoordinate, str], bool]:
    """Build coordinate->formula map from formulas_map."""
    formulas_raw = sheet_data.get("formulas_map")
    if formulas_raw is None:
        return {}, False
    if not isinstance(formulas_raw, dict):
        return {}, False
    formula_map: dict[CellCoordinate, str] = {}
    for formula, positions in formulas_raw.items():
        if not isinstance(formula, str):
            continue
        if not isinstance(positions, list):
            continue
        for position in positions:
            parsed = _parse_formula_position(position)
            if parsed is None:
                continue
            formula_map[parsed] = formula
    return formula_map, True


def _parse_formula_position(position: object) -> CellCoordinate | None:
    """Parse formulas_map position into row/column tuple."""
    if not isinstance(position, list | tuple):
        return None
    if len(position) != 2:
        return None
    row = position[0]
    col = position[1]
    if not isinstance(row, int) or not isinstance(col, int):
        return None
    if row <= 0 or col < 0:
        return None
    return row, col + 1


def _parse_col_key(key: str) -> int | None:
    """Parse cell column key into 1-based column index."""
    stripped = key.strip()
    try:
        zero_based = int(stripped)
    except (TypeError, ValueError):
        pass
    else:
        if zero_based < 0:
            return None
        return zero_based + 1
    normalized = stripped.upper()
    if not normalized.isalpha():
        return None
    return _alpha_to_col(normalized)


def _alpha_to_col(label: str) -> int:
    """Convert alphabetic column label to 1-based index."""
    acc = 0
    for char in label.strip().upper():
        code = ord(char)
        if code < ord("A") or code > ord("Z"):
            raise ValueError(f"Invalid column label: {label}")
        acc = acc * 26 + (code - ord("A") + 1)
    return acc


def _col_to_alpha(col_index: int) -> str:
    """Convert 1-based column index to alphabetic column label."""
    if col_index <= 0:
        raise ValueError(f"Invalid column index: {col_index}")
    current = col_index
    chars: list[str] = []
    while current > 0:
        current -= 1
        chars.append(chr(ord("A") + (current % 26)))
        current //= 26
    return "".join(reversed(chars))


def _format_cell(row: int, col: int) -> str:
    """Format row/column as A1 cell address."""
    return f"{_col_to_alpha(col)}{row}"


def _format_range(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """Format start/end coordinates as A1 range."""
    start = _format_cell(start_row, start_col)
    end = _format_cell(end_row, end_col)
    if start == end:
        return start
    return f"{start}:{end}"


def _in_range(cell: CellCoordinate, bounds: RangeCoordinate) -> bool:
    """Check whether cell is included in range bounds."""
    row, col = cell
    start_row, start_col, end_row, end_col = bounds
    return start_row <= row <= end_row and start_col <= col <= end_col


def _normalize_scalar(value: object) -> JsonScalar:
    """Normalize arbitrary JSON value to scalar."""
    if isinstance(value, str | int | float | bool) or value is None:
        return value
    return str(value)


def _as_optional_str(value: object) -> str | None:
    """Convert object to optional string."""
    if value is None:
        return None
    if isinstance(value, str):
        return value
    return str(value)
