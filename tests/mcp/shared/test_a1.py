from __future__ import annotations

import pytest

from exstruct.mcp.shared.a1 import (
    column_index_to_label,
    column_label_to_index,
    parse_qualified_a1_range,
    parse_range_geometry,
    range_cell_count,
    resolve_sheet_and_range,
    split_a1,
)


def test_column_roundtrip() -> None:
    assert column_label_to_index("A") == 1
    assert column_label_to_index("AA") == 27
    assert column_index_to_label(1) == "A"
    assert column_index_to_label(27) == "AA"


def test_split_a1() -> None:
    assert split_a1("b12") == ("B", 12)


def test_range_cell_count() -> None:
    assert range_cell_count("A1:C3") == 9
    assert range_cell_count("C3:A1") == 9


def test_parse_range_geometry() -> None:
    base, rows, cols = parse_range_geometry("D6:B4")
    assert base == "B4"
    assert rows == 3
    assert cols == 3


def test_split_a1_rejects_invalid() -> None:
    with pytest.raises(ValueError, match="Invalid cell reference"):
        split_a1("1A")


def test_parse_qualified_a1_range_supports_sheet_qualifier() -> None:
    parsed = parse_qualified_a1_range("'Sheet 1'!a1:b2")
    assert parsed.sheet == "Sheet 1"
    assert parsed.range_ref == "A1:B2"


def test_resolve_sheet_and_range_rejects_missing_sheet() -> None:
    with pytest.raises(ValueError, match="sheet is required"):
        resolve_sheet_and_range(None, "A1:B2")


def test_resolve_sheet_and_range_rejects_sheet_mismatch() -> None:
    with pytest.raises(ValueError, match="must match"):
        resolve_sheet_and_range("Summary", "Sheet1!A1:B2")
