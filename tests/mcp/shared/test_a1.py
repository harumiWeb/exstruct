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
    """Convert between column labels and indices in both directions."""
    assert column_label_to_index("A") == 1
    assert column_label_to_index("AA") == 27
    assert column_index_to_label(1) == "A"
    assert column_index_to_label(27) == "AA"


def test_split_a1() -> None:
    """Normalize cell references to uppercase column labels."""
    assert split_a1("b12") == ("B", 12)


def test_range_cell_count() -> None:
    """Count cells in ranges regardless of start/end ordering."""
    assert range_cell_count("A1:C3") == 9
    assert range_cell_count("C3:A1") == 9


def test_parse_range_geometry() -> None:
    """Return top-left anchor with row and column span."""
    base, rows, cols = parse_range_geometry("D6:B4")
    assert base == "B4"
    assert rows == 3
    assert cols == 3


def test_split_a1_rejects_invalid() -> None:
    """Reject malformed cell references."""
    with pytest.raises(ValueError, match="Invalid cell reference"):
        split_a1("1A")


def test_parse_qualified_a1_range_supports_sheet_qualifier() -> None:
    """Parse quoted sheet qualifiers and normalize range case."""
    parsed = parse_qualified_a1_range("'Sheet 1'!a1:b2")
    assert parsed.sheet == "Sheet 1"
    assert parsed.range_ref == "A1:B2"


def test_resolve_sheet_and_range_accepts_qualified_range_without_sheet() -> None:
    """Accept qualified ranges when top-level sheet is omitted."""
    selection = resolve_sheet_and_range(None, "Sheet1!A1:B2")
    assert selection.sheet == "Sheet1"
    assert selection.range_ref == "A1:B2"


def test_resolve_sheet_and_range_rejects_unqualified_range_without_sheet() -> None:
    """Reject unqualified ranges when top-level sheet is omitted."""
    with pytest.raises(ValueError, match="sheet is required"):
        resolve_sheet_and_range(None, "A1:B2")


def test_resolve_sheet_and_range_rejects_sheet_mismatch() -> None:
    """Reject inconsistent sheet names across `sheet` and `range`."""
    with pytest.raises(ValueError, match="must match"):
        resolve_sheet_and_range("Summary", "Sheet1!A1:B2")
