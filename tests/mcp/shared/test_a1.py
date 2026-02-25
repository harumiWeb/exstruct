from __future__ import annotations

import pytest

from exstruct.mcp.shared.a1 import (
    column_index_to_label,
    column_label_to_index,
    parse_range_geometry,
    range_cell_count,
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
