"""Tests for alpha column key conversion utilities."""

from __future__ import annotations

import pytest

from exstruct.models import (
    CellRow,
    MergedCells,
    SheetData,
    WorkbookData,
    col_index_to_alpha,
    convert_row_keys_to_alpha,
    convert_sheet_keys_to_alpha,
    convert_workbook_keys_to_alpha,
)


class TestColIndexToAlpha:
    """Tests for col_index_to_alpha()."""

    def test_known_values(self) -> None:
        cases: list[tuple[int, str]] = [
            (0, "A"),
            (1, "B"),
            (25, "Z"),
            (26, "AA"),
            (27, "AB"),
            (51, "AZ"),
            (52, "BA"),
            (701, "ZZ"),
            (702, "AAA"),
        ]
        for index, expected in cases:
            assert col_index_to_alpha(index) == expected

    def test_negative_raises(self) -> None:
        with pytest.raises(ValueError, match="non-negative"):
            col_index_to_alpha(-1)


class TestConvertRowKeysToAlpha:
    """Tests for convert_row_keys_to_alpha()."""

    def test_basic_numeric_keys(self) -> None:
        row = CellRow(r=1, c={"0": "val_A", "1": "val_B", "2": "val_C"}, links=None)
        result = convert_row_keys_to_alpha(row)
        assert result.c == {"A": "val_A", "B": "val_B", "C": "val_C"}
        assert result.r == 1

    def test_links_converted(self) -> None:
        row = CellRow(
            r=2,
            c={"0": "text"},
            links={"0": "http://example.com"},
        )
        result = convert_row_keys_to_alpha(row)
        assert result.links == {"A": "http://example.com"}

    def test_no_links(self) -> None:
        row = CellRow(r=1, c={"0": 1}, links=None)
        result = convert_row_keys_to_alpha(row)
        assert result.links is None

    def test_non_numeric_keys_pass_through(self) -> None:
        row = CellRow(r=1, c={"X": "val"}, links=None)
        result = convert_row_keys_to_alpha(row)
        assert result.c == {"X": "val"}

    def test_high_columns(self) -> None:
        row = CellRow(r=1, c={"26": "AA_col"}, links=None)
        result = convert_row_keys_to_alpha(row)
        assert result.c == {"AA": "AA_col"}

    def test_raises_on_collision_in_cells(self) -> None:
        row = CellRow(r=3, c={"0": "num_zero", "A": "alpha_a"}, links=None)
        with pytest.raises(ValueError, match="collision"):
            convert_row_keys_to_alpha(row)

    def test_raises_on_collision_in_links(self) -> None:
        row = CellRow(
            r=4,
            c={"0": "value"},
            links={"0": "https://numeric.example", "A": "https://alpha.example"},
        )
        with pytest.raises(ValueError, match="collision"):
            convert_row_keys_to_alpha(row)


class TestConvertSheetKeysToAlpha:
    """Tests for convert_sheet_keys_to_alpha()."""

    def test_converts_all_rows(self) -> None:
        rows = [
            CellRow(r=1, c={"0": "A1", "1": "B1"}, links=None),
            CellRow(r=2, c={"0": "A2", "2": "C2"}, links=None),
        ]
        sheet = SheetData(rows=rows)
        result = convert_sheet_keys_to_alpha(sheet)
        assert result.rows[0].c == {"A": "A1", "B": "B1"}
        assert result.rows[1].c == {"A": "A2", "C": "C2"}

    def test_preserves_other_fields(self) -> None:
        sheet = SheetData(
            rows=[CellRow(r=1, c={"0": "v"}, links=None)],
            shapes=[],
            charts=[],
        )
        result = convert_sheet_keys_to_alpha(sheet)
        assert result.shapes == []
        assert result.charts == []

    def test_moves_merged_cells_to_merged_ranges(self) -> None:
        sheet = SheetData(
            rows=[CellRow(r=1, c={"0": "v"}, links=None)],
            merged_cells=MergedCells(
                items=[(1, 0, 3, 2, "merged"), (10, 25, 10, 26, "x")]
            ),
        )
        result = convert_sheet_keys_to_alpha(sheet)
        assert result.merged_ranges == ["A1:C3", "Z10:AA10"]
        assert result.merged_cells is None

    def test_no_merged_cells_keeps_none(self) -> None:
        sheet = SheetData(
            rows=[CellRow(r=1, c={"0": "v"}, links=None)], merged_cells=None
        )
        result = convert_sheet_keys_to_alpha(sheet)
        assert result.merged_cells is None
        assert result.merged_ranges == []


class TestConvertWorkbookKeysToAlpha:
    """Tests for convert_workbook_keys_to_alpha()."""

    def test_converts_all_sheets(self) -> None:
        wb = WorkbookData(
            book_name="test.xlsx",
            sheets={
                "Sheet1": SheetData(
                    rows=[CellRow(r=1, c={"0": "a", "1": "b"}, links=None)]
                ),
                "Sheet2": SheetData(rows=[CellRow(r=1, c={"2": "c"}, links=None)]),
            },
        )
        result = convert_workbook_keys_to_alpha(wb)
        assert result.sheets["Sheet1"].rows[0].c == {"A": "a", "B": "b"}
        assert result.sheets["Sheet2"].rows[0].c == {"C": "c"}
        assert result.book_name == "test.xlsx"

    def test_empty_workbook(self) -> None:
        wb = WorkbookData(book_name="empty.xlsx", sheets={})
        result = convert_workbook_keys_to_alpha(wb)
        assert result.sheets == {}
