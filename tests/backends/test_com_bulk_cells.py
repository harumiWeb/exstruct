from __future__ import annotations

from collections.abc import Sequence

import pytest

from exstruct.core.backends.com_backend import ComBackend
from exstruct.models import CellRow


class _Count:
    def __init__(self, count: int) -> None:
        self.Count = count


class _UsedRange:
    def __init__(self, row: int, column: int, height: int, width: int) -> None:
        self.Row = row
        self.Column = column
        self.Rows = _Count(height)
        self.Columns = _Count(width)


class _Anchor:
    def __init__(self, row: int, column: int, height: int = 1, width: int = 1) -> None:
        self.Row = row
        self.Column = column
        self.Rows = _Count(height)
        self.Columns = _Count(width)


class _Hyperlink:
    def __init__(self, address: str, anchor: _Anchor, subaddress: str = "") -> None:
        self.Address = address
        self.Range = anchor
        self.SubAddress = subaddress


class _FakeRange:
    def __init__(
        self,
        sheet: _FakeSheet,
        first_row: int,
        first_column: int,
        last_row: int,
        last_column: int,
    ) -> None:
        self.sheet = sheet
        self.first_row = first_row
        self.first_column = first_column
        self.last_row = last_row
        self.last_column = last_column
        self.Address = f"R{first_row}C{first_column}:R{last_row}C{last_column}"
        self.api = self

    @property
    def Value2(self) -> object:
        if self.sheet.value_error is not None:
            raise self.sheet.value_error
        values = self._slice(self.sheet.values)
        if len(values) == 1 and len(values[0]) == 1:
            return values[0][0]
        return tuple(tuple(row) for row in values)

    def error_values(self) -> tuple[tuple[bool, ...], ...]:
        return tuple(tuple(row) for row in self._slice(self.sheet.errors))

    def _slice(self, source: Sequence[Sequence[object]]) -> list[list[object]]:
        row_offset = self.first_row - self.sheet.first_row
        column_offset = self.first_column - self.sheet.first_column
        height = self.last_row - self.first_row + 1
        width = self.last_column - self.first_column + 1
        return [
            list(source[row][column_offset : column_offset + width])
            for row in range(row_offset, row_offset + height)
        ]


class _SheetApi:
    def __init__(self, sheet: _FakeSheet) -> None:
        self.sheet = sheet
        self.UsedRange = _UsedRange(
            sheet.first_row,
            sheet.first_column,
            len(sheet.values),
            len(sheet.values[0]),
        )
        self.Hyperlinks = sheet.hyperlinks

    def Evaluate(self, formula: str) -> object:
        if self.sheet.evaluate_error is not None:
            raise self.sheet.evaluate_error
        expected = f"ISERROR({self.sheet.last_range.Address})"
        assert formula == expected
        errors = self.sheet.last_range.error_values()
        if len(errors) == 1 and len(errors[0]) == 1:
            return errors[0][0]
        return errors


class _FakeSheet:
    def __init__(
        self,
        values: list[list[object]],
        *,
        first_row: int = 1,
        first_column: int = 1,
        errors: list[list[bool]] | None = None,
        hyperlinks: list[_Hyperlink] | None = None,
        value_error: Exception | None = None,
        evaluate_error: Exception | None = None,
        name: str = "Data",
    ) -> None:
        self.name = name
        self.values = values
        self.first_row = first_row
        self.first_column = first_column
        self.errors = errors or [[False for _ in row] for row in values]
        self.hyperlinks = hyperlinks or []
        self.value_error = value_error
        self.evaluate_error = evaluate_error
        self.last_range: _FakeRange
        self.api = _SheetApi(self)

    def range(self, first: tuple[int, int], last: tuple[int, int]) -> _FakeRange:
        self.last_range = _FakeRange(self, *first, *last)
        return self.last_range


class _FakeWorkbook:
    def __init__(self, *sheets: _FakeSheet) -> None:
        self.sheets = list(sheets)


def test_com_extract_cells_preserves_offsets_and_normalizes_values() -> None:
    error_code = 2042
    sheet = _FakeSheet(
        [
            [error_code, 12.0, "42", True, None, "NA"],
            [error_code, 2.75, False, "hello", "N/A", "  "],
        ],
        first_row=7,
        first_column=4,
        errors=[
            [False, False, False, False, False, False],
            [True, False, False, False, False, False],
        ],
    )

    rows = ComBackend(_FakeWorkbook(sheet)).extract_cells()["Data"]

    assert rows == [
        CellRow(
            r=7,
            c={"3": error_code, "4": 12, "5": 42, "6": "True"},
        ),
        CellRow(r=8, c={"4": 2.75, "5": "False", "6": "hello"}),
    ]


@pytest.mark.parametrize(
    ("values", "first_row", "first_column", "expected"),
    [
        ([[3.0]], 9, 6, [CellRow(r=9, c={"5": 3})]),
        ([["left", "right"]], 4, 3, [CellRow(r=4, c={"2": "left", "3": "right"})]),
        (
            [[1], [2]],
            5,
            2,
            [CellRow(r=5, c={"1": 1}), CellRow(r=6, c={"1": 2})],
        ),
    ],
    ids=["scalar", "one-row", "one-column"],
)
def test_com_extract_cells_handles_com_array_shapes(
    values: list[list[object]],
    first_row: int,
    first_column: int,
    expected: list[CellRow],
) -> None:
    sheet = _FakeSheet(values, first_row=first_row, first_column=first_column)

    assert ComBackend(_FakeWorkbook(sheet)).extract_cells()["Data"] == expected


def test_com_extract_cells_masks_errors_without_dropping_equal_integer() -> None:
    error_code = 2042
    sheet = _FakeSheet(
        [[error_code, error_code]],
        errors=[[False, True]],
    )

    rows = ComBackend(_FakeWorkbook(sheet)).extract_cells()["Data"]

    assert rows == [CellRow(r=1, c={"0": error_code})]


def test_com_extract_cells_includes_external_links_but_omits_internal_and_blank_rows() -> (
    None
):
    sheet = _FakeSheet(
        [["linked", None], [None, None], [None, "internal"]],
        first_row=10,
        first_column=3,
        hyperlinks=[
            _Hyperlink("https://example.test/page", _Anchor(10, 4)),
            _Hyperlink("https://blank.test/", _Anchor(11, 3)),
            _Hyperlink("", _Anchor(12, 3), subaddress="Other!A1"),
        ],
    )

    rows = ComBackend(_FakeWorkbook(sheet)).extract_cells(include_links=True)["Data"]

    assert rows == [
        CellRow(
            r=10,
            c={"2": "linked"},
            links={"3": "https://example.test/page"},
        ),
        CellRow(r=12, c={"3": "internal"}),
    ]


@pytest.mark.parametrize("failure_point", ["Value2", "Evaluate"])
def test_com_extract_cells_propagates_com_read_errors(failure_point: str) -> None:
    failure = RuntimeError(f"{failure_point} failed")
    sheet = _FakeSheet(
        [["value"]],
        value_error=failure if failure_point == "Value2" else None,
        evaluate_error=failure if failure_point == "Evaluate" else None,
    )

    with pytest.raises(RuntimeError, match=f"{failure_point} failed"):
        ComBackend(_FakeWorkbook(sheet)).extract_cells()


def test_com_extract_cells_batches_large_rectangles() -> None:
    sheet = _FakeSheet([[None] * 2 for _ in range(50_001)])
    sheet.values[0][0] = "first"
    sheet.values[-1][1] = "last"
    assert ComBackend(_FakeWorkbook(sheet)).extract_cells()["Data"] == [
        CellRow(r=1, c={"0": "first"}),
        CellRow(r=50_001, c={"1": "last"}),
    ]
    assert sheet.last_range.first_row == 50_001
    assert sheet.last_range.last_row == 50_001
