"""Dependency-free Excel coordinate and scalar conversion regressions."""

from datetime import date, datetime, time, timedelta
from pathlib import Path

import pytest

from exstruct.core.ooxml_scalars import (
    MAC_EPOCH,
    WINDOWS_EPOCH,
    UnsupportedOoxmlError,
    coordinate_to_tuple,
    from_excel,
    from_iso8601,
    get_column_letter,
    is_date_format,
    is_timedelta_format,
    range_boundaries,
    translate_formula,
)
from exstruct.core.ooxml_session import OoxmlExtractionSession, _Formula, _SheetData
from exstruct.core.ranges import RangeBounds, parse_range_zero_based


def test_coordinate_and_range_helpers_are_dependency_free() -> None:
    assert coordinate_to_tuple("$AZ$15") == (15, 52)
    assert get_column_letter(16384) == "XFD"
    assert range_boundaries("$B$2:$D$5") == (2, 2, 4, 5)
    assert parse_range_zero_based("'O''Brien, data'!$B$2:$D$5") == RangeBounds(
        r1=1, c1=1, r2=4, c2=3
    )
    assert parse_range_zero_based("Sheet1!1:4") is None


@pytest.mark.parametrize(
    ("serial", "epoch", "expected"),
    [
        (45000, WINDOWS_EPOCH, datetime(2023, 3, 15)),
        (1, MAC_EPOCH, datetime(1904, 1, 2)),
        (0.5, WINDOWS_EPOCH, time(12, 0)),
    ],
)
def test_builtin_date_serial_parity(
    serial: int | float, epoch: datetime, expected: datetime | time
) -> None:
    assert from_excel(serial, epoch=epoch) == expected


def test_date_formats_iso_values_and_elapsed_durations() -> None:
    assert is_date_format("mm-dd-yy")
    assert is_date_format('0 "days";[Red]-0 "days"') is False
    assert is_timedelta_format("[h]:mm:ss")
    assert is_timedelta_format("[H]:MM:SS")
    assert from_iso8601("2026-01-02T03:04:05.120Z") == datetime(
        2026, 1, 2, 3, 4, 5, 120000
    )
    assert from_iso8601("2026-01-02") == date(2026, 1, 2)
    assert from_iso8601("12:30:00") == time(12, 30)
    assert from_iso8601("PT2H3M4.5S") == timedelta(hours=2, minutes=3, seconds=4.5)


def test_shared_formula_translation_ignores_quoted_names_and_strings() -> None:
    assert (
        translate_formula('=IF("A1"="A1",\'A1\'!B1+$C$2,0)', origin="B2", target="C3")
        == '=IF("A1"="A1",\'A1\'!C2+$C$2,0)'
    )
    with pytest.raises(UnsupportedOoxmlError, match="Whole-row/column"):
        translate_formula("=SUM(1:2)", origin="A1", target="A2")


def test_unsupported_shared_formula_translation_propagates() -> None:
    data = _SheetData(
        formulas=[
            _Formula("A2", "SUM(B:B)", "shared", "0"),
            _Formula("A3", None, "shared", "0"),
        ]
    )
    session = OoxmlExtractionSession(Path("unused.xlsx"))
    with pytest.raises(UnsupportedOoxmlError, match="Whole-row/column"):
        session._formulas_map("Sheet", data)


@pytest.mark.parametrize("formula", ["=A1!B1", "=A1:B2!C3", "=項目A1+B1"])
def test_shared_formula_qualifiers_and_names_request_compatibility(
    formula: str,
) -> None:
    with pytest.raises(UnsupportedOoxmlError):
        translate_formula(formula, origin="A1", target="A2")


def test_shared_formula_out_of_bounds_requests_compatibility() -> None:
    with pytest.raises(UnsupportedOoxmlError, match="bounds"):
        translate_formula("=A1", origin="B1", target="A2")
