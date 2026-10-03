"""Pure-Python scalar helpers for reading common OOXML workbook values."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta as duration
import re


class UnsupportedOoxmlError(Exception):
    """Signal that an OOXML construct needs the compatibility backend."""


_MAX_COLUMN = 18278  # openpyxl's accepted three-letter column boundary
_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")
_COLUMN_RE = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW_RE = re.compile(r"^\$?([1-9][0-9]*)$")
_FORMULA_CELL_RE = re.compile(
    r"(?<![A-Za-z0-9_.\[\]])(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)(?![A-Za-z0-9_.\[(\]])"
)
_WHOLE_RANGE_RE = re.compile(
    r"(?<![A-Za-z0-9_.])\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|(?<![A-Za-z0-9_.])\$?[1-9][0-9]*:\$?[1-9][0-9]*"
)
_ISO_RE = re.compile(
    r"(?P<date>(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2}))?T?"
    r"(?P<time>(?P<hour>\d{2}):(?P<minute>\d{2})"
    r"(?::(?P<second>\d{2})(?P<microsecond>\.\d{1,3})?)?)?Z?"
)
_ISO_DURATION_RE = re.compile(
    r"PT(?:(?P<hours>\d+)H)?(?:(?P<minutes>\d+)M)?"
    r"(?:(?P<seconds>\d+(?:\.\d{1,3})?)S)?"
)
_FORMAT_STRIP_RE = re.compile(r'".*?"|\[(?!hh?\]|mm?\]|ss?\])[^\]]*\]')
_TIMEDELTA_RE = re.compile(
    r"\[hh?\](:mm(:ss(\.0*)?)?)?|\[mm?\](:ss(\.0*)?)?|\[ss?\](\.0*)?",
    re.IGNORECASE,
)
_UNQUOTED_SHEET_RE = re.compile(r"[\w.]+(?::[\w.]+)?!")

BUILTIN_FORMATS: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: '"$"#,##0_);("$"#,##0)',
    6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)',
    8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0_);(#,##0)",
    38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

MAC_EPOCH = datetime(1904, 1, 1)
WINDOWS_EPOCH = datetime(1899, 12, 30)


def _column_index(letters: str) -> int:
    """Convert one to three ASCII letters to a one-based column number."""
    index = 0
    for letter in letters.upper():
        index = index * 26 + ord(letter) - ord("A") + 1
    if not 1 <= index <= _MAX_COLUMN:
        raise ValueError(f"Invalid column index {letters!r}")
    return index


def coordinate_to_tuple(coordinate: str) -> tuple[int, int]:
    """Return a cell coordinate as one-based ``(row, column)`` numbers."""
    match = _CELL_RE.fullmatch(coordinate)
    if match is None:
        raise ValueError(f"Invalid cell coordinate {coordinate!r}")
    return int(match.group(2)), _column_index(match.group(1))


def get_column_letter(index: int) -> str:
    """Convert a one-based column number into its Excel letter label."""
    if not isinstance(index, int) or not 1 <= index <= _MAX_COLUMN:
        raise ValueError(f"Invalid column index {index!r}")
    letters: list[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def _range_endpoint(
    text: str,
) -> tuple[int | None, int | None, str]:
    """Parse one cell, column, or row endpoint."""
    match = _CELL_RE.fullmatch(text)
    if match is not None:
        return _column_index(match.group(1)), int(match.group(2)), "cell"
    match = _COLUMN_RE.fullmatch(text)
    if match is not None:
        return _column_index(match.group(1)), None, "column"
    match = _ROW_RE.fullmatch(text)
    if match is not None:
        return None, int(match.group(1)), "row"
    raise ValueError(f"Invalid cell range endpoint {text!r}")


def range_boundaries(
    reference: str,
) -> tuple[int | None, int | None, int | None, int | None]:
    """Return one-based ``(min_col, min_row, max_col, max_row)`` bounds."""
    endpoints = reference.split(":")
    if len(endpoints) not in (1, 2):
        raise ValueError(f"Invalid cell range {reference!r}")
    start = _range_endpoint(endpoints[0])
    end = _range_endpoint(endpoints[-1])
    if start[2] != end[2] and len(endpoints) == 2:
        raise ValueError(f"Mismatched cell range endpoints {reference!r}")
    return start[0], start[1], end[0], end[1]


def is_date_format(format_code: str | None) -> bool:
    """Return whether the first number-format section represents a date/time."""
    if format_code is None:
        return False
    first_section = _FORMAT_STRIP_RE.sub("", format_code.split(";", 1)[0])
    return re.search(r"(?<![_\\])[dmhysDMHYS]", first_section) is not None


def is_timedelta_format(format_code: str | None) -> bool:
    """Return whether a format uses elapsed hours, minutes, or seconds."""
    if format_code is None:
        return False
    return _TIMEDELTA_RE.search(format_code.split(";", 1)[0]) is not None


def from_excel(
    value: int | float,
    epoch: datetime = WINDOWS_EPOCH,
    timedelta: bool = False,
) -> datetime | date | time | duration:
    """Convert an Excel serial into a date, time, or elapsed duration."""
    if timedelta:
        return duration(milliseconds=round(value * 86_400_000))
    day, fraction = divmod(value, 1)
    diff = duration(milliseconds=round(fraction * 86_400_000))
    if 0 <= value < 1 and diff.days == 0:
        return time(
            diff.seconds // 3600,
            (diff.seconds % 3600) // 60,
            diff.seconds % 60,
            diff.microseconds,
        )
    if 0 < value < 60 and epoch == WINDOWS_EPOCH:
        day += 1
    return epoch + duration(days=day) + diff


def from_iso8601(value: str) -> datetime | date | time | duration:
    """Decode ISO date/time values and OOXML's ISO elapsed-time form."""
    if not value:
        raise ValueError("Invalid datetime value")
    match = _ISO_RE.fullmatch(value)
    if match is not None and (match.group("date") or match.group("time")):
        parts = match.groupdict()
        year = int(parts["year"]) if parts["year"] else None
        month = int(parts["month"]) if parts["month"] else None
        day = int(parts["day"]) if parts["day"] else None
        hour = int(parts["hour"]) if parts["hour"] else 0
        minute = int(parts["minute"]) if parts["minute"] else 0
        second = int(parts["second"]) if parts["second"] else 0
        fraction = parts["microsecond"]
        microsecond = int(float(fraction) * 1_000_000) if fraction else 0
        if parts["date"] and parts["time"]:
            assert year is not None and month is not None and day is not None
            return datetime(year, month, day, hour, minute, second, microsecond)
        if parts["date"]:
            assert year is not None and month is not None and day is not None
            return date(year, month, day)
        return time(hour, minute, second, microsecond)
    duration_match = _ISO_DURATION_RE.fullmatch(value)
    if duration_match is not None and any(duration_match.groupdict().values()):
        parts = duration_match.groupdict()
        return duration(
            hours=float(parts["hours"] or 0),
            minutes=float(parts["minutes"] or 0),
            seconds=float(parts["seconds"] or 0),
        )
    raise ValueError(f"Invalid datetime value {value!r}")


def _translate_formula_segment(text: str, row_delta: int, col_delta: int) -> str:
    """Translate references in text outside quoted formula tokens."""
    if _WHOLE_RANGE_RE.search(text):
        raise UnsupportedOoxmlError("Whole-row/column formula references")
    if _UNQUOTED_SHEET_RE.search(text) or any(ord(char) > 127 for char in text):
        raise UnsupportedOoxmlError("Qualified or non-ASCII shared formula references")

    def replace(match: re.Match[str]) -> str:
        col_abs, col_letters, row_abs, row_text = match.groups()
        row = int(row_text)
        col = _column_index(col_letters)
        if not col_abs:
            col += col_delta
        if not row_abs:
            row += row_delta
        if not 1 <= col <= 16384 or not 1 <= row <= 1048576:
            raise UnsupportedOoxmlError(
                "Shared formula reference crosses worksheet bounds"
            )
        return f"{col_abs}{get_column_letter(col)}{row_abs}{row}"

    return _FORMULA_CELL_RE.sub(replace, text)


def _quoted_formula_part(formula: str, start: int) -> tuple[str, int]:
    """Return a string, quoted sheet name, or bracket token and next offset."""
    opening = formula[start]
    closing = "]" if opening == "[" else opening
    index = start + 1
    while index < len(formula):
        if formula[index] == closing:
            if closing != "]" and index + 1 < len(formula):
                if formula[index + 1] == closing:
                    index += 2
                    continue
            return formula[start : index + 1], index + 1
        index += 1
    raise UnsupportedOoxmlError("Unterminated formula string, sheet name, or bracket")


def translate_formula(formula: str, origin: str, target: str) -> str:
    """Translate ordinary A1 references in a shared formula without dependencies."""
    origin_row, origin_col = coordinate_to_tuple(origin)
    target_row, target_col = coordinate_to_tuple(target)
    row_delta = target_row - origin_row
    col_delta = target_col - origin_col
    chunks: list[str] = []
    plain: list[str] = []
    index = 0
    while index < len(formula):
        char = formula[index]
        if char in {'"', "'", "["}:
            chunks.append(
                _translate_formula_segment("".join(plain), row_delta, col_delta)
            )
            plain.clear()
            quoted, index = _quoted_formula_part(formula, index)
            chunks.append(quoted)
        else:
            plain.append(char)
            index += 1
    chunks.append(_translate_formula_segment("".join(plain), row_delta, col_delta))
    return "".join(chunks)
