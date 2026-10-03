from __future__ import annotations

from dataclasses import dataclass

from .ooxml_scalars import range_boundaries


@dataclass(frozen=True)
class RangeBounds:
    """Normalized range bounds.

    Attributes:
        r1: Top row (zero-based).
        c1: Left column (zero-based).
        r2: Bottom row (zero-based).
        c2: Right column (zero-based).
    """

    r1: int
    c1: int
    r2: int
    c2: int


def parse_range_zero_based(range_str: str) -> RangeBounds | None:
    """Parse an Excel range string into zero-based bounds.

    Args:
        range_str: Excel range string (e.g., "Sheet1!A1:B2").

    Returns:
        RangeBounds in zero-based coordinates, or None on failure.
    """
    cleaned = range_str.strip()
    if not cleaned:
        return None
    quoted = False
    index = 0
    sheet_separator = -1
    while index < len(cleaned):
        char = cleaned[index]
        if char == "'":
            if quoted and index + 1 < len(cleaned) and cleaned[index + 1] == "'":
                index += 2
                continue
            quoted = not quoted
        elif char == "!" and not quoted:
            sheet_separator = index
        index += 1
    if sheet_separator >= 0:
        cleaned = cleaned[sheet_separator + 1 :]
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cleaned)
    except Exception:
        return None
    if min_col is None or min_row is None or max_col is None or max_row is None:
        return None
    return RangeBounds(
        r1=min_row - 1,
        c1=min_col - 1,
        r2=max_row - 1,
        c2=max_col - 1,
    )
