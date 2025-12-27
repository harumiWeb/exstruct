from __future__ import annotations

from dataclasses import dataclass

from openpyxl.utils import range_boundaries


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
    if "!" in cleaned:
        cleaned = cleaned.split("!", 1)[1]
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cleaned)
    except Exception:
        return None
    return RangeBounds(
        r1=min_row - 1,
        c1=min_col - 1,
        r2=max_row - 1,
        c2=max_col - 1,
    )
