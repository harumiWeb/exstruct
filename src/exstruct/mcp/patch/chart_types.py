from __future__ import annotations

from typing import Final

# Explicit ordered pairs of (chart_type, Excel COM ChartType ID).
# This is the single source of truth; both the ordered tuple and mapping are derived from it.
_CHART_TYPE_ENTRIES: Final[tuple[tuple[str, int], ...]] = (
    ("line", 4),
    ("column", 51),
    ("bar", 57),
    ("area", 1),
    ("pie", 5),
    ("doughnut", -4120),
    ("scatter", -4169),
    ("radar", -4151),
)

# Ordered tuple of canonical chart type names; ordering here determines error messages and docs.
SUPPORTED_CHART_TYPES: Final[tuple[str, ...]] = tuple(t for t, _ in _CHART_TYPE_ENTRIES)

# Mapping from canonical chart type to Excel COM ChartType ID.
CHART_TYPE_TO_COM_ID: Final[dict[str, int]] = dict(_CHART_TYPE_ENTRIES)

CHART_TYPE_ALIASES: Final[dict[str, str]] = {
    "column_clustered": "column",
    "bar_clustered": "bar",
    "xy_scatter": "scatter",
    "donut": "doughnut",
}

SUPPORTED_CHART_TYPES_SET: Final[frozenset[str]] = frozenset(SUPPORTED_CHART_TYPES)
SUPPORTED_CHART_TYPES_CSV: Final[str] = ", ".join(SUPPORTED_CHART_TYPES)


def normalize_chart_type(chart_type: str) -> str | None:
    """Normalize chart type input to a canonical key.

    Args:
        chart_type: Raw chart type value from request payload.

    Returns:
        Canonical chart type key when supported; otherwise ``None``.
    """
    candidate = chart_type.strip().lower()
    canonical = CHART_TYPE_ALIASES.get(candidate, candidate)
    if canonical in SUPPORTED_CHART_TYPES_SET:
        return canonical
    return None


def resolve_chart_type_id(chart_type: str) -> int | None:
    """Resolve canonical chart type key to Excel COM ChartType ID.

    Args:
        chart_type: Raw or canonical chart type value.

    Returns:
        Excel COM ChartType ID when supported; otherwise ``None``.
    """
    normalized = normalize_chart_type(chart_type)
    if normalized is None:
        return None
    return CHART_TYPE_TO_COM_ID[normalized]
