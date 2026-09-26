"""Type aliases and scalar coercion for the public workbook editing contract."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta
from typing import Literal

# Scalar cell payloads. datetime/date/time/timedelta are what openpyxl hands back
# for date- and duration-formatted cells; without them a patch that merely reads
# such a cell for its before-snapshot fails validation before doing any work.
PatchScalar = str | int | float | datetime | date | time | timedelta | None

PatchValueType = Literal["auto", "date"]

PatchOpType = Literal[
    "set_value",
    "set_formula",
    "add_sheet",
    "set_range_values",
    "fill_formula",
    "set_value_if",
    "set_formula_if",
    "draw_grid_border",
    "set_bold",
    "set_font_size",
    "set_font_color",
    "set_fill_color",
    "set_dimensions",
    "auto_fit_columns",
    "merge_cells",
    "unmerge_cells",
    "set_alignment",
    "set_style",
    "apply_table_style",
    "create_chart",
    "restore_design_snapshot",
]
PatchStatus = Literal["applied", "skipped"]
PatchValueKind = Literal["value", "formula", "sheet", "style", "dimension", "chart"]
PatchBackend = Literal["auto", "com", "openpyxl"]
PatchEngine = Literal["com", "openpyxl"]
OnConflictPolicy = Literal["overwrite", "skip", "rename"]
FormulaIssueLevel = Literal["warning", "error"]
FormulaIssueCode = Literal[
    "invalid_token",
    "ref_error",
    "name_error",
    "div0_error",
    "value_error",
    "na_error",
    "circular_ref_suspected",
]

HorizontalAlignType = Literal[
    "general",
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "centerContinuous",
    "distributed",
]
VerticalAlignType = Literal["top", "center", "bottom", "justify", "distributed"]

__all__ = [
    "FormulaIssueCode",
    "FormulaIssueLevel",
    "HorizontalAlignType",
    "OnConflictPolicy",
    "PatchBackend",
    "PatchEngine",
    "PatchOpType",
    "PatchScalar",
    "PatchStatus",
    "PatchValueKind",
    "PatchValueType",
    "VerticalAlignType",
    "coerce_patch_scalar",
]


def coerce_patch_scalar(value: PatchScalar, value_type: PatchValueType) -> PatchScalar:
    """Apply an explicit ``value_type`` hint to a scalar payload.

    JSON has no date literal, so a datetime survives serialization only as an ISO
    string -- and a string round-trips back as a string, silently downgrading a
    date cell to text. ``value_type="date"`` is how a caller (and the inverse-op
    builder) says "this really is a date".
    """

    if value_type != "date" or not isinstance(value, str):
        return value
    for parse in (datetime.fromisoformat, date.fromisoformat, time.fromisoformat):
        try:
            return parse(value)
        except ValueError:
            continue
    raise ValueError(f"value_type='date' but {value!r} is not an ISO date/time.")
