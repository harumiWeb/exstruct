from __future__ import annotations

from typing import Literal

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
    "restore_design_snapshot",
]
PatchStatus = Literal["applied", "skipped"]
PatchValueKind = Literal["value", "formula", "sheet", "style", "dimension"]
PatchBackend = Literal["auto", "com", "openpyxl"]
PatchEngine = Literal["com", "openpyxl"]
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
