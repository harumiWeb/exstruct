from __future__ import annotations

from typing import Final, cast

from pydantic import BaseModel, Field

from .types import PatchOpType


class PatchOpSpec(BaseModel):
    """Specification metadata used by patch-op normalization."""

    op: PatchOpType
    aliases: dict[str, str] = Field(default_factory=dict)


PATCH_OP_SPECS: Final[dict[PatchOpType, PatchOpSpec]] = {
    "set_value": PatchOpSpec(op="set_value"),
    "set_formula": PatchOpSpec(op="set_formula"),
    "add_sheet": PatchOpSpec(op="add_sheet", aliases={"name": "sheet"}),
    "set_range_values": PatchOpSpec(op="set_range_values"),
    "fill_formula": PatchOpSpec(op="fill_formula"),
    "set_value_if": PatchOpSpec(op="set_value_if"),
    "set_formula_if": PatchOpSpec(op="set_formula_if"),
    "draw_grid_border": PatchOpSpec(op="draw_grid_border"),
    "set_bold": PatchOpSpec(op="set_bold"),
    "set_font_size": PatchOpSpec(op="set_font_size"),
    "set_font_color": PatchOpSpec(op="set_font_color"),
    "set_fill_color": PatchOpSpec(op="set_fill_color", aliases={"color": "fill_color"}),
    "set_dimensions": PatchOpSpec(
        op="set_dimensions",
        aliases={
            "row": "rows",
            "col": "columns",
            "height": "row_height",
            "width": "column_width",
        },
    ),
    "auto_fit_columns": PatchOpSpec(op="auto_fit_columns"),
    "merge_cells": PatchOpSpec(op="merge_cells"),
    "unmerge_cells": PatchOpSpec(op="unmerge_cells"),
    "set_alignment": PatchOpSpec(
        op="set_alignment",
        aliases={"horizontal": "horizontal_align", "vertical": "vertical_align"},
    ),
    "set_style": PatchOpSpec(op="set_style"),
    "apply_table_style": PatchOpSpec(op="apply_table_style"),
    "create_chart": PatchOpSpec(op="create_chart"),
    "restore_design_snapshot": PatchOpSpec(op="restore_design_snapshot"),
}


def get_alias_map_for_op(op_name: str) -> dict[str, str]:
    """Return alias mapping for one operation name."""
    if op_name not in PATCH_OP_SPECS:
        return {}
    spec = PATCH_OP_SPECS[cast(PatchOpType, op_name)]
    return dict(spec.aliases)
