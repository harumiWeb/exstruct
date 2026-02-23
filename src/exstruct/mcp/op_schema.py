from __future__ import annotations

from typing import Any, get_args

from pydantic import BaseModel, Field

from .patch_runner import PatchOpType


class PatchOpSchema(BaseModel):
    """Mini schema metadata for a patch operation."""

    op: PatchOpType
    description: str
    required: list[str] = Field(default_factory=list)
    optional: list[str] = Field(default_factory=list)
    constraints: list[str] = Field(default_factory=list)
    example: dict[str, Any]
    aliases: dict[str, str] = Field(default_factory=dict)


def list_patch_op_schemas() -> list[PatchOpSchema]:
    """Return patch operation schemas in canonical op order."""
    ordered_ops = list(get_args(PatchOpType))
    return [_PATCH_OP_SCHEMA_BY_NAME[op] for op in ordered_ops]


def get_patch_op_schema(op: str) -> PatchOpSchema | None:
    """Return schema for one patch op name."""
    schema = _PATCH_OP_SCHEMA_BY_NAME.get(op)
    return schema


def build_patch_tool_mini_schema() -> str:
    """Build a human-readable mini schema section for exstruct_patch doc."""
    lines: list[str] = []
    lines.append("Mini op schema (required/optional/constraints/example/aliases):")
    lines.append(
        "Sheet resolution: non-add_sheet ops allow top-level sheet fallback; "
        "op.sheet overrides top-level sheet."
    )
    for schema in list_patch_op_schemas():
        display_schema = schema_with_sheet_resolution_rules(schema)
        lines.append(f"- {schema.op}: {schema.description}")
        lines.append(
            "  required: "
            + (
                ", ".join(display_schema.required)
                if display_schema.required
                else "(none)"
            )
        )
        lines.append(
            "  optional: "
            + (
                ", ".join(display_schema.optional)
                if display_schema.optional
                else "(none)"
            )
        )
        lines.append(
            "  constraints: "
            + (
                ", ".join(display_schema.constraints)
                if display_schema.constraints
                else "(none)"
            )
        )
        lines.append(f"  example: {display_schema.example}")
        lines.append(
            "  aliases: "
            + (
                ", ".join(
                    f"{alias} -> {canonical}"
                    for alias, canonical in sorted(display_schema.aliases.items())
                )
                if display_schema.aliases
                else "(none)"
            )
        )
    return "\n".join(lines)


def schema_with_sheet_resolution_rules(schema: PatchOpSchema) -> PatchOpSchema:
    """Return display schema with top-level sheet resolution notes."""
    if "sheet" not in schema.required:
        return schema
    updated_required = [
        (
            "sheet (or top-level sheet)"
            if item == "sheet" and schema.op != "add_sheet"
            else item
        )
        for item in schema.required
    ]
    updated_constraints = list(schema.constraints)
    if schema.op == "add_sheet":
        updated_constraints.append("top-level sheet is not used for add_sheet")
    else:
        updated_constraints.append(
            "op.sheet overrides top-level sheet when both are set"
        )
    return schema.model_copy(
        update={"required": updated_required, "constraints": updated_constraints}
    )


_PATCH_OP_SCHEMA_BY_NAME: dict[str, PatchOpSchema] = {
    "set_value": PatchOpSchema(
        op="set_value",
        description="Set a scalar value to one cell.",
        required=["sheet", "cell", "value"],
        optional=[],
        constraints=[
            "cell target only",
            "use auto_formula=true to allow values starting with '='",
        ],
        example={"op": "set_value", "sheet": "Sheet1", "cell": "A1", "value": "Hello"},
    ),
    "set_formula": PatchOpSchema(
        op="set_formula",
        description="Set one formula string to one cell.",
        required=["sheet", "cell", "formula"],
        optional=[],
        constraints=["formula must start with '='"],
        example={
            "op": "set_formula",
            "sheet": "Sheet1",
            "cell": "B2",
            "formula": "=SUM(B1:B10)",
        },
    ),
    "add_sheet": PatchOpSchema(
        op="add_sheet",
        description="Add a new worksheet by name.",
        required=["sheet"],
        optional=[],
        constraints=["sheet name must be unique in workbook"],
        example={"op": "add_sheet", "sheet": "Data"},
        aliases={"name": "sheet"},
    ),
    "set_range_values": PatchOpSchema(
        op="set_range_values",
        description="Set a 2D values matrix to a rectangular range.",
        required=["sheet", "range", "values"],
        optional=[],
        constraints=["values shape must match range rows x cols"],
        example={
            "op": "set_range_values",
            "sheet": "Sheet1",
            "range": "A1:B2",
            "values": [[1, 2], [3, 4]],
        },
    ),
    "fill_formula": PatchOpSchema(
        op="fill_formula",
        description="Fill a base formula across target range.",
        required=["sheet", "range", "base_cell", "formula"],
        optional=[],
        constraints=["formula must start with '='"],
        example={
            "op": "fill_formula",
            "sheet": "Sheet1",
            "range": "C2:C10",
            "base_cell": "C2",
            "formula": "=A2+B2",
        },
    ),
    "set_value_if": PatchOpSchema(
        op="set_value_if",
        description="Set value when current value matches expected.",
        required=["sheet", "cell", "expected", "value"],
        optional=[],
        constraints=["no-op when expected mismatch"],
        example={
            "op": "set_value_if",
            "sheet": "Sheet1",
            "cell": "A1",
            "expected": "old",
            "value": "new",
        },
    ),
    "set_formula_if": PatchOpSchema(
        op="set_formula_if",
        description="Set formula when current value matches expected.",
        required=["sheet", "cell", "expected", "formula"],
        optional=[],
        constraints=["formula must start with '='", "no-op when expected mismatch"],
        example={
            "op": "set_formula_if",
            "sheet": "Sheet1",
            "cell": "C5",
            "expected": 0,
            "formula": "=A5+B5",
        },
    ),
    "draw_grid_border": PatchOpSchema(
        op="draw_grid_border",
        description="Draw thin black borders for a rectangular region.",
        required=["sheet", "base_cell", "row_count", "col_count"],
        optional=[],
        constraints=["row_count > 0", "col_count > 0", "or use range shorthand alias"],
        example={
            "op": "draw_grid_border",
            "sheet": "Sheet1",
            "base_cell": "A1",
            "row_count": 5,
            "col_count": 4,
        },
        aliases={"range": "base_cell + row_count + col_count"},
    ),
    "set_bold": PatchOpSchema(
        op="set_bold",
        description="Apply bold font to one cell or one range.",
        required=["sheet"],
        optional=["cell", "range", "bold"],
        constraints=["exactly one of cell or range"],
        example={"op": "set_bold", "sheet": "Sheet1", "range": "A1:D1", "bold": True},
    ),
    "set_font_size": PatchOpSchema(
        op="set_font_size",
        description="Apply font size to one cell or one range.",
        required=["sheet", "font_size"],
        optional=["cell", "range"],
        constraints=["exactly one of cell or range", "font_size > 0"],
        example={
            "op": "set_font_size",
            "sheet": "Sheet1",
            "cell": "A1",
            "font_size": 14,
        },
    ),
    "set_font_color": PatchOpSchema(
        op="set_font_color",
        description="Apply font color to one cell or one range.",
        required=["sheet", "color"],
        optional=["cell", "range"],
        constraints=[
            "exactly one of cell or range",
            "hex color (#RRGGBB or #AARRGGBB)",
        ],
        example={
            "op": "set_font_color",
            "sheet": "Sheet1",
            "range": "A1:D1",
            "color": "#1F4E79",
        },
    ),
    "set_fill_color": PatchOpSchema(
        op="set_fill_color",
        description="Apply fill color to one cell or one range.",
        required=["sheet", "fill_color"],
        optional=["cell", "range"],
        constraints=[
            "exactly one of cell or range",
            "hex color (#RRGGBB or #AARRGGBB)",
        ],
        example={
            "op": "set_fill_color",
            "sheet": "Sheet1",
            "range": "A1:D1",
            "fill_color": "#D9E1F2",
        },
        aliases={"color": "fill_color"},
    ),
    "set_dimensions": PatchOpSchema(
        op="set_dimensions",
        description="Set row height and/or column width.",
        required=["sheet"],
        optional=["rows", "columns", "row_height", "column_width"],
        constraints=[
            "at least one of row_height or column_width",
            "row_height > 0, column_width > 0",
        ],
        example={
            "op": "set_dimensions",
            "sheet": "Sheet1",
            "rows": [1, 2],
            "row_height": 22,
            "columns": ["A", "B"],
            "column_width": 18,
        },
        aliases={
            "row": "rows",
            "col": "columns",
            "height": "row_height",
            "width": "column_width",
        },
    ),
    "auto_fit_columns": PatchOpSchema(
        op="auto_fit_columns",
        description="Auto-fit column widths with optional width bounds.",
        required=["sheet"],
        optional=["columns", "min_width", "max_width"],
        constraints=[
            "columns optional (uses used columns when omitted)",
            "min_width > 0, max_width > 0 when provided",
            "min_width <= max_width when both are provided",
        ],
        example={
            "op": "auto_fit_columns",
            "sheet": "Sheet1",
            "columns": ["A", 2],
            "min_width": 8,
            "max_width": 40,
        },
    ),
    "merge_cells": PatchOpSchema(
        op="merge_cells",
        description="Merge one rectangular range.",
        required=["sheet", "range"],
        optional=[],
        constraints=["range must be rectangular"],
        example={"op": "merge_cells", "sheet": "Sheet1", "range": "A1:C1"},
    ),
    "unmerge_cells": PatchOpSchema(
        op="unmerge_cells",
        description="Unmerge merged cells intersecting range.",
        required=["sheet", "range"],
        optional=[],
        constraints=["range must be rectangular"],
        example={"op": "unmerge_cells", "sheet": "Sheet1", "range": "A1:C1"},
    ),
    "set_alignment": PatchOpSchema(
        op="set_alignment",
        description="Set alignment flags to one cell or one range.",
        required=["sheet"],
        optional=["cell", "range", "horizontal_align", "vertical_align", "wrap_text"],
        constraints=[
            "exactly one of cell or range",
            "specify at least one alignment field",
        ],
        example={
            "op": "set_alignment",
            "sheet": "Sheet1",
            "range": "A1:D1",
            "horizontal_align": "center",
            "vertical_align": "center",
            "wrap_text": True,
        },
        aliases={"horizontal": "horizontal_align", "vertical": "vertical_align"},
    ),
    "set_style": PatchOpSchema(
        op="set_style",
        description="Apply multiple style attributes in one op.",
        required=["sheet"],
        optional=[
            "cell",
            "range",
            "bold",
            "font_size",
            "color",
            "fill_color",
            "horizontal_align",
            "vertical_align",
            "wrap_text",
        ],
        constraints=[
            "exactly one of cell or range",
            "specify at least one style field",
            "font_size > 0",
            "target cell count <= style limit",
        ],
        example={
            "op": "set_style",
            "sheet": "Sheet1",
            "range": "A1:D1",
            "bold": True,
            "color": "#FFFFFF",
            "fill_color": "#1F3864",
            "horizontal_align": "center",
        },
    ),
    "apply_table_style": PatchOpSchema(
        op="apply_table_style",
        description="Create table and apply Excel table style.",
        required=["sheet", "range", "style"],
        optional=["table_name"],
        constraints=[
            "range must include header row",
            "table name must be unique when provided",
            "range must not intersect existing table",
            "backend='com' falls back to openpyxl with warning",
        ],
        example={
            "op": "apply_table_style",
            "sheet": "Sheet1",
            "range": "A1:D11",
            "style": "TableStyleMedium9",
            "table_name": "SalesTable",
        },
    ),
    "restore_design_snapshot": PatchOpSchema(
        op="restore_design_snapshot",
        description="Internal inverse op to restore style snapshot.",
        required=["sheet", "design_snapshot"],
        optional=[],
        constraints=["internal use for inverse operations"],
        example={
            "op": "restore_design_snapshot",
            "sheet": "Sheet1",
            "design_snapshot": {"range": "A1:A1", "cells": []},
        },
    ),
}
