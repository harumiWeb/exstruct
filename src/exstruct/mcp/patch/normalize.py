from __future__ import annotations

import json
from typing import Any, cast

from exstruct.mcp.shared.a1 import parse_range_geometry

from .specs import get_alias_map_for_op


def coerce_patch_ops(ops_data: list[dict[str, Any] | str]) -> list[dict[str, Any]]:
    """Normalize patch operations payload for MCP clients."""
    normalized_ops: list[dict[str, Any]] = []
    for index, raw_op in enumerate(ops_data):
        parsed_op = (
            dict(raw_op)
            if isinstance(raw_op, dict)
            else parse_patch_op_json(raw_op, index=index)
        )
        normalized_ops.append(normalize_patch_op_aliases(parsed_op, index=index))
    return normalized_ops


def resolve_top_level_sheet_for_payload(data: object) -> object:
    """Resolve top-level sheet default into operation dict payloads."""
    if not isinstance(data, dict):
        return data
    ops_raw = data.get("ops")
    if not isinstance(ops_raw, list):
        return data
    top_level_sheet = normalize_top_level_sheet(data.get("sheet"))
    resolved_ops: list[object] = []
    for index, op_raw in enumerate(ops_raw):
        if not isinstance(op_raw, dict):
            resolved_ops.append(op_raw)
            continue
        op_copy = normalize_patch_op_aliases(dict(op_raw), index=index)
        op_name_raw = op_copy.get("op")
        op_name = op_name_raw if isinstance(op_name_raw, str) else ""
        op_sheet = op_copy.get("sheet")
        if op_name == "add_sheet":
            if op_copy.get("sheet") is None:
                raise ValueError(
                    build_missing_sheet_message(index=index, op_name="add_sheet")
                )
            resolved_ops.append(op_copy)
            continue
        if op_sheet is None:
            if top_level_sheet is None:
                raise ValueError(
                    build_missing_sheet_message(index=index, op_name=op_name)
                )
            op_copy["sheet"] = top_level_sheet
        resolved_ops.append(op_copy)
    payload = dict(data)
    payload["ops"] = resolved_ops
    if top_level_sheet is not None:
        payload["sheet"] = top_level_sheet
    return payload


def normalize_patch_op_aliases(
    op_data: dict[str, Any], *, index: int
) -> dict[str, Any]:
    """Normalize MCP-friendly aliases to canonical patch operation fields."""
    normalized = dict(op_data)
    op_name = normalized.get("op")
    if not isinstance(op_name, str):
        return normalized
    alias_map = get_alias_map_for_op(op_name)
    for alias, canonical in alias_map.items():
        alias_to_canonical_with_conflict_check(
            normalized,
            index=index,
            alias=alias,
            canonical=canonical,
            op_name=op_name,
        )
    normalize_draw_grid_border_range(normalized, index=index)
    return normalized


def alias_to_canonical_with_conflict_check(
    op_data: dict[str, Any],
    *,
    index: int,
    alias: str,
    canonical: str,
    op_name: str,
) -> None:
    """Map alias field to canonical field when operation type matches."""
    if op_data.get("op") != op_name or alias not in op_data:
        return
    alias_value = op_data[alias]
    canonical_value = op_data.get(canonical)
    if canonical in op_data:
        if canonical_value != alias_value:
            raise ValueError(
                build_patch_op_error_message(
                    index,
                    f"conflicting fields: '{canonical}' and alias '{alias}'",
                )
            )
    else:
        op_data[canonical] = alias_value
    del op_data[alias]


def normalize_draw_grid_border_range(op_data: dict[str, Any], *, index: int) -> None:
    """Convert draw_grid_border range shorthand to base/size fields."""
    if op_data.get("op") != "draw_grid_border" or "range" not in op_data:
        return
    if "base_cell" in op_data or "row_count" in op_data or "col_count" in op_data:
        raise ValueError(
            build_patch_op_error_message(
                index,
                "draw_grid_border does not allow mixing 'range' with 'base_cell/row_count/col_count'",
            )
        )
    range_ref = op_data.get("range")
    if not isinstance(range_ref, str):
        raise ValueError(
            build_patch_op_error_message(
                index, "draw_grid_border range must be a string A1 range"
            )
        )
    try:
        start, row_count, col_count = parse_range_geometry(range_ref)
    except ValueError as exc:
        raise ValueError(
            build_patch_op_error_message(
                index, "draw_grid_border range must be like 'A1:C3'"
            )
        ) from exc
    op_data["base_cell"] = start
    op_data["row_count"] = row_count
    op_data["col_count"] = col_count
    del op_data["range"]


def parse_patch_op_json(raw_op: str, *, index: int) -> dict[str, Any]:
    """Parse a JSON string patch operation into object form."""
    text = raw_op.strip()
    if not text:
        raise ValueError(build_patch_op_error_message(index, "empty string"))
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(build_patch_op_error_message(index, "invalid JSON")) from exc
    if not isinstance(parsed, dict):
        raise ValueError(
            build_patch_op_error_message(index, "JSON value must be an object")
        )
    return cast(dict[str, Any], parsed)


def normalize_top_level_sheet(value: object) -> str | None:
    """Normalize optional top-level sheet text."""
    if not isinstance(value, str):
        return None
    candidate = value.strip()
    return candidate or None


def build_missing_sheet_message(*, index: int, op_name: str) -> str:
    """Build self-healing error for unresolved sheet selection."""
    target_op = op_name or "<unknown>"
    return (
        f"ops[{index}] ({target_op}) is missing sheet. "
        "Set op.sheet, or set top-level sheet for non-add_sheet ops. "
        "For add_sheet, op.sheet (or alias name) is required."
    )


def build_patch_op_error_message(index: int, reason: str) -> str:
    """Build a consistent validation message for invalid patch ops."""
    example = '{"op":"set_value","sheet":"Sheet1","cell":"A1","value":"sample"}'
    return (
        f"Invalid patch operation at ops[{index}]: {reason}. "
        f"Use object form like {example}."
    )
