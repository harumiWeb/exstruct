"""Compatibility shim for legacy runtime helpers."""

from __future__ import annotations

from pathlib import Path

from exstruct.edit import internal as edit_internal
from exstruct.edit.runtime import (
    ComAvailability,
    PatchOpError,
    allow_auto_openpyxl_fallback,
    append_large_ops_warning,
    apply_conflict_policy,
    build_make_seed_path,
    contains_apply_table_style_op,
    contains_create_chart_op,
    contains_design_ops,
    create_seed_workbook,
    ensure_output_dir,
    ensure_supported_extension,
    expand_range_coordinates,
    get_com_availability,
    requires_openpyxl_backend,
    resolve_make_initial_sheet_name,
    select_patch_engine,
    validate_make_request_constraints,
)
from exstruct.mcp.io import PathPolicy


def resolve_make_output_path(path: Path, *, policy: PathPolicy | None = None) -> Path:
    """Resolve output path for make requests via the legacy policy-aware shim."""
    return edit_internal._resolve_make_output_path(path, policy=policy)


def resolve_input_path(path: Path, *, policy: PathPolicy | None = None) -> Path:
    """Resolve and validate input workbook path via the legacy policy-aware shim."""
    return edit_internal._resolve_input_path(path, policy=policy)


def resolve_output_path(
    input_path: Path,
    *,
    out_dir: Path | None,
    out_name: str | None,
    policy: PathPolicy | None = None,
) -> Path:
    """Resolve and validate output workbook path via the legacy policy-aware shim."""
    return edit_internal._resolve_output_path(
        input_path,
        out_dir=out_dir,
        out_name=out_name,
        policy=policy,
    )


__all__ = [
    "PatchOpError",
    "allow_auto_openpyxl_fallback",
    "append_large_ops_warning",
    "apply_conflict_policy",
    "build_make_seed_path",
    "contains_apply_table_style_op",
    "contains_create_chart_op",
    "contains_design_ops",
    "create_seed_workbook",
    "ensure_output_dir",
    "ensure_supported_extension",
    "expand_range_coordinates",
    "get_com_availability",
    "requires_openpyxl_backend",
    "resolve_input_path",
    "resolve_make_initial_sheet_name",
    "resolve_make_output_path",
    "resolve_output_path",
    "select_patch_engine",
    "validate_make_request_constraints",
    "ComAvailability",
]
