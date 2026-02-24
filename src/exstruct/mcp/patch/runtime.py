from __future__ import annotations

from pathlib import Path

from exstruct.cli.availability import ComAvailability
from exstruct.mcp.extract_runner import OnConflictPolicy
from exstruct.mcp.io import PathPolicy

from . import legacy_runner as _legacy
from .models import MakeRequest, PatchOp, PatchRequest
from .types import PatchEngine

PatchOpError = _legacy.PatchOpError


def get_com_availability() -> ComAvailability:
    """Return COM availability via the compatibility layer."""
    return _legacy.get_com_availability()


def append_large_ops_warning(warnings: list[str], ops: list[PatchOp]) -> None:
    """Append warnings when patch operation count is large."""
    _legacy._append_large_ops_warning(warnings, ops)


def contains_apply_table_style_op(ops: list[PatchOp]) -> bool:
    """Return whether operations include apply_table_style."""
    return _legacy._contains_apply_table_style_op(ops)


def contains_design_ops(ops: list[PatchOp]) -> bool:
    """Return whether operations include design-affecting ops."""
    return _legacy._contains_design_ops(ops)


def resolve_make_output_path(path: Path, *, policy: PathPolicy | None) -> Path:
    """Resolve output path for make requests."""
    return _legacy._resolve_make_output_path(path, policy=policy)


def ensure_supported_extension(path: Path) -> None:
    """Validate workbook extension for patch/make operations."""
    _legacy._ensure_supported_extension(path)


def validate_make_request_constraints(request: MakeRequest, output_path: Path) -> None:
    """Validate make-request constraints against target output."""
    _legacy._validate_make_request_constraints(request, output_path)


def build_make_seed_path(output_path: Path) -> Path:
    """Return temporary seed workbook path for make operations."""
    return _legacy._build_make_seed_path(output_path)


def resolve_make_initial_sheet_name(request: MakeRequest) -> str:
    """Resolve initial sheet name for make operations."""
    return _legacy._resolve_make_initial_sheet_name(request)


def create_seed_workbook(
    seed_path: Path, extension: str, *, initial_sheet_name: str
) -> None:
    """Create seed workbook used by make operation orchestration."""
    _legacy._create_seed_workbook(
        seed_path,
        extension,
        initial_sheet_name=initial_sheet_name,
    )


def resolve_input_path(path: Path, *, policy: PathPolicy | None) -> Path:
    """Resolve and validate input workbook path."""
    return _legacy._resolve_input_path(path, policy=policy)


def resolve_output_path(
    input_path: Path,
    *,
    out_dir: Path | None,
    out_name: str | None,
    policy: PathPolicy | None,
) -> Path:
    """Resolve and validate output workbook path."""
    return _legacy._resolve_output_path(
        input_path,
        out_dir=out_dir,
        out_name=out_name,
        policy=policy,
    )


def select_patch_engine(
    *, request: PatchRequest, input_path: Path, com_available: bool
) -> PatchEngine:
    """Select runtime patch engine based on request and environment."""
    return _legacy._select_patch_engine(
        request=request,
        input_path=input_path,
        com_available=com_available,
    )


def apply_conflict_policy(
    output_path: Path, on_conflict: OnConflictPolicy
) -> tuple[Path, str | None, bool]:
    """Apply conflict policy to an output path."""
    return _legacy._apply_conflict_policy(output_path, on_conflict)


def requires_openpyxl_backend(request: PatchRequest) -> bool:
    """Return whether request requires openpyxl backend."""
    return _legacy._requires_openpyxl_backend(request)


def ensure_output_dir(path: Path) -> None:
    """Ensure parent directory exists for output path."""
    _legacy._ensure_output_dir(path)


def allow_auto_openpyxl_fallback(request: PatchRequest, input_path: Path) -> bool:
    """Return whether COM failures should fallback to openpyxl."""
    return _legacy._allow_auto_openpyxl_fallback(request, input_path)


def expand_range_coordinates(range_ref: str) -> list[list[str]]:
    """Expand an A1 range into 2D cell coordinates."""
    return _legacy._expand_range_coordinates(range_ref)


__all__ = [
    "PatchOpError",
    "allow_auto_openpyxl_fallback",
    "append_large_ops_warning",
    "apply_conflict_policy",
    "build_make_seed_path",
    "contains_apply_table_style_op",
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
