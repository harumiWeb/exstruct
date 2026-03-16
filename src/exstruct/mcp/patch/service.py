"""Compatibility wrappers for the legacy patch service import path."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

from exstruct.edit.models import MakeRequest, PatchRequest, PatchResult
import exstruct.edit.runtime as edit_runtime
import exstruct.edit.service as edit_service
from exstruct.mcp.io import PathPolicy
from exstruct.mcp.patch.engine import (
    openpyxl_engine as legacy_openpyxl_engine,
    xlwings_engine as legacy_xlwings_engine,
)
import exstruct.mcp.patch.runtime as runtime


def apply_openpyxl_engine(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
) -> object:
    """Call the current legacy openpyxl engine boundary via live module lookup."""
    return legacy_openpyxl_engine.apply_openpyxl_engine(
        request,
        input_path,
        output_path,
    )


def apply_xlwings_engine(
    input_path: Path,
    output_path: Path,
    ops: list[object],
    auto_formula: bool,
) -> list[object]:
    """Call the current legacy xlwings engine boundary via live module lookup."""
    return legacy_xlwings_engine.apply_xlwings_engine(
        input_path,
        output_path,
        cast(list[Any], ops),
        auto_formula,
    )


def _sync_compat_overrides() -> None:
    """Propagate legacy monkeypatch targets into the canonical edit service."""
    service_module = cast(Any, edit_service)
    runtime_module = cast(Any, edit_runtime)
    service_module.apply_openpyxl_engine = apply_openpyxl_engine
    service_module.apply_xlwings_engine = apply_xlwings_engine
    runtime_module.get_com_availability = runtime.get_com_availability
    runtime_module.PatchOpError = runtime.PatchOpError


def run_make(request: MakeRequest, *, policy: PathPolicy | None = None) -> PatchResult:
    """Compatibility wrapper for legacy make orchestration."""
    _sync_compat_overrides()
    resolved_request = _resolve_make_request_paths(request, policy=policy)
    return edit_service.make_workbook(resolved_request)


def run_patch(
    request: PatchRequest, *, policy: PathPolicy | None = None
) -> PatchResult:
    """Compatibility wrapper for legacy patch orchestration."""
    _sync_compat_overrides()
    resolved_request = _resolve_patch_request_paths(request, policy=policy)
    return edit_service.patch_workbook(resolved_request)


def _resolve_make_request_paths(
    request: MakeRequest, *, policy: PathPolicy | None
) -> MakeRequest:
    """Canonicalize make-request paths under the host path policy."""
    if policy is None:
        return request
    return request.model_copy(
        update={"out_path": policy.ensure_allowed(request.out_path)}
    )


def _resolve_patch_request_paths(
    request: PatchRequest, *, policy: PathPolicy | None
) -> PatchRequest:
    """Canonicalize patch-request paths under the host path policy."""
    if policy is None:
        return request
    resolved_out_dir = (
        policy.ensure_allowed(Path(request.out_dir))
        if request.out_dir is not None
        else None
    )
    return request.model_copy(
        update={
            "xlsx_path": policy.ensure_allowed(request.xlsx_path),
            "out_dir": resolved_out_dir,
        }
    )


__all__ = [
    "apply_openpyxl_engine",
    "apply_xlwings_engine",
    "run_make",
    "run_patch",
    "runtime",
]
