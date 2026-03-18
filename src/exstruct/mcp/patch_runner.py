from __future__ import annotations

from exstruct.edit.errors import PatchOpError
import exstruct.edit.internal as edit_internal
from exstruct.edit.models import (
    AlignmentSnapshot,
    BorderSideSnapshot,
    BorderSnapshot,
    ColumnDimensionSnapshot,
    DesignSnapshot,
    FillSnapshot,
    FontSnapshot,
    FormulaIssue,
    MakeRequest,
    MergeStateSnapshot,
    OpenpyxlWorksheetProtocol,
    PatchDiffItem,
    PatchErrorDetail,
    PatchOp,
    PatchRequest,
    PatchResult,
    PatchValue,
    RowDimensionSnapshot,
    XlwingsRangeProtocol,
)
import exstruct.edit.runtime as edit_runtime

from .io import PathPolicy
from .patch import internal as _internal, runtime as patch_runtime, service

get_com_availability = _internal.get_com_availability


def _sync_legacy_overrides() -> None:
    """Propagate supported monkeypatch overrides to edit and legacy internals."""
    _internal.get_com_availability = get_com_availability
    patch_runtime.get_com_availability = get_com_availability
    edit_runtime.get_com_availability = get_com_availability
    edit_internal.get_com_availability = get_com_availability


def run_make(request: MakeRequest, *, policy: PathPolicy | None = None) -> PatchResult:
    """Compatibility wrapper for make runner."""
    _sync_legacy_overrides()
    return service.run_make(request, policy=policy)


def run_patch(
    request: PatchRequest, *, policy: PathPolicy | None = None
) -> PatchResult:
    """Compatibility wrapper for patch runner."""
    _sync_legacy_overrides()
    return service.run_patch(request, policy=policy)


__all__ = [
    "AlignmentSnapshot",
    "BorderSideSnapshot",
    "BorderSnapshot",
    "ColumnDimensionSnapshot",
    "DesignSnapshot",
    "FillSnapshot",
    "FontSnapshot",
    "FormulaIssue",
    "MakeRequest",
    "MergeStateSnapshot",
    "OpenpyxlWorksheetProtocol",
    "PatchDiffItem",
    "PatchErrorDetail",
    "PatchOp",
    "PatchOpError",
    "PatchRequest",
    "PatchResult",
    "PatchValue",
    "RowDimensionSnapshot",
    "XlwingsRangeProtocol",
    "get_com_availability",
    "run_make",
    "run_patch",
]
