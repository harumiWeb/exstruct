from __future__ import annotations

from .io import PathPolicy
from .patch import legacy_runner as _legacy

AlignmentSnapshot = _legacy.AlignmentSnapshot
BorderSideSnapshot = _legacy.BorderSideSnapshot
BorderSnapshot = _legacy.BorderSnapshot
ColumnDimensionSnapshot = _legacy.ColumnDimensionSnapshot
DesignSnapshot = _legacy.DesignSnapshot
FillSnapshot = _legacy.FillSnapshot
FontSnapshot = _legacy.FontSnapshot
FormulaIssue = _legacy.FormulaIssue
MakeRequest = _legacy.MakeRequest
MergeStateSnapshot = _legacy.MergeStateSnapshot
OpenpyxlWorksheetProtocol = _legacy.OpenpyxlWorksheetProtocol
PatchDiffItem = _legacy.PatchDiffItem
PatchErrorDetail = _legacy.PatchErrorDetail
PatchOp = _legacy.PatchOp
PatchOpError = _legacy.PatchOpError
PatchRequest = _legacy.PatchRequest
PatchResult = _legacy.PatchResult
PatchValue = _legacy.PatchValue
RowDimensionSnapshot = _legacy.RowDimensionSnapshot
XlwingsRangeProtocol = _legacy.XlwingsRangeProtocol
get_com_availability = _legacy.get_com_availability


def _sync_legacy_overrides() -> None:
    """Propagate supported monkeypatch overrides to legacy module."""
    _legacy.get_com_availability = get_com_availability


def run_make(request: MakeRequest, *, policy: PathPolicy | None = None) -> PatchResult:
    """Compatibility wrapper for make runner."""
    _sync_legacy_overrides()
    return _legacy.run_make(request, policy=policy)


def run_patch(
    request: PatchRequest, *, policy: PathPolicy | None = None
) -> PatchResult:
    """Compatibility wrapper for patch runner."""
    _sync_legacy_overrides()
    return _legacy.run_patch(request, policy=policy)


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
