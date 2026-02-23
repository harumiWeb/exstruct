from __future__ import annotations

import sys

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
    """Propagate monkeypatched private helpers to legacy module."""
    module = sys.modules[__name__]
    for name, value in vars(module).items():
        if name in {
            "__name__",
            "__doc__",
            "__package__",
            "__loader__",
            "__spec__",
            "__file__",
            "__cached__",
            "__builtins__",
        }:
            continue
        if name in {"run_make", "run_patch", "_sync_legacy_overrides"}:
            continue
        if hasattr(_legacy, name):
            setattr(_legacy, name, value)


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


# Re-export private helpers used by tests and internal modules.
_apply_openpyxl_set_fill_color = _legacy._apply_openpyxl_set_fill_color
_apply_ops_openpyxl = _legacy._apply_ops_openpyxl
_apply_ops_xlwings = _legacy._apply_ops_xlwings
_apply_xlwings_set_font_size = _legacy._apply_xlwings_set_font_size
_collect_openpyxl_target_column_max_lengths = (
    _legacy._collect_openpyxl_target_column_max_lengths
)

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
