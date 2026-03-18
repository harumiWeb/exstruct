"""Compatibility shim for legacy patch model imports."""

from __future__ import annotations

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
    OpenpyxlEngineResult,
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
    "OpenpyxlEngineResult",
    "PatchDiffItem",
    "PatchErrorDetail",
    "PatchOp",
    "PatchRequest",
    "PatchResult",
    "PatchValue",
    "RowDimensionSnapshot",
    "XlwingsRangeProtocol",
]
