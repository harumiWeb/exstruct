from __future__ import annotations

from pathlib import Path
from typing import Any, cast

from exstruct.mcp.patch import internal as _internal
from exstruct.mcp.patch.models import OpenpyxlEngineResult, PatchRequest


def apply_openpyxl_ops(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
) -> OpenpyxlEngineResult:
    """Apply patch operations using the openpyxl implementation."""
    diff, inverse_ops, formula_issues, op_warnings = _internal._apply_ops_openpyxl(
        cast(Any, request),
        input_path,
        output_path,
    )
    return OpenpyxlEngineResult(
        patch_diff=list(diff),
        inverse_ops=list(inverse_ops),
        formula_issues=list(formula_issues),
        op_warnings=list(op_warnings),
    )


__all__ = ["apply_openpyxl_ops"]
