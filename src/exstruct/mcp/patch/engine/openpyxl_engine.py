from __future__ import annotations

from pathlib import Path

import exstruct.mcp.patch.legacy_runner as runner
from exstruct.mcp.patch.legacy_runner import PatchRequest


def apply_openpyxl_engine(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
) -> tuple[list[object], list[object], list[object], list[str]]:
    """Apply patch operations using the existing openpyxl backend implementation."""
    diff, inverse_ops, formula_issues, op_warnings = runner._apply_ops_openpyxl(
        request,
        input_path,
        output_path,
    )
    return (
        list(diff),
        list(inverse_ops),
        list(formula_issues),
        list(op_warnings),
    )


__all__ = ["apply_openpyxl_engine"]
