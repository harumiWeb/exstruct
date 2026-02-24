from __future__ import annotations

from pathlib import Path
from typing import Any, cast

from exstruct.mcp.patch import internal as _internal
from exstruct.mcp.patch.models import PatchRequest


def apply_openpyxl_ops(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
) -> tuple[list[object], list[object], list[object], list[str]]:
    """Apply patch operations using the openpyxl implementation."""
    diff, inverse_ops, formula_issues, op_warnings = _internal._apply_ops_openpyxl(
        cast(Any, request),
        input_path,
        output_path,
    )
    return (
        list(diff),
        list(inverse_ops),
        list(formula_issues),
        list(op_warnings),
    )


__all__ = ["apply_openpyxl_ops"]
