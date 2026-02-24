from __future__ import annotations

from pathlib import Path

from exstruct.mcp.patch.models import PatchOp
from exstruct.mcp.patch.ops.xlwings_ops import apply_xlwings_ops


def apply_xlwings_engine(
    input_path: Path,
    output_path: Path,
    ops: list[PatchOp],
    auto_formula: bool,
) -> list[object]:
    """Apply patch operations using the existing xlwings backend implementation."""
    return apply_xlwings_ops(input_path, output_path, ops, auto_formula)


__all__ = ["apply_xlwings_engine"]
