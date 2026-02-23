from __future__ import annotations

from pathlib import Path

import exstruct.mcp.patch.legacy_runner as runner
from exstruct.mcp.patch.legacy_runner import PatchOp


def apply_xlwings_engine(
    input_path: Path,
    output_path: Path,
    ops: list[PatchOp],
    auto_formula: bool,
) -> list[object]:
    """Apply patch operations using the existing xlwings backend implementation."""
    diff = runner._apply_ops_xlwings(
        input_path,
        output_path,
        ops,
        auto_formula,
    )
    return list(diff)


__all__ = ["apply_xlwings_engine"]
