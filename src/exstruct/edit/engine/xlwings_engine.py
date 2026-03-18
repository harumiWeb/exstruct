"""Canonical xlwings engine boundary for workbook editing."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

from exstruct.edit import internal as _internal
from exstruct.edit.models import PatchOp


def apply_xlwings_engine(
    input_path: Path,
    output_path: Path,
    ops: list[PatchOp],
    auto_formula: bool,
) -> list[object]:
    """Apply patch operations using the edit-owned xlwings implementation."""
    diff = _internal._apply_ops_xlwings(
        input_path,
        output_path,
        cast(list[Any], ops),
        auto_formula,
    )
    return list(diff)


__all__ = ["apply_xlwings_engine"]
