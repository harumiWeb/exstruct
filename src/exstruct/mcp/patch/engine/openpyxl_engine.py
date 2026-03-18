"""Compatibility wrapper for the legacy openpyxl engine import path."""

from __future__ import annotations

from pathlib import Path

from exstruct.mcp.patch.models import OpenpyxlEngineResult, PatchRequest
from exstruct.mcp.patch.ops.openpyxl_ops import apply_openpyxl_ops


def apply_openpyxl_engine(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
) -> OpenpyxlEngineResult:
    """Apply patch operations using the legacy openpyxl engine boundary."""
    return apply_openpyxl_ops(request, input_path, output_path)


__all__ = ["apply_openpyxl_engine"]
