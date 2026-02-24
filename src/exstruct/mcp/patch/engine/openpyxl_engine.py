from __future__ import annotations

from pathlib import Path

from exstruct.mcp.patch.models import PatchRequest
from exstruct.mcp.patch.ops.openpyxl_ops import apply_openpyxl_ops


def apply_openpyxl_engine(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
) -> tuple[list[object], list[object], list[object], list[str]]:
    """Apply patch operations using the existing openpyxl backend implementation."""
    return apply_openpyxl_ops(request, input_path, output_path)


__all__ = ["apply_openpyxl_engine"]
