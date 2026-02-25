from __future__ import annotations

from pathlib import Path
from typing import Protocol

from exstruct.mcp.patch.types import PatchOpType


class OpenpyxlPatchEngine(Protocol):
    """Protocol for openpyxl patch engine adapters."""

    def apply(
        self,
        request: object,
        input_path: Path,
        output_path: Path,
    ) -> tuple[list[object], list[object], list[object], list[str]]:
        """Apply patch operations via openpyxl-compatible engine."""


class XlwingsPatchEngine(Protocol):
    """Protocol for xlwings patch engine adapters."""

    def apply(
        self,
        input_path: Path,
        output_path: Path,
        ops: list[object],
        auto_formula: bool,
    ) -> list[object]:
        """Apply patch operations via xlwings-compatible engine."""


__all__ = ["OpenpyxlPatchEngine", "PatchOpType", "XlwingsPatchEngine"]
