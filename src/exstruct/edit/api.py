"""Public function entry points for workbook editing."""

from __future__ import annotations

from .models import MakeRequest, PatchRequest, PatchResult
from .service import make_workbook as _make_workbook, patch_workbook as _patch_workbook


def patch_workbook(request: PatchRequest) -> PatchResult:
    """Edit an existing workbook without MCP path policy enforcement."""

    return _patch_workbook(request)


def make_workbook(request: MakeRequest) -> PatchResult:
    """Create a new workbook and apply initial patch operations."""

    return _make_workbook(request)


__all__ = [
    "make_workbook",
    "patch_workbook",
    "MakeRequest",
    "PatchRequest",
    "PatchResult",
]
