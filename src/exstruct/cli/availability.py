from __future__ import annotations

import logging
import os
import sys

from pydantic import BaseModel, Field
import xlwings as xw

logger = logging.getLogger(__name__)


class ComAvailability(BaseModel):
    """Availability information for Excel COM-dependent features."""

    available: bool = Field(
        ..., description="True when Excel COM can be used from this environment."
    )
    reason: str | None = Field(
        default=None, description="Reason COM features are unavailable."
    )


def get_com_availability() -> ComAvailability:
    """Detect whether Excel COM is available for CLI features.

    Returns:
        ComAvailability describing whether COM features can be used.
    """
    if os.getenv("SKIP_COM_TESTS"):
        return ComAvailability(available=False, reason="SKIP_COM_TESTS is set.")

    if sys.platform != "win32":
        return ComAvailability(available=False, reason="Non-Windows platform.")

    try:
        app = xw.App(add_book=False, visible=False)
    except Exception as exc:
        return ComAvailability(
            available=False,
            reason=f"Excel COM is unavailable ({exc.__class__.__name__}).",
        )

    try:
        app.quit()
    except Exception:
        logger.warning("Failed to quit Excel during COM availability check.")

    return ComAvailability(available=True, reason=None)
