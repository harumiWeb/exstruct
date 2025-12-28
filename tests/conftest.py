from __future__ import annotations

from functools import lru_cache
import importlib.util
import os
import sys

import pytest

IS_WINDOWS = sys.platform == "win32"
SKIP_COM_TESTS = os.getenv("SKIP_COM_TESTS") == "1"
FORCE_COM_TESTS = os.getenv("FORCE_COM_TESTS") == "1"
RUN_RENDER_SMOKE = os.getenv("RUN_RENDER_SMOKE") == "1"


def pytest_configure(config: pytest.Config) -> None:
    """Register custom markers to avoid pytest warnings."""
    config.addinivalue_line("markers", "com: requires Excel COM (Windows + Excel).")
    config.addinivalue_line(
        "markers",
        "render: requires Excel COM and pypdfium2; set RUN_RENDER_SMOKE=1 to enable.",
    )


@lru_cache(maxsize=1)
def _has_excel_com() -> bool:
    """Return True if Excel COM can be opened via xlwings."""
    try:
        import xlwings as xw

        app = xw.App(add_book=False, visible=False)
        app.quit()
        return True
    except Exception:
        return False


@lru_cache(maxsize=1)
def _has_pdfium() -> bool:
    """Return True if pypdfium2 is importable."""
    return importlib.util.find_spec("pypdfium2") is not None


def _com_skip_reason() -> str | None:
    """
    Return a skip reason for COM-marked tests, or None when they should run.

    If FORCE_COM_TESTS=1 and COM is unavailable, raises RuntimeError to fail fast.
    """
    if SKIP_COM_TESTS:
        return "COM tests skipped via SKIP_COM_TESTS=1."
    if not IS_WINDOWS:
        return "COM tests require Windows."
    if not _has_excel_com():
        if FORCE_COM_TESTS:
            raise RuntimeError("Excel COM is unavailable but FORCE_COM_TESTS=1 is set.")
        return "Excel COM is unavailable."
    return None


def _render_skip_reason() -> str | None:
    """
    Return a skip reason for render-marked tests, or None when they should run.

    Render tests require COM + pypdfium2 and are gated by RUN_RENDER_SMOKE=1.
    """
    if not RUN_RENDER_SMOKE:
        return "Render tests disabled; set RUN_RENDER_SMOKE=1 to enable."
    reason = _com_skip_reason()
    if reason:
        return reason
    if not _has_pdfium():
        return "pypdfium2 is unavailable."
    return None


def pytest_runtest_setup(item: pytest.Item) -> None:
    """Skip tests based on resource markers and environment availability."""
    if item.get_closest_marker("render") is not None:
        reason = _render_skip_reason()
        if reason:
            pytest.skip(reason)
        return
    if item.get_closest_marker("com") is not None:
        reason = _com_skip_reason()
        if reason:
            pytest.skip(reason)
