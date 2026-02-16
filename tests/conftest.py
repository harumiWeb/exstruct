from __future__ import annotations

from functools import lru_cache
import importlib.util
import os
import re
import sys

import pytest

IS_WINDOWS = sys.platform == "win32"
SKIP_COM_TESTS = os.getenv("SKIP_COM_TESTS") == "1"
FORCE_COM_TESTS = os.getenv("FORCE_COM_TESTS") == "1"
RUN_RENDER_SMOKE = os.getenv("RUN_RENDER_SMOKE") == "1"


def _markexpr_requests_com(markexpr: str) -> bool:
    """Return True when markexpr explicitly requests the ``com`` marker."""
    tokens = re.findall(r"[A-Za-z_][A-Za-z0-9_]*", markexpr.lower())
    for index, token in enumerate(tokens):
        if token != "com":
            continue
        prev = tokens[index - 1] if index > 0 else ""
        if prev != "not":
            return True
    return False


def pytest_configure(config: pytest.Config) -> None:
    """Register custom markers to avoid pytest warnings."""
    markexpr = getattr(config.option, "markexpr", "") or ""
    if _markexpr_requests_com(markexpr):
        os.environ.pop("SKIP_COM_TESTS", None)
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


@lru_cache(maxsize=1)
def _has_pillow() -> bool:
    """Return True if Pillow (PIL) is importable."""
    return importlib.util.find_spec("PIL") is not None


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
    if not _has_pillow():
        return "Pillow (PIL) is unavailable."
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


@pytest.fixture(autouse=True)  # type: ignore[misc]
def _skip_com_for_non_com_tests(
    request: pytest.FixtureRequest, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Disable COM usage for tests that are not marked as COM/render."""
    node_path = str(request.node.path)
    markexpr = getattr(request.config.option, "markexpr", "") or ""
    if _markexpr_requests_com(markexpr):
        monkeypatch.delenv("SKIP_COM_TESTS", raising=False)
        return
    if "tests\\com\\" in node_path or "tests/com/" in node_path:
        monkeypatch.delenv("SKIP_COM_TESTS", raising=False)
        return
    if request.node.get_closest_marker("render") is not None:
        return
    if request.node.get_closest_marker("com") is not None:
        return
    monkeypatch.setenv("SKIP_COM_TESTS", "1")
