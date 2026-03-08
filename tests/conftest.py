"""Pytest configuration and environment gating for the test suite."""

from __future__ import annotations

from functools import lru_cache
import importlib.util
import os
from pathlib import Path
import re
import subprocess
import sys
from types import ModuleType

import pytest

IS_WINDOWS = sys.platform == "win32"
SKIP_COM_TESTS = os.getenv("SKIP_COM_TESTS") == "1"
FORCE_COM_TESTS = os.getenv("FORCE_COM_TESTS") == "1"
RUN_RENDER_SMOKE = os.getenv("RUN_RENDER_SMOKE") == "1"
RUN_LIBREOFFICE_SMOKE = os.getenv("RUN_LIBREOFFICE_SMOKE") == "1"


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
    config.addinivalue_line(
        "markers",
        "libreoffice: requires LibreOffice runtime; set RUN_LIBREOFFICE_SMOKE=1 to enable.",
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


@lru_cache(maxsize=1)
def _has_libreoffice_runtime() -> bool:
    """Return True if a runnable LibreOffice executable is available."""
    libreoffice_module = _load_libreoffice_runtime_module()
    raw_path = os.getenv("EXSTRUCT_LIBREOFFICE_PATH")
    which_soffice = libreoffice_module._which_soffice
    resolve_python_path = libreoffice_module._resolve_python_path
    soffice_path = Path(raw_path) if raw_path else which_soffice()
    if not isinstance(soffice_path, Path) or not soffice_path.exists():
        return False
    try:
        python_path = resolve_python_path(soffice_path)
    except Exception:
        return False
    if not isinstance(python_path, Path) or not python_path.exists():
        return False
    try:
        subprocess.run(
            [str(soffice_path), "--version"],
            capture_output=True,
            check=True,
            text=True,
            timeout=5.0,
        )
    except Exception:
        return False
    return True


@lru_cache(maxsize=1)
def _load_libreoffice_runtime_module() -> ModuleType:
    """Load the LibreOffice runtime helper without importing the full package."""

    module_path = (
        Path(__file__).resolve().parents[1] / "src/exstruct/core/libreoffice.py"
    )
    module_name = "tests._libreoffice_runtime_probe"
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Failed to load LibreOffice runtime helper.")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


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


def _libreoffice_skip_reason() -> str | None:
    """
    Return a skip reason for libreoffice-marked tests, or None when they should run.

    LibreOffice smoke tests are gated by RUN_LIBREOFFICE_SMOKE=1.
    """
    if not RUN_LIBREOFFICE_SMOKE:
        return (
            "LibreOffice smoke tests disabled; set RUN_LIBREOFFICE_SMOKE=1 to enable."
        )
    if not _has_libreoffice_runtime():
        return "LibreOffice runtime is unavailable."
    return None


def pytest_runtest_setup(item: pytest.Item) -> None:
    """Skip tests based on resource markers and environment availability."""
    if item.get_closest_marker("render") is not None:
        reason = _render_skip_reason()
        if reason:
            pytest.skip(reason)
        return
    if item.get_closest_marker("libreoffice") is not None:
        reason = _libreoffice_skip_reason()
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
