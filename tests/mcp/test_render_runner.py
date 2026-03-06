from __future__ import annotations

import logging

import pytest

from exstruct import render
from exstruct.mcp import render_runner


class _FailingQuitApp:
    """Excel app stub whose quit method fails."""

    def quit(self) -> None:
        raise RuntimeError("quit failed")


def test_ensure_com_available_ignores_quit_errors(
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Do not fail COM probe when app teardown raises.

    Args:
        monkeypatch: Fixture for patching module attributes.
        caplog: Fixture for asserting emitted log messages.

    Returns:
        None.
    """
    monkeypatch.setattr(
        render,
        "_require_excel_app",
        lambda: _FailingQuitApp(),
    )

    caplog.set_level(logging.WARNING, logger=render_runner.__name__)
    render_runner._ensure_com_available()

    assert "Failed to close Excel app after COM probe" in caplog.text


def test_ensure_com_available_raises_value_error_when_com_missing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Wrap COM availability errors as ValueError for tool-level handling.

    Args:
        monkeypatch: Fixture for patching module attributes.

    Returns:
        None.
    """

    def _raise() -> object:
        raise RuntimeError("com missing")

    monkeypatch.setattr(render, "_require_excel_app", _raise)

    with pytest.raises(ValueError, match="Excel \\(COM\\) is not available"):
        render_runner._ensure_com_available()
