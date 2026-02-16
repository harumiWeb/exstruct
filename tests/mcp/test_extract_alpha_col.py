"""Tests for alpha_col support in MCP extract_runner."""

from __future__ import annotations

from exstruct.mcp.extract_runner import ExtractOptions


def test_extract_options_alpha_col_default() -> None:
    """alpha_col should default to True in MCP."""
    opts = ExtractOptions()
    assert opts.alpha_col is True


def test_extract_options_alpha_col_true() -> None:
    """alpha_col should be settable to True."""
    opts = ExtractOptions(alpha_col=True)
    assert opts.alpha_col is True


def test_extract_options_alpha_col_false() -> None:
    """alpha_col should be settable to False to opt out."""
    opts = ExtractOptions(alpha_col=False)
    assert opts.alpha_col is False


def test_extract_options_from_dict() -> None:
    """alpha_col should be parseable from a plain dict (MCP JSON payload)."""
    opts = ExtractOptions.model_validate({"alpha_col": True, "pretty": True})
    assert opts.alpha_col is True
    assert opts.pretty is True
