"""Tests for the --alpha-col CLI flag."""

from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook

from exstruct.cli.main import main as cli_main


def _create_workbook(tmp_path: Path) -> Path:
    """Create a minimal Excel workbook for CLI tests.

    Args:
        tmp_path: Temporary directory provided by pytest.

    Returns:
        Path to the generated workbook.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Hello", "World"])
    ws.append([1, 2])
    dest = tmp_path / "alpha_test.xlsx"
    wb.save(dest)
    wb.close()
    return dest


def test_cli_alpha_col_flag(tmp_path: Path) -> None:
    """--alpha-col should produce ABC-style column keys in output."""
    xlsx = _create_workbook(tmp_path)
    out = tmp_path / "out.json"
    code = cli_main([str(xlsx), "-o", str(out), "--alpha-col", "--pretty"])
    assert code == 0
    data = json.loads(out.read_text(encoding="utf-8"))
    rows = data["sheets"]["Sheet1"]["rows"]
    # With alpha_col, column keys should be A, B, ... not 0, 1, ...
    for row in rows:
        cell_keys = set(row["c"].keys())
        assert "A" in cell_keys
        assert "B" in cell_keys
        assert "0" not in cell_keys


def test_cli_without_alpha_col_flag(tmp_path: Path) -> None:
    """Without --alpha-col, column keys should remain 0-based numeric."""
    xlsx = _create_workbook(tmp_path)
    out = tmp_path / "out.json"
    code = cli_main([str(xlsx), "-o", str(out), "--pretty"])
    assert code == 0
    data = json.loads(out.read_text(encoding="utf-8"))
    rows = data["sheets"]["Sheet1"]["rows"]
    for row in rows:
        cell_keys = set(row["c"].keys())
        assert "0" in cell_keys
        assert "A" not in cell_keys


def test_cli_alpha_col_parser_present() -> None:
    """The --alpha-col flag should be accepted by the parser."""
    from exstruct.cli.main import build_parser

    parser = build_parser()
    args = parser.parse_args(["dummy.xlsx", "--alpha-col"])
    assert args.alpha_col is True


def test_cli_alpha_col_default_false() -> None:
    """Without --alpha-col, default should be False."""
    from exstruct.cli.main import build_parser

    parser = build_parser()
    args = parser.parse_args(["dummy.xlsx"])
    assert args.alpha_col is False
