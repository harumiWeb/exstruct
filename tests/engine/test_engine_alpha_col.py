"""Tests for alpha_col option through the engine and public API layers."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from exstruct.engine import ExStructEngine, StructOptions
from exstruct.models import CellRow, MergedCells, SheetData, WorkbookData


def _fake_workbook(path: Path, **kwargs: object) -> WorkbookData:
    """Return a minimal WorkbookData with numeric column keys.

    Args:
        path: Input path (unused but required by signature).
        **kwargs: Extra keyword arguments accepted for forward-compatibility.

    Returns:
        A WorkbookData with one sheet and one row.
    """
    return WorkbookData(
        book_name=path.name,
        sheets={
            "Sheet1": SheetData(
                rows=[
                    CellRow(r=1, c={"0": "A1", "1": "B1", "25": "Z1"}, links=None),
                ],
                merged_cells=MergedCells(items=[(1, 0, 3, 2, "merged")]),
            )
        },
    )


def test_engine_alpha_col_false(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """With alpha_col=False, column keys remain numeric."""
    monkeypatch.setattr("exstruct.engine.extract_workbook", _fake_workbook)
    engine = ExStructEngine(options=StructOptions(mode="light", alpha_col=False))
    result = engine.extract(tmp_path / "book.xlsx", mode="light")
    row = result.sheets["Sheet1"].rows[0]
    assert "0" in row.c
    assert "1" in row.c
    assert result.sheets["Sheet1"].merged_ranges == []


def test_engine_alpha_col_true(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """With alpha_col=True, column keys are converted to ABC-style."""
    monkeypatch.setattr("exstruct.engine.extract_workbook", _fake_workbook)
    engine = ExStructEngine(options=StructOptions(mode="light", alpha_col=True))
    result = engine.extract(tmp_path / "book.xlsx", mode="light")
    row = result.sheets["Sheet1"].rows[0]
    assert row.c == {"A": "A1", "B": "B1", "Z": "Z1"}
    assert result.sheets["Sheet1"].merged_ranges == ["A1:C3"]
    assert result.sheets["Sheet1"].merged_cells is None


def test_engine_serialize_alpha_col_includes_merged_ranges(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Serialized output should include merged_ranges when alpha_col=True."""
    monkeypatch.setattr("exstruct.engine.extract_workbook", _fake_workbook)
    engine = ExStructEngine(options=StructOptions(mode="light", alpha_col=True))
    result = engine.extract(tmp_path / "book.xlsx", mode="light")
    payload = json.loads(engine.serialize(result, fmt="json"))
    sheet = payload["sheets"]["Sheet1"]
    assert sheet["merged_ranges"] == ["A1:C3"]
    assert "merged_cells" not in sheet
