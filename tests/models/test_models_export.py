from collections.abc import Callable
from importlib import util
import json
from pathlib import Path
from typing import Any, cast

import pytest

from exstruct.errors import MissingDependencyError
from exstruct.models import (
    CellRow,
    MergedCells,
    SheetData,
    SmartArt,
    SmartArtNode,
    WorkbookData,
)

HAS_PYYAML = util.find_spec("yaml") is not None
HAS_TOON = util.find_spec("toon") is not None
_SkipIf = Callable[[Callable[..., Any]], Callable[..., Any]]


def _sheet() -> SheetData:
    """
    Create a sample SheetData containing one row, no shapes or charts, and a single table candidate.
    
    Returns:
        SheetData: A SheetData instance with one CellRow (r=1, c={"0": "A"}), empty shapes and charts lists, and table_candidates set to ["A1:B2"].
    """
    return SheetData(
        rows=[CellRow(r=1, c={"0": "A"})],
        shapes=[],
        charts=[],
        table_candidates=["A1:B2"],
    )


def _workbook() -> WorkbookData:
    return WorkbookData(book_name="sample.xlsx", sheets={"Sheet1": _sheet()})


def test_workbook_to_json_pretty() -> None:
    wb = _workbook()
    text = wb.to_json(pretty=True)
    assert '"book_name": "sample.xlsx"' in text
    assert '\n  "' in text  # pretty indent applied


def test_sheet_to_json_compact() -> None:
    sheet = _sheet()
    text = sheet.to_json()
    assert '"r": 1' in text
    assert "table_candidates" in text


def test_save_respects_suffix(tmp_path: Path) -> None:
    wb = _workbook()
    dest = tmp_path / "out.json"
    wb.save(dest, pretty=False)
    assert dest.exists()
    content = dest.read_text(encoding="utf-8")
    assert '"book_name": "sample.xlsx"' in content

    sheet = _sheet()
    sdest = tmp_path / "sheet.json"
    sheet.save(sdest)
    assert sdest.exists()
    assert '"table_candidates": [' in sdest.read_text(encoding="utf-8")


def test_save_unsupported_format_raises(tmp_path: Path) -> None:
    wb = _workbook()
    bad = tmp_path / "out.bin"
    with pytest.raises(ValueError):
        wb.save(bad)


# cast to _SkipIf satisfies mypy strict mode for decorator typing
@cast(_SkipIf, pytest.mark.skipif(not HAS_PYYAML, reason="pyyaml not installed"))
def test_sheet_to_yaml_roundtrip() -> None:
    sheet = _sheet()
    text = sheet.to_yaml()
    assert "table_candidates" in text
    assert "SheetData" not in text  # not a repr


@cast(_SkipIf, pytest.mark.skipif(not HAS_PYYAML, reason="pyyaml not installed"))
def test_workbook_to_yaml() -> None:
    wb = _workbook()
    text = wb.to_yaml()
    assert "book_name: sample.xlsx" in text


def test_sheet_to_toon_dependency() -> None:
    sheet = _sheet()
    if HAS_TOON:
        text = sheet.to_toon()
        assert isinstance(text, str) and text
    else:
        with pytest.raises((RuntimeError, MissingDependencyError)):
            sheet.to_toon()


def test_workbook_iter_and_getitem() -> None:
    """
    Verify Workbook supports lookup by sheet name and iteration over (name, sheet) pairs, and that missing sheet names raise KeyError.

    Asserts that retrieving a known sheet by key returns the SheetData instance, that iterating the workbook yields a single (name, sheet) pair matching the retrieved sheet, and that accessing a nonexistent sheet raises KeyError.
    """
    wb = _workbook()
    first = wb["Sheet1"]
    assert isinstance(first, SheetData)
    pairs = list(wb)
    assert len(pairs) == 1
    assert pairs[0][0] == "Sheet1"
    assert pairs[0][1] is first
    with pytest.raises(KeyError):
        _ = wb["Nope"]


def test_sheet_json_includes_smartart_nodes() -> None:
    smartart = SmartArt(
        id=1,
        text="sa",
        l=0,
        t=0,
        w=10,
        h=10,
        layout="Layout",
        nodes=[
            SmartArtNode(
                text="root",
                kids=[SmartArtNode(text="child", kids=[])],
            )
        ],
    )
    sheet = SheetData(
        rows=[],
        shapes=[smartart],
        charts=[],
        table_candidates=[],
    )
    data = json.loads(sheet.to_json())
    assert data["shapes"][0]["kind"] == "smartart"
    assert data["shapes"][0]["nodes"][0]["text"] == "root"
    assert data["shapes"][0]["nodes"][0]["kids"][0]["text"] == "child"


def test_sheet_json_includes_merged_cells_schema() -> None:
    """
    Verify that SheetData.to_json serializes merged_cells with schema and items.

    Asserts that the JSON output includes a merged_cells object with a schema
    field containing ["r1", "c1", "r2", "c2", "v"] and an items array with the
    provided merged cell data as a 5-element array.
    """
    sheet = SheetData(
        rows=[],
        merged_cells=MergedCells(items=[(1, 0, 1, 1, "merged")]),
    )
    data = json.loads(sheet.to_json())
    assert data["merged_cells"]["schema"] == ["r1", "c1", "r2", "c2", "v"]
    assert data["merged_cells"]["items"][0] == [1, 0, 1, 1, "merged"]