from importlib import util
from pathlib import Path

import pytest

from exstruct.errors import MissingDependencyError
from exstruct.models import CellRow, SheetData, WorkbookData

HAS_PYYAML = util.find_spec("yaml") is not None
HAS_TOON = util.find_spec("toon") is not None


def _sheet() -> SheetData:
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


# pytest.skipif is typed; no ignore needed
@pytest.mark.skipif(not HAS_PYYAML, reason="pyyaml not installed")  # type: ignore[misc]
def test_sheet_to_yaml_roundtrip() -> None:
    sheet = _sheet()
    text = sheet.to_yaml()
    assert "table_candidates" in text
    assert "SheetData" not in text  # not a repr


@pytest.mark.skipif(not HAS_PYYAML, reason="pyyaml not installed")  # type: ignore[misc]
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
    wb = _workbook()
    first = wb["Sheet1"]
    assert isinstance(first, SheetData)
    pairs = list(wb)
    assert len(pairs) == 1
    assert pairs[0][0] == "Sheet1"
    assert pairs[0][1] is first
    with pytest.raises(KeyError):
        _ = wb["Nope"]
