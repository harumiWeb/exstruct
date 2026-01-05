import json
from pathlib import Path

import pytest

from exstruct import (
    CellRow,
    Chart,
    ChartSeries,
    Shape,
    SheetData,
    WorkbookData,
    export,
    export_sheets,
    export_sheets_as,
)
from exstruct.io import dict_without_empty_values


def test_空値はdict_without_empty_valuesで除外される() -> None:
    data = {
        "a": None,
        "b": "",
        "c": [],
        "d": {},
        "e": {"x": "", "y": 1},
        "f": [None, "", 2],
    }
    filtered = dict_without_empty_values(data)
    assert filtered == {"e": {"y": 1}, "f": [2]}


def test_JSON出力はUTF8で保存される(tmp_path: Path) -> None:
    wb = WorkbookData(book_name="b.xlsx", sheets={})
    out = tmp_path / "out.json"
    export(wb, out, fmt="json")
    raw = out.read_bytes()
    assert raw.startswith(b"{")
    text = raw.decode("utf-8")
    assert '"book_name": "b.xlsx"' in text


def test_YAML出力はsort_keys_false(tmp_path: Path) -> None:
    pytest.importorskip("yaml")
    wb = WorkbookData(book_name="b.xlsx", sheets={})
    out = tmp_path / "out.yaml"
    export(wb, out, fmt="yaml")
    content = out.read_text(encoding="utf-8")
    # order should preserve book_name before sheets when dumped without sort_keys
    assert content.strip().startswith("book_name:")


def test_TOON出力が生成される(tmp_path: Path) -> None:
    pytest.importorskip("toon")
    wb = WorkbookData(book_name="b.xlsx", sheets={})
    out = tmp_path / "out.toon"
    export(wb, out, fmt="toon")
    content = out.read_text(encoding="utf-8")
    assert "book_name" in content


def test_JSON_roundtripが破壊的変更なし(tmp_path: Path) -> None:
    sheet = SheetData(
        rows=[CellRow(r=1, c={"0": 123})],
        shapes=[Shape(id=1, text="t", l=1, t=2, w=None, h=None)],
        charts=[
            Chart(
                name="c",
                chart_type="Line",
                title=None,
                y_axis_title="",
                y_axis_range=[],
                series=[ChartSeries(name="s")],
                l=0,
                t=0,
            )
        ],
        table_candidates=["A1:B2"],
    )
    wb = WorkbookData(book_name="b.xlsx", sheets={"S": sheet})
    out = tmp_path / "round.json"
    export(wb, out, fmt="json")
    loaded = json.loads(out.read_text(encoding="utf-8"))
    assert loaded["book_name"] == "b.xlsx"
    assert loaded["sheets"]["S"]["rows"][0]["c"]["0"] == 123
    assert loaded["sheets"]["S"]["shapes"][0]["text"] == "t"
    assert loaded["sheets"]["S"]["charts"][0]["series"][0]["name"] == "s"


def test_export_sheetsでシートごとにファイルが出力される(tmp_path: Path) -> None:
    sheet = SheetData(rows=[CellRow(r=1, c={"0": "v"})])
    wb = WorkbookData(
        book_name="book.xlsx", sheets={"SheetA": sheet, "SheetB": SheetData()}
    )
    outdir = tmp_path / "sheets"
    paths = export_sheets(wb, outdir)
    assert set(paths.keys()) == {"SheetA", "SheetB"}
    for p in paths.values():
        assert p.exists()


def test_export_sheets_yamlとtoonが出力される(tmp_path: Path) -> None:
    sheet = SheetData(rows=[CellRow(r=1, c={"0": "v"})])
    wb = WorkbookData(book_name="book.xlsx", sheets={"SheetA": sheet})
    yaml_dir = tmp_path / "yaml_sheets"
    yaml_paths = export_sheets_as(wb, yaml_dir, fmt="yaml")
    assert yaml_paths["SheetA"].suffix == ".yaml"
    assert yaml_paths["SheetA"].exists()

    pytest.importorskip("toon")
    toon_dir = tmp_path / "toon_sheets"
    toon_paths = export_sheets_as(wb, toon_dir, fmt="toon")
    assert toon_paths["SheetA"].suffix == ".toon"
    assert toon_paths["SheetA"].exists()


def test_merged_cells_empty_is_omitted_in_sheet_json() -> None:
    sheet = SheetData(rows=[CellRow(r=1, c={"0": "v"})], merged_cells=[])
    payload = sheet.to_json()
    assert "merged_cells" not in payload
