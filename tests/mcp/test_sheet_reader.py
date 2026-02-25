from __future__ import annotations

from collections.abc import Mapping
import json
from pathlib import Path

import pytest

from exstruct.mcp.sheet_reader import (
    ReadCellsRequest,
    ReadFormulasRequest,
    ReadRangeRequest,
    read_cells,
    read_formulas,
    read_range,
)


def _write_json(path: Path, data: Mapping[str, object]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def test_read_range_with_a1_range(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {
            "Data": {
                "rows": [
                    {"r": 1, "c": {"0": "A1", "1": "B1"}},
                    {"r": 2, "c": {"0": "A2"}},
                ]
            }
        },
    }
    out = tmp_path / "out.json"
    _write_json(out, data)

    request = ReadRangeRequest(out_path=out, sheet="Data", range="A1:B2")
    result = read_range(request)

    assert result.range == "A1:B2"
    assert [item.cell for item in result.cells] == ["A1", "B1", "A2", "B2"]
    assert [item.value for item in result.cells] == ["A1", "B1", "A2", None]


def test_read_range_rejects_invalid_range(tmp_path: Path) -> None:
    out = tmp_path / "out.json"
    _write_json(out, {"book_name": "book", "sheets": {"Data": {"rows": []}}})
    request = ReadRangeRequest(out_path=out, sheet="Data", range="1A:B2")
    with pytest.raises(ValueError, match="Invalid A1 cell address"):
        read_range(request)


def test_read_range_rejects_too_large_range(tmp_path: Path) -> None:
    out = tmp_path / "out.json"
    _write_json(out, {"book_name": "book", "sheets": {"Data": {"rows": []}}})
    request = ReadRangeRequest(out_path=out, sheet="Data", range="A1:C3", max_cells=8)
    with pytest.raises(ValueError, match="Requested range exceeds max_cells"):
        read_range(request)


def test_read_cells_keeps_input_order_and_reports_missing(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {
            "Data": {
                "rows": [
                    {"r": 1, "c": {"0": "head"}},
                    {"r": 98, "c": {"9": 12345}},
                ],
                "formulas_map": {"=SUM(D124:F124)*G124*(1-H124)": [[124, 9]]},
            }
        },
    }
    out = tmp_path / "out.json"
    _write_json(out, data)

    request = ReadCellsRequest(
        out_path=out,
        sheet="Data",
        addresses=["J124", "B2", "J98", "A1"],
        include_formulas=True,
    )
    result = read_cells(request)

    assert [item.cell for item in result.cells] == ["J124", "B2", "J98", "A1"]
    assert result.missing_cells == ["B2"]
    assert result.cells[0].formula == "=SUM(D124:F124)*G124*(1-H124)"
    assert result.cells[2].value == 12345
    assert result.cells[3].value == "head"


def test_read_formulas_with_range_and_values_alpha_col(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {
            "Data": {
                "rows": [
                    {"r": 1, "c": {"A": "header"}},
                    {"r": 98, "c": {"J": 98765}},
                ],
                "formulas_map": {
                    "=SUM(D98:F98)*G98*(1-H98)": [[98, 9]],
                    "=A1": [[1, 0]],
                },
            }
        },
    }
    out = tmp_path / "out.json"
    _write_json(out, data)

    request = ReadFormulasRequest(
        out_path=out,
        sheet="Data",
        range="J2:J200",
        include_values=True,
    )
    result = read_formulas(request)

    assert result.range == "J2:J200"
    assert len(result.formulas) == 1
    assert result.formulas[0].cell == "J98"
    assert result.formulas[0].formula == "=SUM(D98:F98)*G98*(1-H98)"
    assert result.formulas[0].value == 98765


def test_read_formulas_warns_when_formulas_map_missing(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Data": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)

    request = ReadFormulasRequest(out_path=out, sheet="Data")
    result = read_formulas(request)

    assert result.formulas == []
    assert any("mode='verbose'" in warning for warning in result.warnings)


def test_read_range_warns_when_include_formulas_without_map(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Data": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)

    request = ReadRangeRequest(
        out_path=out, sheet="Data", range="A1:A1", include_formulas=True
    )
    result = read_range(request)

    assert len(result.cells) == 1
    assert result.cells[0].cell == "A1"
    assert any("mode='verbose'" in warning for warning in result.warnings)


def test_read_cells_requires_sheet_for_multi_sheet_payload(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"A": {"rows": []}, "B": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)

    request = ReadCellsRequest(out_path=out, addresses=["A1"])
    with pytest.raises(ValueError, match=r"Available sheets: A, B"):
        read_cells(request)


def test_read_range_excludes_empty_cells_when_include_empty_false(
    tmp_path: Path,
) -> None:
    data = {"book_name": "book", "sheets": {"Data": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)

    request = ReadRangeRequest(
        out_path=out,
        sheet="Data",
        range="A1:B1",
        include_empty=False,
    )
    result = read_range(request)
    assert result.cells == []


def test_read_range_rejects_missing_output_file(tmp_path: Path) -> None:
    request = ReadRangeRequest(
        out_path=tmp_path / "missing.json", sheet="Data", range="A1"
    )
    with pytest.raises(FileNotFoundError, match="Output file not found"):
        read_range(request)


def test_sheet_reader_private_helpers_cover_invalid_payload_paths() -> None:
    from exstruct.mcp import sheet_reader

    with pytest.raises(ValueError, match="expected object at root"):
        sheet_reader._parse_json("[]")

    with pytest.raises(ValueError, match="sheets is not a mapping"):
        sheet_reader._select_sheet({"sheets": []}, None)
    with pytest.raises(ValueError, match="sheet payload is not an object"):
        sheet_reader._select_sheet({"sheets": {"Only": []}}, None)
    with pytest.raises(ValueError, match="Invalid A1 range"):
        sheet_reader._parse_range("B2:A1")


def test_sheet_reader_private_parsers_and_normalizers() -> None:
    from exstruct.mcp import sheet_reader

    assert sheet_reader._build_value_map({"rows": "not-list"}) == {}
    value_map = sheet_reader._build_value_map(
        {
            "rows": [
                {"r": 1, "c": {"0": "v0", "AA": "vaa", "-1": "skip"}},
                {"r": 0, "c": {"0": "skip"}},
                {"r": 2, "c": ["skip"]},
            ]
        }
    )
    assert value_map[(1, 1)] == "v0"
    assert value_map[(1, 27)] == "vaa"

    formula_map, has_formula = sheet_reader._build_formula_map(
        {
            "formulas_map": {
                1: [[1, 0]],
                "=OK": [[1, 0], [0, 0], [1], "bad"],
            }
        }
    )
    assert has_formula is True
    assert formula_map == {(1, 1): "=OK"}

    assert sheet_reader._parse_formula_position("bad") is None
    assert sheet_reader._parse_formula_position([1]) is None
    assert sheet_reader._parse_formula_position([1, "x"]) is None
    assert sheet_reader._parse_formula_position([0, 0]) is None
    assert sheet_reader._parse_formula_position([1, -1]) is None

    assert sheet_reader._parse_col_key("-1") is None
    assert sheet_reader._parse_col_key("3") == 4
    assert sheet_reader._parse_col_key("AB") == 28
    assert sheet_reader._parse_col_key("A1") is None

    with pytest.raises(ValueError, match="Invalid column index"):
        sheet_reader._col_to_alpha(0)
    with pytest.raises(ValueError, match="Invalid column label"):
        sheet_reader._alpha_to_col("A!")

    assert sheet_reader._normalize_scalar({"a": 1}) == "{'a': 1}"
    assert sheet_reader._as_optional_str(10) == "10"
