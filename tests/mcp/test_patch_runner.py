from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from pydantic import ValidationError
import pytest

from exstruct.cli.availability import ComAvailability
from exstruct.mcp import patch_runner
from exstruct.mcp.io import PathPolicy
from exstruct.mcp.patch_runner import PatchOp, PatchRequest, run_patch


def _create_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "old"
    sheet["B1"] = 1
    workbook.save(path)
    workbook.close()


def _disable_com(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        patch_runner,
        "get_com_availability",
        lambda: ComAvailability(available=False, reason="test"),
    )


def test_run_patch_set_value_and_formula(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _disable_com(monkeypatch)
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    ops = [
        PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="new"),
        PatchOp(op="set_formula", sheet="Sheet1", cell="B1", formula="=SUM(1,1)"),
    ]
    request = PatchRequest(xlsx_path=input_path, ops=ops, on_conflict="rename")
    result = run_patch(request, policy=PathPolicy(root=tmp_path))
    assert Path(result.out_path).exists()
    workbook = load_workbook(result.out_path)
    try:
        sheet = workbook["Sheet1"]
        assert sheet["A1"].value == "new"
        formula_value = sheet["B1"].value
        if isinstance(formula_value, str) and not formula_value.startswith("="):
            formula_value = f"={formula_value}"
        assert formula_value == "=SUM(1,1)"
    finally:
        workbook.close()
    assert len(result.patch_diff) == 2
    assert result.patch_diff[0].after is not None


def test_run_patch_add_sheet_and_set_value(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _disable_com(monkeypatch)
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    ops = [
        PatchOp(op="add_sheet", sheet="NewSheet"),
        PatchOp(op="set_value", sheet="NewSheet", cell="A1", value="ok"),
    ]
    request = PatchRequest(xlsx_path=input_path, ops=ops, on_conflict="rename")
    result = run_patch(request, policy=PathPolicy(root=tmp_path))
    workbook = load_workbook(result.out_path)
    try:
        assert "NewSheet" in workbook.sheetnames
        assert workbook["NewSheet"]["A1"].value == "ok"
    finally:
        workbook.close()


def test_run_patch_add_sheet_rejects_duplicate(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _disable_com(monkeypatch)
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    ops = [PatchOp(op="add_sheet", sheet="Sheet1")]
    request = PatchRequest(xlsx_path=input_path, ops=ops, on_conflict="rename")
    with pytest.raises(ValueError):
        run_patch(request, policy=PathPolicy(root=tmp_path))


def test_run_patch_rejects_value_starting_with_equal() -> None:
    with pytest.raises(ValidationError):
        PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="=SUM(1,1)")


def test_run_patch_rejects_formula_without_equal() -> None:
    with pytest.raises(ValidationError):
        PatchOp(op="set_formula", sheet="Sheet1", cell="A1", formula="SUM(1,1)")


def test_run_patch_rejects_path_outside_root(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _disable_com(monkeypatch)
    root = tmp_path / "root"
    root.mkdir()
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    ops = [PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="x")]
    request = PatchRequest(xlsx_path=input_path, ops=ops, on_conflict="rename")
    with pytest.raises(ValueError):
        run_patch(request, policy=PathPolicy(root=root))


def test_run_patch_conflict_rename(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _disable_com(monkeypatch)
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    default_out = tmp_path / "book_patched.xlsx"
    default_out.write_text("dummy", encoding="utf-8")
    ops = [PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="x")]
    request = PatchRequest(xlsx_path=input_path, ops=ops, on_conflict="rename")
    result = run_patch(request, policy=PathPolicy(root=tmp_path))
    assert result.out_path != str(default_out)
    assert Path(result.out_path).exists()


def test_run_patch_conflict_skip(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _disable_com(monkeypatch)
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    default_out = tmp_path / "book_patched.xlsx"
    default_out.write_text("dummy", encoding="utf-8")
    ops = [PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="x")]
    request = PatchRequest(xlsx_path=input_path, ops=ops, on_conflict="skip")
    result = run_patch(request, policy=PathPolicy(root=tmp_path))
    assert result.out_path == str(default_out)
    assert result.patch_diff == []
    assert any("skipping" in warning for warning in result.warnings)


def test_run_patch_conflict_overwrite(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _disable_com(monkeypatch)
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    default_out = tmp_path / "book_patched.xlsx"
    _create_workbook(default_out)
    ops = [PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="new")]
    request = PatchRequest(xlsx_path=input_path, ops=ops, on_conflict="overwrite")
    result = run_patch(request, policy=PathPolicy(root=tmp_path))
    assert result.out_path == str(default_out)
    workbook = load_workbook(result.out_path)
    try:
        assert workbook["Sheet1"]["A1"].value == "new"
    finally:
        workbook.close()


def test_run_patch_atomicity(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _disable_com(monkeypatch)
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    ops = [
        PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="x"),
        PatchOp(op="set_value", sheet="Missing", cell="A1", value="y"),
    ]
    request = PatchRequest(xlsx_path=input_path, ops=ops, on_conflict="rename")
    output_path = tmp_path / "book_patched.xlsx"
    with pytest.raises(ValueError):
        run_patch(request, policy=PathPolicy(root=tmp_path))
    assert not output_path.exists()
