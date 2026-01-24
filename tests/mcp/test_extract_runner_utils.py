from __future__ import annotations

import builtins
from collections.abc import Mapping, Sequence
from pathlib import Path

import pytest

from exstruct.mcp import extract_runner
from exstruct.mcp.io import PathPolicy


def test_resolve_input_path_missing(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError):
        extract_runner._resolve_input_path(tmp_path / "missing.xlsx", policy=None)


def test_resolve_input_path_rejects_directory(tmp_path: Path) -> None:
    path = tmp_path / "input.xlsx"
    path.mkdir()
    with pytest.raises(ValueError):
        extract_runner._resolve_input_path(path, policy=None)


def test_format_suffix() -> None:
    assert extract_runner._format_suffix("yml") == ".yml"
    assert extract_runner._format_suffix("json") == ".json"


def test_normalize_output_name(tmp_path: Path) -> None:
    input_path = tmp_path / "book.xlsx"
    assert (
        extract_runner._normalize_output_name(input_path, None, ".json") == "book.json"
    )
    assert (
        extract_runner._normalize_output_name(input_path, "out", ".json") == "out.json"
    )
    assert (
        extract_runner._normalize_output_name(input_path, "out.yaml", ".json")
        == "out.yaml"
    )


def test_resolve_output_path_denies_outside_root(tmp_path: Path) -> None:
    policy = PathPolicy(root=tmp_path)
    input_path = tmp_path / "book.xlsx"
    input_path.write_text("x", encoding="utf-8")
    outside = tmp_path.parent
    with pytest.raises(ValueError):
        extract_runner._resolve_output_path(
            input_path,
            "json",
            out_dir=outside,
            out_name=None,
            policy=policy,
        )


def test_try_read_workbook_meta(tmp_path: Path) -> None:
    from openpyxl import Workbook

    path = tmp_path / "book.xlsx"
    wb = Workbook()
    wb.active.title = "Sheet1"
    wb.save(path)
    meta, warnings = extract_runner._try_read_workbook_meta(path)
    assert meta is not None
    assert meta.sheet_count == 1
    assert meta.sheet_names == ["Sheet1"]
    assert warnings == []


def test_run_extract_skips_when_output_exists(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "input.json"
    input_path.write_text("x", encoding="utf-8")
    output_path.write_text("y", encoding="utf-8")

    def _raise(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("process_excel should not be called")

    monkeypatch.setattr(extract_runner, "process_excel", _raise)
    request = extract_runner.ExtractRequest(
        xlsx_path=input_path,
        on_conflict="skip",
        format="json",
    )
    result = extract_runner.run_extract(request)
    assert result.workbook_meta is None
    assert any("skipping write" in warning for warning in result.warnings)


def test_run_extract_creates_output_dir(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    input_path = tmp_path / "input.xlsx"
    input_path.write_text("x", encoding="utf-8")
    out_dir = tmp_path / "nested" / "out"

    def _noop(*_args: object, **_kwargs: object) -> None:
        return None

    monkeypatch.setattr(extract_runner, "process_excel", _noop)
    monkeypatch.setattr(extract_runner, "_try_read_workbook_meta", lambda _: (None, []))
    request = extract_runner.ExtractRequest(
        xlsx_path=input_path,
        out_dir=out_dir,
        format="json",
    )
    extract_runner.run_extract(request)
    assert out_dir.exists()


def test_run_extract_applies_options(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    input_path = tmp_path / "input.xlsx"
    input_path.write_text("x", encoding="utf-8")
    capture: dict[str, object] = {}

    def _capture(*_args: object, **kwargs: object) -> None:
        capture.update(kwargs)

    monkeypatch.setattr(extract_runner, "process_excel", _capture)
    monkeypatch.setattr(extract_runner, "_try_read_workbook_meta", lambda _: (None, []))

    options = extract_runner.ExtractOptions(
        pretty=True,
        indent=2,
        sheets_dir=tmp_path / "sheets",
        print_areas_dir=tmp_path / "print_areas",
    )
    request = extract_runner.ExtractRequest(
        xlsx_path=input_path,
        out_dir=tmp_path,
        format="json",
        options=options,
    )
    extract_runner.run_extract(request)
    assert capture["pretty"] is True
    assert capture["indent"] == 2
    assert capture["sheets_dir"] == options.sheets_dir
    assert capture["print_areas_dir"] == options.print_areas_dir


def test_try_read_workbook_meta_import_error(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "missing.xlsx"

    def _import(
        name: str,
        globals_: Mapping[str, object] | None = None,
        locals_: Mapping[str, object] | None = None,
        fromlist: Sequence[str] = (),
        level: int = 0,
    ) -> object:
        if name == "openpyxl":
            raise ImportError("missing")
        return builtins.__import__(name, globals_, locals_, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _import)
    meta, warnings = extract_runner._try_read_workbook_meta(path)
    assert meta is None
    assert any("openpyxl is not available" in warning for warning in warnings)


def test_try_read_workbook_meta_load_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "broken.xlsx"
    path.write_text("x", encoding="utf-8")

    def _raise(*_args: object, **_kwargs: object) -> None:
        raise ValueError("boom")

    import openpyxl

    monkeypatch.setattr(openpyxl, "load_workbook", _raise)
    meta, warnings = extract_runner._try_read_workbook_meta(path)
    assert meta is None
    assert any("Failed to read workbook metadata" in warning for warning in warnings)
