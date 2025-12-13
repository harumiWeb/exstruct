from __future__ import annotations

import importlib
from pathlib import Path

import pytest

from exstruct import export_auto_page_breaks
from exstruct.errors import (
    MissingDependencyError,
    OutputError,
    PrintAreaError,
    SerializationError,
)
from exstruct.io import save_as_json, save_as_yaml, serialize_workbook
from exstruct.models import SheetData, WorkbookData


def _minimal_workbook() -> WorkbookData:
    """Create a minimal workbook for error-path testing."""
    return WorkbookData(book_name="book.xlsx", sheets={"Sheet1": SheetData()})


def test_serialize_workbook_unsupported_format() -> None:
    """Unsupported formats should raise SerializationError."""
    workbook = _minimal_workbook()
    with pytest.raises(SerializationError):
        serialize_workbook(workbook, fmt="invalid")  # type: ignore[arg-type]


def test_save_as_yaml_missing_dependency(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """YAML dependency missing should raise MissingDependencyError."""
    original_import = importlib.import_module

    def _fake_import(name: str, package: str | None = None) -> object:
        if name == "yaml":
            raise ImportError("yaml not installed")
        return original_import(name, package=package)

    monkeypatch.setattr(importlib, "import_module", _fake_import)
    workbook = _minimal_workbook()
    with pytest.raises(MissingDependencyError):
        save_as_yaml(workbook, tmp_path / "out.yaml")


def test_save_as_json_write_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """IO failures should surface as OutputError."""
    workbook = _minimal_workbook()

    def _fail_write(self: Path, *args: object, **kwargs: object) -> None:
        raise OSError("disk full")

    monkeypatch.setattr(Path, "write_text", _fail_write)
    with pytest.raises(OutputError):
        save_as_json(workbook, tmp_path / "out.json")


def test_export_auto_page_breaks_raises_print_area_error(tmp_path: Path) -> None:
    """Auto page-break export without data should raise PrintAreaError."""
    workbook = _minimal_workbook()
    with pytest.raises(PrintAreaError):
        export_auto_page_breaks(workbook, tmp_path)
