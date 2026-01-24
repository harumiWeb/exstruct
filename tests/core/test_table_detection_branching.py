import builtins
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from types import ModuleType

from _pytest.monkeypatch import MonkeyPatch

from exstruct.core import cells


@dataclass(frozen=True)
class _DummyBook:
    fullname: str


@dataclass(frozen=True)
class _DummySheet:
    book: _DummyBook
    name: str


def test_detect_tables_xlsx_uses_openpyxl(monkeypatch: MonkeyPatch) -> None:
    """xlsx は openpyxl 経由で検出されることを確認する。"""
    sheet = _DummySheet(book=_DummyBook("C:/tmp/book.xlsx"), name="Sheet1")

    def _openpyxl_tables(_path: object, _name: str, **_kwargs: object) -> list[str]:
        return ["A1:B2"]

    def _com_tables(_sheet: object) -> list[str]:
        raise AssertionError("COM fallback should not be used for xlsx.")

    monkeypatch.setattr("exstruct.core.cells.detect_tables_openpyxl", _openpyxl_tables)
    monkeypatch.setattr("exstruct.core.cells.detect_tables_xlwings", _com_tables)

    assert cells.detect_tables(sheet) == ["A1:B2"]


def test_detect_tables_xls_falls_back_to_com(monkeypatch: MonkeyPatch) -> None:
    """xls は COM 経由にフォールバックすることを確認する。"""
    sheet = _DummySheet(book=_DummyBook("C:/tmp/book.xls"), name="Sheet1")

    def _com_tables(_sheet: object) -> list[str]:
        return ["C3:D4"]

    monkeypatch.setattr("exstruct.core.cells.detect_tables_xlwings", _com_tables)

    assert cells.detect_tables(sheet) == ["C3:D4"]


def test_detect_tables_openpyxl_missing_falls_back_to_com(
    monkeypatch: MonkeyPatch,
) -> None:
    """openpyxl 不在時は COM 経由にフォールバックすることを確認する。"""
    sheet = _DummySheet(book=_DummyBook("C:/tmp/book.xlsx"), name="Sheet1")
    original_import = builtins.__import__

    def _fake_import(
        name: str,
        globals_: Mapping[str, object] | None = None,
        locals_: Mapping[str, object] | None = None,
        fromlist: Sequence[str] | None = (),
        level: int = 0,
    ) -> ModuleType:
        if name == "openpyxl":
            raise ImportError("openpyxl missing")
        fromlist_seq: Sequence[str] = () if fromlist is None else fromlist
        return original_import(name, globals_, locals_, fromlist_seq, level)

    def _com_tables(_sheet: object) -> list[str]:
        return ["E5:F6"]

    monkeypatch.setattr(builtins, "__import__", _fake_import)
    monkeypatch.setattr("exstruct.core.cells.detect_tables_xlwings", _com_tables)

    assert cells.detect_tables(sheet) == ["E5:F6"]
