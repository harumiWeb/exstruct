from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path
from typing import cast

import pytest

from exstruct.core import workbook


def test_openpyxl_workbook_close_error_is_suppressed(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    class _DummyWorkbook:
        def close(self) -> None:
            raise RuntimeError("close failed")

    dummy = _DummyWorkbook()

    def _fake_load_workbook(*_args: object, **_kwargs: object) -> _DummyWorkbook:
        return dummy

    monkeypatch.setattr(workbook, "load_workbook", _fake_load_workbook)

    with workbook.openpyxl_workbook(
        tmp_path / "book.xlsx", data_only=True, read_only=False
    ) as wb:
        assert wb is dummy


def test_xlwings_workbook_leaves_existing_workbook_and_app_open(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    class _DummyBook:
        close_calls = 0
        fullname = str(tmp_path / "book.xlsx")

        def close(self) -> None:
            self.close_calls += 1

    dummy = _DummyBook()

    class _DummyApp:
        books = [dummy]
        quit_calls = 0

        def quit(self) -> None:
            self.quit_calls += 1

    app = _DummyApp()

    def _boom(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("xlwings App should not be created when existing")

    monkeypatch.setattr(workbook.xw, "apps", [app])
    monkeypatch.setattr(workbook.xw, "App", _boom)

    with workbook.xlwings_workbook(tmp_path / "book.xlsx") as wb:
        assert wb is dummy

    assert dummy.close_calls == 0
    assert app.quit_calls == 0


def test_xlwings_workbook_quits_app_when_open_fails(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    class _DummyBooks:
        def open(self, _path: str) -> None:
            raise RuntimeError("open failed")

    class _DummyApp:
        books = _DummyBooks()
        quit_calls = 0
        kill_calls = 0

        def quit(self) -> None:
            self.quit_calls += 1

        def kill(self) -> None:
            self.kill_calls += 1

    app = _DummyApp()
    monkeypatch.setattr(workbook, "_find_open_workbook", lambda _path: None)
    monkeypatch.setattr(workbook.xw, "App", lambda **_kwargs: app)

    with pytest.raises(RuntimeError, match="open failed"):
        with workbook.xlwings_workbook(tmp_path / "book.xlsx"):
            pytest.fail("open failure must prevent yielding")

    assert app.quit_calls == 1
    assert app.kill_calls == 0


def test_xlwings_workbook_closes_owned_book_after_extraction_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    class _DummyBook:
        close_calls = 0

        def close(self) -> None:
            self.close_calls += 1

    class _DummyBooks:
        def __init__(self, book: _DummyBook) -> None:
            self._book = book

        def open(self, _path: str) -> _DummyBook:
            return self._book

    class _DummyApp:
        def __init__(self, book: _DummyBook) -> None:
            self.books = _DummyBooks(book)
            self.quit_calls = 0

        def quit(self) -> None:
            self.quit_calls += 1

    book = _DummyBook()
    app = _DummyApp(book)
    monkeypatch.setattr(workbook, "_find_open_workbook", lambda _path: None)
    monkeypatch.setattr(workbook.xw, "App", lambda **_kwargs: app)

    with pytest.raises(RuntimeError, match="extraction failed"):
        with workbook.xlwings_workbook(tmp_path / "book.xlsx"):
            raise RuntimeError("extraction failed")

    assert book.close_calls == 1
    assert app.quit_calls == 1


@pytest.mark.parametrize("close_fails", [False, True])
def test_xlwings_workbook_kills_owned_app_when_quit_fails(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, close_fails: bool
) -> None:
    class _DummyBook:
        def close(self) -> None:
            if close_fails:
                raise RuntimeError("close failed")

    class _DummyBooks:
        def open(self, _path: str) -> _DummyBook:
            return _DummyBook()

    class _DummyApp:
        books = _DummyBooks()
        quit_calls = 0
        kill_calls = 0

        def quit(self) -> None:
            self.quit_calls += 1
            raise RuntimeError("quit failed")

        def kill(self) -> None:
            self.kill_calls += 1

    app = _DummyApp()
    monkeypatch.setattr(workbook, "_find_open_workbook", lambda _path: None)
    monkeypatch.setattr(workbook.xw, "App", lambda **_kwargs: app)

    with workbook.xlwings_workbook(tmp_path / "book.xlsx"):
        pass

    assert app.quit_calls == 1
    assert app.kill_calls == 1


def test_find_open_workbook_handles_fullname_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class _BadBook:
        @property
        def fullname(self) -> str:
            raise RuntimeError("boom")

    class _DummyApp:
        books = [_BadBook()]

    monkeypatch.setattr(workbook.xw, "apps", [_DummyApp()])
    assert workbook._find_open_workbook(Path("book.xlsx")) is None


def test_find_open_workbook_handles_resolve_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class _DummyPath:
        def __init__(self, value: str) -> None:
            self._value = value

        def resolve(self) -> _DummyPath:
            if self._value == "bad":
                raise RuntimeError("resolve failed")
            return self

        def __eq__(self, other: object) -> bool:
            return isinstance(other, _DummyPath) and self._value == other._value

    class _DummyBook:
        fullname = "bad"

    class _DummyApp:
        books = [_DummyBook()]

    monkeypatch.setattr(workbook, "Path", _DummyPath)
    monkeypatch.setattr(workbook.xw, "apps", [_DummyApp()])

    file_path = _DummyPath("good")
    assert workbook._find_open_workbook(cast(Path, file_path)) is None


def test_find_open_workbook_returns_none_on_iter_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class _BadApps:
        def __iter__(self) -> Iterator[object]:
            raise RuntimeError("apps failure")

    monkeypatch.setattr(workbook.xw, "apps", _BadApps())
    assert workbook._find_open_workbook(Path("book.xlsx")) is None
