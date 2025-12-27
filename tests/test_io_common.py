from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch

from exstruct.io import save_sheets
from exstruct.models import SheetData, WorkbookData


def test_save_sheets_accepts_yml(monkeypatch: MonkeyPatch, tmp_path: Path) -> None:
    class DummyYaml:
        @staticmethod
        def safe_dump(
            payload: object, *, allow_unicode: bool, sort_keys: bool, indent: int
        ) -> str:
            return "dummy-yaml"

    monkeypatch.setattr("exstruct.io.serialize._require_yaml", lambda: DummyYaml())

    wb = WorkbookData(book_name="book.xlsx", sheets={"Sheet1": SheetData()})
    written = save_sheets(wb, tmp_path, fmt="yml")
    path = next(iter(written.values()))
    assert path.suffix == ".yaml"
    assert path.read_text(encoding="utf-8") == "dummy-yaml"
