from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch

from exstruct.core.workbook import openpyxl_workbook


def test_openpyxl_workbook_closes(monkeypatch: MonkeyPatch, tmp_path: Path) -> None:
    calls: dict[str, int] = {"close": 0}

    class DummyWorkbook:
        def close(self) -> None:
            calls["close"] += 1

    def fake_load_workbook(
        path: Path, *, data_only: bool, read_only: bool
    ) -> DummyWorkbook:
        return DummyWorkbook()

    monkeypatch.setattr("exstruct.core.workbook.load_workbook", fake_load_workbook)

    with openpyxl_workbook(tmp_path / "book.xlsx", data_only=True, read_only=True):
        pass

    assert calls["close"] == 1


def test_openpyxl_workbook_sets_warning_filters(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    calls: list[tuple[str, str, type[Warning], str]] = []

    def fake_filterwarnings(
        action: str,
        message: str = "",
        category: type[Warning] = Warning,
        module: str = "",
        **_kwargs: object,
    ) -> None:
        calls.append((action, message, category, module))

    class DummyWorkbook:
        def close(self) -> None:
            pass

    def fake_load_workbook(
        path: Path, *, data_only: bool, read_only: bool
    ) -> DummyWorkbook:
        return DummyWorkbook()

    monkeypatch.setattr(
        "exstruct.core.workbook.warnings.filterwarnings", fake_filterwarnings
    )
    monkeypatch.setattr("exstruct.core.workbook.load_workbook", fake_load_workbook)

    with openpyxl_workbook(tmp_path / "book.xlsx", data_only=True, read_only=True):
        pass

    expected_messages = {
        "Unknown extension is not supported and will be removed",
        "Conditional Formatting extension is not supported and will be removed",
        "Cannot parse header or footer so it will be ignored",
    }
    recorded = {message for _action, message, _category, _module in calls}
    assert expected_messages.issubset(recorded)
