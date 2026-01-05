from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch

from exstruct import (
    DestinationOptions,
    ExStructEngine,
    OutputOptions,
    StructOptions,
    export_auto_page_breaks,
)
from exstruct.core.backends.com_backend import ComBackend
from exstruct.models import PrintArea, SheetData, WorkbookData


def test_extract_passes_auto_page_break_flag(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    called: dict[str, object] = {}

    def fake_extract(
        path: Path,
        mode: str,
        include_cell_links: bool = False,
        include_print_areas: bool = True,
        include_auto_page_breaks: bool = False,
        include_colors_map: bool = False,
        include_default_background: bool = False,
        ignore_colors: set[str] | None = None,
        include_merged_cells: bool | None = None,
    ) -> WorkbookData:
        called["include_auto_page_breaks"] = include_auto_page_breaks
        return WorkbookData(book_name=path.name, sheets={})

    monkeypatch.setattr("exstruct.engine.extract_workbook", fake_extract)

    engine = ExStructEngine(
        options=StructOptions(mode="standard"),
        output=OutputOptions(
            destinations=DestinationOptions(auto_page_breaks_dir=tmp_path / "out")
        ),
    )
    engine.extract(tmp_path / "book.xlsx")

    assert called["include_auto_page_breaks"] is True


def test_export_auto_page_breaks_writes_files(tmp_path: Path) -> None:
    area = PrintArea(r1=1, c1=0, r2=2, c2=1)
    sheet = SheetData(auto_print_areas=[area])
    wb = WorkbookData(book_name="b.xlsx", sheets={"Sheet1": sheet})

    written = export_auto_page_breaks(wb, tmp_path, pretty=True)
    assert written
    assert (tmp_path / next(iter(written.values())).name).exists()


def test_export_auto_page_breaks_raises_when_empty(tmp_path: Path) -> None:
    wb = WorkbookData(book_name="b.xlsx", sheets={"Sheet1": SheetData()})
    try:
        export_auto_page_breaks(wb, tmp_path)
    except ValueError:
        return
    raise AssertionError(
        "export_auto_page_breaks should raise when no auto_print_areas"
    )


def test_com_backend_extract_auto_page_breaks_handles_failure() -> None:
    class _FailingSheetApi:
        @property
        def DisplayPageBreaks(self) -> bool:
            raise RuntimeError("boom")

    class _FailingSheet:
        name = "Sheet1"
        api = _FailingSheetApi()

    class _DummyWorkbook:
        sheets = [_FailingSheet()]

    backend = ComBackend(_DummyWorkbook())
    assert backend.extract_auto_page_breaks() == {}
