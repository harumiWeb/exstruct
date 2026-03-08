"""Tests for auto page-break extraction and export behavior."""

from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch
import pytest

from exstruct import (
    ConfigError,
    DestinationOptions,
    ExStructEngine,
    FilterOptions,
    OutputOptions,
    StructOptions,
    export_auto_page_breaks,
)
from exstruct.core.backends.com_backend import ComBackend
from exstruct.models import PrintArea, SheetData, WorkbookData


def test_extract_passes_auto_page_break_flag(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """
    Verify that extract_workbook is invoked with include_auto_page_breaks set to True.

    Creates a fake extractor that captures the include_auto_page_breaks argument, replaces
    exstruct.engine.extract_workbook with it, runs ExStructEngine.extract against a dummy
    workbook path configured to export auto page breaks, and asserts the captured flag is True.
    """
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
        include_formulas_map: bool | None = None,
        include_merged_cells: bool | None = None,
        include_merged_values_in_rows: bool = True,
    ) -> WorkbookData:
        """
        Test stub for workbook extraction that records the auto page breaks flag.

        This fake extractor captures the value of `include_auto_page_breaks` in the outer
        `called` mapping and returns a minimal `WorkbookData` with `book_name` set to
        the provided path's filename and an empty `sheets` mapping.

        Parameters:
            path (Path): Filesystem path used to derive the returned `WorkbookData.book_name`.
            include_auto_page_breaks (bool): Flag whose value is written to `called["include_auto_page_breaks"]`.

        Returns:
            WorkbookData: A minimal workbook data object with `book_name` set to `path.name` and no sheets.
        """
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


def test_extract_rejects_auto_page_break_flag_in_libreoffice_mode(
    tmp_path: Path,
) -> None:
    """Verify that extract rejects auto page break flag in LibreOffice mode."""

    engine = ExStructEngine(
        options=StructOptions(mode="libreoffice"),
        output=OutputOptions(
            filters=FilterOptions(include_auto_print_areas=True),
        ),
    )

    with pytest.raises(ConfigError, match="does not support auto page-break export"):
        engine.extract(tmp_path / "book.xlsx")


def test_process_passes_auto_page_break_flag_from_per_call_destination(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that process-time auto-page-break overrides reach extraction."""

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
        include_formulas_map: bool | None = None,
        include_merged_cells: bool | None = None,
        include_merged_values_in_rows: bool = True,
    ) -> WorkbookData:
        _ = (
            include_cell_links,
            include_print_areas,
            include_colors_map,
            include_default_background,
            ignore_colors,
            include_formulas_map,
            include_merged_cells,
            include_merged_values_in_rows,
        )
        called["path"] = path
        called["mode"] = mode
        called["include_auto_page_breaks"] = include_auto_page_breaks
        return WorkbookData(book_name=path.name, sheets={})

    monkeypatch.setattr("exstruct.engine.extract_workbook", fake_extract)

    engine = ExStructEngine(options=StructOptions(mode="standard"))
    engine.process(
        tmp_path / "book.xlsx",
        output_path=tmp_path / "out.json",
        auto_page_breaks_dir=tmp_path / "auto",
    )

    assert called["path"] == tmp_path / "book.xlsx"
    assert called["mode"] == "standard"
    assert called["include_auto_page_breaks"] is True


def test_process_uses_engine_default_auto_page_break_destination_without_mutation(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that process forwards engine defaults without mutating engine state."""

    calls: dict[str, object] = {}
    default_dir = tmp_path / "auto-default"

    def fake_extract(
        self: ExStructEngine,
        path: Path,
        *,
        mode: str | None = None,
    ) -> WorkbookData:
        calls["extract_path"] = path
        calls["extract_mode"] = mode
        calls["engine_dir_inside_extract"] = (
            self.output.destinations.auto_page_breaks_dir
        )
        return WorkbookData(book_name=path.name, sheets={})

    def fake_export(
        self: ExStructEngine,
        wb: WorkbookData,
        *,
        auto_page_breaks_dir: Path | None = None,
        **kwargs: object,
    ) -> None:
        _ = kwargs
        _engine = self
        _ = _engine
        calls["export_book_name"] = wb.book_name
        calls["export_dir"] = auto_page_breaks_dir

    monkeypatch.setattr(ExStructEngine, "extract", fake_extract, raising=True)
    monkeypatch.setattr(ExStructEngine, "export", fake_export, raising=True)

    engine = ExStructEngine(
        output=OutputOptions(
            destinations=DestinationOptions(auto_page_breaks_dir=default_dir)
        )
    )
    input_path = tmp_path / "book.xlsx"
    input_path.write_text("", encoding="utf-8")

    engine.process(input_path, output_path=tmp_path / "out.json")

    assert calls["extract_path"] == input_path
    assert calls["extract_mode"] == "standard"
    assert calls["engine_dir_inside_extract"] == default_dir
    assert calls["export_dir"] == default_dir
    assert calls["export_book_name"] == input_path.name
    assert engine.output.destinations.auto_page_breaks_dir == default_dir


def test_export_auto_page_breaks_writes_files(tmp_path: Path) -> None:
    """Verify that export auto page breaks writes files."""

    area = PrintArea(r1=1, c1=0, r2=2, c2=1)
    sheet = SheetData(auto_print_areas=[area])
    wb = WorkbookData(book_name="b.xlsx", sheets={"Sheet1": sheet})

    written = export_auto_page_breaks(wb, tmp_path, pretty=True)
    assert written
    assert (tmp_path / next(iter(written.values())).name).exists()


def test_export_auto_page_breaks_raises_when_empty(tmp_path: Path) -> None:
    """Verify that export auto page breaks raises when empty."""

    wb = WorkbookData(book_name="b.xlsx", sheets={"Sheet1": SheetData()})
    try:
        export_auto_page_breaks(wb, tmp_path)
    except ValueError:
        return
    raise AssertionError(
        "export_auto_page_breaks should raise when no auto_print_areas"
    )


def test_com_backend_extract_auto_page_breaks_handles_failure() -> None:
    """Verify that COM backend extract auto page breaks handles failure."""

    class _FailingSheetApi:
        """Failing sheet API double used in tests."""

        @property
        def DisplayPageBreaks(self) -> bool:
            """Raise when page-break display state is requested."""

            raise RuntimeError("boom")

    class _FailingSheet:
        """Worksheet double that exposes the failing sheet API."""

        name = "Sheet1"
        api = _FailingSheetApi()

    class _DummyWorkbook:
        """Workbook double that wraps the failing sheet test case."""

        sheets = [_FailingSheet()]

    backend = ComBackend(_DummyWorkbook())
    assert backend.extract_auto_page_breaks() == {}
