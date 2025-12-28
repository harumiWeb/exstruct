from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch
from openpyxl import Workbook

from exstruct.core.pipeline import resolve_extraction_inputs, run_extraction_pipeline
from exstruct.errors import FallbackReason


def _make_basic_book(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "v1"
    wb.save(path)
    wb.close()


def test_pipeline_fallback_skip_com_tests(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "book.xlsx"
    _make_basic_book(path)
    monkeypatch.setenv("SKIP_COM_TESTS", "1")

    inputs = resolve_extraction_inputs(
        path,
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
    )
    result = run_extraction_pipeline(inputs)

    assert result.state.fallback_reason == FallbackReason.SKIP_COM_TESTS
    assert result.state.com_attempted is False
    sheet = result.workbook.sheets["Sheet1"]
    assert sheet.shapes == []
    assert sheet.charts == []
    assert sheet.rows


def test_pipeline_fallback_com_unavailable(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "book.xlsx"
    _make_basic_book(path)
    monkeypatch.delenv("SKIP_COM_TESTS", raising=False)

    def _raise(*_args: object, **_kwargs: object) -> None:
        raise RuntimeError("no COM")

    monkeypatch.setattr("exstruct.core.pipeline.xlwings_workbook", _raise)

    inputs = resolve_extraction_inputs(
        path,
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
    )
    result = run_extraction_pipeline(inputs)

    assert result.state.fallback_reason == FallbackReason.COM_UNAVAILABLE
    assert result.state.com_attempted is False
    sheet = result.workbook.sheets["Sheet1"]
    assert sheet.shapes == []
    assert sheet.charts == []
    assert sheet.rows


def test_pipeline_fallback_com_pipeline_failed(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "book.xlsx"
    _make_basic_book(path)
    monkeypatch.delenv("SKIP_COM_TESTS", raising=False)

    @contextmanager
    def _dummy_workbook(_path: Path) -> Iterator[object]:
        yield object()

    def _raise(
        *_args: object,
        **_kwargs: object,
    ) -> None:
        raise RuntimeError("pipeline failed")

    monkeypatch.setattr("exstruct.core.pipeline.xlwings_workbook", _dummy_workbook)
    monkeypatch.setattr("exstruct.core.pipeline.run_com_pipeline", _raise)

    inputs = resolve_extraction_inputs(
        path,
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
    )
    result = run_extraction_pipeline(inputs)

    assert result.state.fallback_reason == FallbackReason.COM_PIPELINE_FAILED
    assert result.state.com_attempted is True
    sheet = result.workbook.sheets["Sheet1"]
    assert sheet.shapes == []
    assert sheet.charts == []
    assert sheet.rows
