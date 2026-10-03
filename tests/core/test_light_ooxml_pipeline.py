"""Public light backend selection, complete-output parity and compatibility (#150)."""

from dataclasses import replace
import hashlib
import logging
from pathlib import Path
from unittest.mock import Mock
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Border, Side
from openpyxl.worksheet.table import Table
import pytest

import exstruct
from exstruct.core import ooxml_session, pipeline, workbook
from exstruct.errors import FallbackReason
from exstruct.models import CellRow


def make_light_book(path: Path) -> None:
    """Combine sparse cells, tables, links, formulas, merges, print areas and a chart."""
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = "Data"
    sheet.append(["Name", "Count"])
    sheet.append(["item", 7])
    sheet["A2"].hyperlink = "https://example.com"
    sheet.add_table(Table(displayName="Items", ref="A1:B2"))
    sheet["D1"] = "merged"
    sheet.merge_cells("D1:E2")
    sheet["G1"] = "=SUM(B2)"
    sheet["A10000"] = "sparse"
    edge = Side(style="thin")
    for row in sheet.iter_rows(min_row=1, max_row=2, min_col=1, max_col=2):
        for cell in row:
            cell.border = Border(top=edge, bottom=edge, left=edge, right=edge)
    sheet.print_area = ["A1:B2", "D1:E2"]
    chart = BarChart()
    chart.add_data(
        Reference(sheet, min_col=2, min_row=1, max_row=2), titles_from_data=True
    )
    sheet.add_chart(chart, "I1")
    book.create_sheet("Empty")
    book.save(path)
    book.close()


def inputs_for(path: Path) -> pipeline.ExtractionInputs:
    return pipeline.resolve_extraction_inputs(
        path,
        mode="light",
        include_cell_links=True,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=True,
        include_merged_cells=True,
        include_merged_values_in_rows=True,
    )


@pytest.mark.parametrize("suffix", [".xlsx", ".xlsm"])
@pytest.mark.parametrize("merged_values", [False, True])
def test_full_light_output_matches_previous_pipeline(
    tmp_path: Path,
    suffix: str,
    merged_values: bool,
) -> None:
    path = tmp_path / f"book{suffix}"
    make_light_book(path)
    if suffix == ".xlsm":
        with ZipFile(path, "a") as archive:
            archive.writestr("xl/vbaProject.bin", b"opaque VBA fixture payload")
    before = hashlib.sha256(path.read_bytes()).digest()
    inputs = replace(inputs_for(path), include_merged_values_in_rows=merged_values)
    expected = pipeline._run_openpyxl_pipeline(inputs).workbook
    actual = pipeline.run_extraction_pipeline(inputs)
    assert actual.state.fallback_reason is None
    assert actual.workbook.model_dump() == expected.model_dump()
    assert actual.artifacts.openpyxl_session is None
    assert hashlib.sha256(path.read_bytes()).digest() == before


@pytest.mark.parametrize(
    "path", sorted((Path(__file__).resolve().parents[2] / "sample").rglob("*.xlsx"))
)
def test_tracked_sample_light_output_parity(path: Path) -> None:
    """Compare emitted models including diagrams against the previous implementation."""
    inputs = inputs_for(path)
    expected = pipeline._run_openpyxl_pipeline(inputs).workbook
    actual = pipeline.run_extraction_pipeline(inputs).workbook
    assert actual.model_dump() == expected.model_dump()


def test_public_extraction_owns_one_zip_and_closes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = tmp_path / "book.xlsx"
    make_light_book(path)
    opened: list[ZipFile] = []

    def open_zip(path: Path) -> ZipFile:
        archive = ZipFile(path)
        opened.append(archive)
        return archive

    monkeypatch.setattr(ooxml_session, "ZipFile", open_zip)
    result = exstruct.extract(path, mode="light")
    assert result.sheets["Data"].charts
    assert len(opened) == 1
    assert opened[0].fp is None


def test_failure_discards_partial_core_artifacts_and_logs_reason(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    path = tmp_path / "book.xlsx"
    make_light_book(path)
    inputs = inputs_for(path)
    expected = pipeline._run_openpyxl_pipeline(inputs).workbook
    original = ooxml_session.OoxmlExtractionSession.extract_cells
    closed: list[ooxml_session.OoxmlExtractionSession] = []
    original_close = ooxml_session.OoxmlExtractionSession.close

    def partial(
        session: ooxml_session.OoxmlExtractionSession, *, include_links: bool = False
    ) -> dict[str, list[CellRow]]:
        original(session, include_links=include_links)
        return {"WrongPartialSheet": [CellRow(r=1, c={"0": "discard me"})]}

    def close(session: ooxml_session.OoxmlExtractionSession) -> None:
        original_close(session)
        closed.append(session)

    monkeypatch.setattr(ooxml_session.OoxmlExtractionSession, "extract_cells", partial)
    monkeypatch.setattr(
        ooxml_session.OoxmlExtractionSession,
        "extract_formulas_map",
        Mock(side_effect=ValueError("unsupported fixture formula")),
    )
    monkeypatch.setattr(ooxml_session.OoxmlExtractionSession, "close", close)
    with caplog.at_level(logging.WARNING):
        actual = pipeline.run_extraction_pipeline(inputs)
    assert actual.workbook.model_dump() == expected.model_dump()
    assert actual.state.fallback_reason == FallbackReason.OOXML_COMPATIBILITY
    assert "[ooxml_compatibility]" in caplog.text
    assert "unsupported fixture formula" in caplog.text
    assert "restarting complete light extraction" in caplog.text
    assert closed and closed[0]._archive is not None
    assert closed[0]._archive.fp is None


def test_colors_opt_in_uses_complete_compatibility_pipeline(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    path = tmp_path / "book.xlsx"
    make_light_book(path)
    inputs = replace(inputs_for(path), include_colors_map=True)
    expected = pipeline._run_openpyxl_pipeline(inputs).workbook
    opener = Mock(
        side_effect=AssertionError(
            "unsupported options must be selected before ZIP access"
        )
    )
    monkeypatch.setattr(ooxml_session, "ZipFile", opener)
    actual = pipeline.run_extraction_pipeline(inputs)
    assert actual.workbook.model_dump() == expected.model_dump()
    assert actual.state.fallback_reason == FallbackReason.OOXML_COMPATIBILITY
    assert "colors_map" in caplog.text
    opener.assert_not_called()


def test_legacy_workbook_loader_override_is_honored_by_public_light(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = tmp_path / "book.xlsx"
    make_light_book(path)
    original = workbook.load_workbook

    def load(*args: object, **kwargs: object) -> Workbook:
        book: Workbook = original(*args, **kwargs)
        book["Data"]["A2"] = "legacy override"
        return book

    monkeypatch.setattr(workbook, "load_workbook", load)
    result = exstruct.extract(path, mode="light")
    assert result.sheets["Data"].rows[1].c["0"] == "legacy override"


@pytest.mark.parametrize("formula", ["SUM(B:B)", "A1!B1", "A1:B2!C3"])
def test_unsupported_shared_formula_uses_compatibility_without_losing_followers(
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
    formula: str,
) -> None:
    """A valid whole-column shared formula must survive capability fallback."""
    path = tmp_path / "shared.xlsx"
    make_light_book(path)
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts[
        "xl/worksheets/sheet1.xml"
    ] = f"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
      <row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2">{formula}</f><v>7</v></c><c r="B1"><v>7</v></c></row>
      <row r="2"><c r="A2"><f t="shared" si="0"/><v>7</v></c></row>
    </sheetData></worksheet>""".encode()
    with ZipFile(path, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    inputs = inputs_for(path)
    expected = pipeline._run_openpyxl_pipeline(inputs).workbook
    actual = pipeline.run_extraction_pipeline(inputs)
    assert actual.workbook.model_dump() == expected.model_dump()
    assert actual.state.fallback_reason == FallbackReason.OOXML_COMPATIBILITY
    assert "formula" in caplog.text.lower()
    assert (
        sum(
            len(positions)
            for positions in actual.workbook.sheets["Data"].formulas_map.values()
        )
        == 2
    )


def test_uppercase_elapsed_time_format_preserves_public_cell_and_merge_values(
    tmp_path: Path,
) -> None:
    path = tmp_path / "duration.xlsx"
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = "Duration"
    sheet["A1"] = 1.5
    sheet["A1"].number_format = "[H]:MM:SS"
    sheet.merge_cells("A1:B1")
    book.save(path)
    book.close()
    inputs = inputs_for(path)
    expected = pipeline._run_openpyxl_pipeline(inputs).workbook
    actual = pipeline.run_extraction_pipeline(inputs)
    assert actual.state.fallback_reason is None
    assert actual.workbook.model_dump() == expected.model_dump()
    assert actual.workbook.sheets["Duration"].rows[0].c == {"0": "1 day, 12:00:00"}


@pytest.mark.parametrize("mode", ["standard", "verbose", "libreoffice"])
def test_other_modes_keep_previous_session_selection(
    monkeypatch: pytest.MonkeyPatch,
    mode: pipeline.ExtractionMode,
) -> None:
    inputs = replace(inputs_for(Path("unused.xlsx")), mode=mode)
    sentinel = Mock()
    runner = Mock(return_value=sentinel)
    monkeypatch.setattr(pipeline, "_run_openpyxl_pipeline", runner)
    assert pipeline.run_extraction_pipeline(inputs) is sentinel
    runner.assert_called_once_with(inputs)
