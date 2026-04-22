"""Tests for extraction pipeline planning and step orchestration."""

import logging
from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch
from openpyxl import Workbook
import pytest

from exstruct.core.backends.com_backend import ComBackend
from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend
from exstruct.core.cells import (
    MergedCellRange,
    SheetColorsMap,
    SheetFormulasMap,
    WorkbookColorsMap,
    WorkbookFormulasMap,
)
from exstruct.core.libreoffice import LibreOfficeUnavailableError
from exstruct.core.ooxml_drawing import (
    DrawingShapeRef,
    OoxmlChartInfo,
    OoxmlShapeInfo,
    SheetDrawingData,
)
from exstruct.core.pipeline import (
    ExtractionArtifacts,
    ExtractionInputs,
    PipelinePlan,
    _col_in_intervals,
    _filter_rows_excluding_merged_values,
    _merge_intervals,
    _resolve_sheet_colors_map,
    _resolve_sheet_formulas_map,
    build_cells_tables_workbook,
    build_com_pipeline,
    build_pre_com_pipeline,
    resolve_extraction_inputs,
    run_com_pipeline,
    run_extraction_pipeline,
    step_extract_auto_page_breaks_com,
    step_extract_charts_com,
    step_extract_colors_map_com,
    step_extract_colors_map_openpyxl,
    step_extract_formulas_map_com,
    step_extract_formulas_map_openpyxl,
    step_extract_print_areas_com,
    step_extract_shapes_com,
)
from exstruct.errors import FallbackReason
from exstruct.models import CellRow, Chart, PrintArea, Shape


def test_build_pre_com_pipeline_respects_flags(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that the pre-COM pipeline only includes the requested steps."""

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    steps = build_pre_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert step_names == ["step_extract_cells"]


def test_build_pre_com_pipeline_includes_colors_map_for_light(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that light mode keeps the colors-map step in the pre-COM pipeline."""

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="light",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=True,
        include_merged_values_in_rows=True,
    )
    steps = build_pre_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert step_names == [
        "step_extract_cells",
        "step_extract_print_areas_openpyxl",
        "step_extract_colors_map_openpyxl",
        "step_extract_merged_cells_openpyxl",
    ]


def test_build_pre_com_pipeline_skips_merged_cells_when_disabled(
    tmp_path: Path,
) -> None:
    """Verify that merged-cell extraction is omitted when the flag is disabled."""

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    steps = build_pre_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert "step_extract_merged_cells_openpyxl" not in step_names


def test_build_com_pipeline_respects_flags(tmp_path: Path) -> None:
    """Verify that the COM pipeline includes only the enabled COM steps."""

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=True,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    steps = build_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert step_names == [
        "step_extract_shapes_com",
        "step_extract_charts_com",
        "step_extract_auto_page_breaks_com",
    ]


def test_build_com_pipeline_excludes_auto_page_breaks_when_disabled(
    tmp_path: Path,
) -> None:
    """Verify that auto page-break extraction is skipped when disabled."""

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    steps = build_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert "step_extract_auto_page_breaks_com" not in step_names


def test_build_com_pipeline_empty_for_light(tmp_path: Path) -> None:
    """Verify that light mode does not schedule any COM-only steps."""

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="light",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=True,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    steps = build_com_pipeline(inputs)
    assert steps == []


def test_resolve_extraction_inputs_defaults(tmp_path: Path) -> None:
    """Verify that standard-mode defaults are populated consistently."""

    inputs = resolve_extraction_inputs(
        tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=None,
        include_print_areas=None,
        include_auto_page_breaks=False,
        include_colors_map=None,
        include_default_background=True,
        ignore_colors=None,
        include_formulas_map=None,
        include_merged_cells=None,
        include_merged_values_in_rows=True,
    )
    assert inputs.include_cell_links is False
    assert inputs.include_print_areas is True
    assert inputs.include_colors_map is False
    assert inputs.include_formulas_map is False
    assert inputs.include_default_background is False
    assert inputs.include_merged_cells is True


def test_resolve_extraction_inputs_defaults_for_libreoffice(tmp_path: Path) -> None:
    """Verify that LibreOffice mode uses the same default data-selection flags."""

    inputs = resolve_extraction_inputs(
        tmp_path / "book.xlsx",
        mode="libreoffice",
        include_cell_links=None,
        include_print_areas=None,
        include_auto_page_breaks=False,
        include_colors_map=None,
        include_default_background=True,
        ignore_colors=None,
        include_formulas_map=None,
        include_merged_cells=None,
        include_merged_values_in_rows=True,
    )
    assert inputs.include_cell_links is False
    assert inputs.include_print_areas is True
    assert inputs.include_colors_map is False
    assert inputs.include_formulas_map is False
    assert inputs.include_default_background is False
    assert inputs.include_merged_cells is True


def test_resolve_extraction_inputs_forces_merged_cells_when_excluding_values(
    tmp_path: Path,
) -> None:
    """Verify that merged-cell metadata stays enabled when merged values are excluded."""

    inputs = resolve_extraction_inputs(
        tmp_path / "book.xlsx",
        mode="light",
        include_cell_links=None,
        include_print_areas=None,
        include_auto_page_breaks=False,
        include_colors_map=None,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=None,
        include_merged_cells=False,
        include_merged_values_in_rows=False,
    )
    assert inputs.include_merged_cells is True


def test_resolve_extraction_inputs_warns_on_xls_formulas(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that `.xls` formula extraction emits the compatibility warning."""

    calls: list[str] = []

    def _warn_once(key: str, message: str) -> None:
        """Record the warning key while ignoring the rendered message."""
        calls.append(key)
        _ = message

    monkeypatch.setattr("exstruct.core.pipeline.warn_once", _warn_once)

    inputs = resolve_extraction_inputs(
        tmp_path / "book.xls",
        mode="standard",
        include_cell_links=None,
        include_print_areas=None,
        include_auto_page_breaks=False,
        include_colors_map=None,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=True,
        include_merged_cells=None,
        include_merged_values_in_rows=True,
    )
    assert inputs.use_com_for_formulas is True
    assert calls


def test_resolve_extraction_inputs_rejects_xls_for_libreoffice(tmp_path: Path) -> None:
    """Verify that LibreOffice mode rejects legacy `.xls` workbooks."""

    with pytest.raises(ValueError, match="not supported in libreoffice mode"):
        resolve_extraction_inputs(
            tmp_path / "book.xls",
            mode="libreoffice",
            include_cell_links=None,
            include_print_areas=None,
            include_auto_page_breaks=False,
            include_colors_map=None,
            include_default_background=False,
            ignore_colors=None,
            include_formulas_map=None,
            include_merged_cells=None,
            include_merged_values_in_rows=True,
        )


def test_resolve_extraction_inputs_sets_ignore_colors(tmp_path: Path) -> None:
    """Verify that verbose mode normalizes a missing ignore-colors set."""

    inputs = resolve_extraction_inputs(
        tmp_path / "book.xlsx",
        mode="verbose",
        include_cell_links=None,
        include_print_areas=None,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=None,
        include_merged_cells=None,
        include_merged_values_in_rows=True,
    )
    assert inputs.ignore_colors == set()


def test_build_cells_tables_workbook_uses_print_areas(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that table detection receives the worksheet print areas."""

    def fake_detect_tables(_: Path, __: str, **_kwargs: object) -> list[str]:
        """Return a fixed table range for the test workbook."""

        return ["A1:B2"]

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.detect_tables_openpyxl",
        fake_detect_tables,
    )

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=True,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts(
        cell_data={"Sheet1": [CellRow(r=1, c={"0": "v"})]},
        print_area_data={"Sheet1": [PrintArea(r1=1, c1=0, r2=1, c2=0)]},
        merged_cell_data={"Sheet1": []},
    )
    wb = build_cells_tables_workbook(
        inputs=inputs,
        artifacts=artifacts,
        reason="test",
    )
    sheet = wb.sheets["Sheet1"]
    assert sheet.print_areas
    assert sheet.table_candidates == ["A1:B2"]
    assert sheet.merged_cells is None


def test_build_cells_tables_workbook_excludes_merged_values_in_rows(
    tmp_path: Path,
) -> None:
    """Verify that merged values are removed from row payloads when requested."""

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=True,
        include_merged_values_in_rows=False,
    )
    artifacts = ExtractionArtifacts(
        cell_data={"Sheet1": [CellRow(r=1, c={"0": "A", "1": "B", "2": "C"})]},
        merged_cell_data={"Sheet1": [MergedCellRange(r1=1, c1=0, r2=1, c2=1, v="A")]},
    )
    wb = build_cells_tables_workbook(inputs=inputs, artifacts=artifacts, reason="test")
    sheet = wb.sheets["Sheet1"]
    assert sheet.rows[0].c == {"2": "C"}


def test_build_cells_tables_workbook_keeps_rich_artifacts_in_light(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that fallback workbook assembly preserves rich artifacts in light mode."""

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.detect_tables_openpyxl",
        lambda *_args, **_kwargs: [],
    )

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="light",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts(
        cell_data={"Sheet1": [CellRow(r=1, c={"0": "A"})]},
        shape_data={"Sheet1": [Shape(id=1, text="shape", l=0, t=0)]},
        chart_data={
            "Sheet1": [
                Chart(
                    name="Chart 1",
                    chart_type="Line",
                    title=None,
                    y_axis_title="",
                    y_axis_range=[],
                    series=[],
                    l=0,
                    t=0,
                )
            ]
        },
        merged_cell_data={"Sheet1": []},
    )

    wb = build_cells_tables_workbook(
        inputs=inputs,
        artifacts=artifacts,
        reason="test",
        include_rich_artifacts=True,
    )

    assert len(wb.sheets["Sheet1"].shapes) == 1
    assert len(wb.sheets["Sheet1"].charts) == 1


def test_filter_rows_excluding_merged_values_updates_links() -> None:
    """Verify that row filtering keeps links aligned with the remaining cells."""

    rows = [
        CellRow(
            r=1,
            c={"0": "A", "1": "B", "x": "keep"},
            links={"0": "L0", "1": "L1", "x": "LX"},
        )
    ]
    merged_cells = [MergedCellRange(r1=1, c1=0, r2=1, c2=1, v="A")]
    filtered = _filter_rows_excluding_merged_values(rows, merged_cells)
    assert filtered[0].c == {"x": "keep"}
    assert filtered[0].links == {"x": "LX"}


def test_filter_rows_excluding_merged_values_drops_empty_rows() -> None:
    """Verify that row filtering drops rows that become empty."""

    rows = [CellRow(r=1, c={"0": "A"}, links={"0": "L0"})]
    merged_cells = [MergedCellRange(r1=1, c1=0, r2=1, c2=0, v="A")]
    filtered = _filter_rows_excluding_merged_values(rows, merged_cells)
    assert filtered == []


def test_filter_rows_excluding_merged_values_returns_when_empty() -> None:
    """Verify that row filtering returns an empty list unchanged."""

    assert _filter_rows_excluding_merged_values([], []) == []


def test_filter_rows_excluding_merged_values_keeps_rows_without_intervals() -> None:
    """Verify that row filtering leaves unrelated rows untouched."""

    rows = [CellRow(r=1, c={"0": "A"})]
    merged_cells = [MergedCellRange(r1=2, c1=0, r2=2, c2=1, v="B")]
    filtered = _filter_rows_excluding_merged_values(rows, merged_cells)
    assert filtered == rows


def test_filter_rows_excluding_merged_values_drops_links_when_filtered() -> None:
    """Verify that orphaned links are dropped after row filtering."""

    rows = [CellRow(r=1, c={"0": "A", "1": "B"}, links={"0": "L0"})]
    merged_cells = [MergedCellRange(r1=1, c1=0, r2=1, c2=0, v="A")]
    filtered = _filter_rows_excluding_merged_values(rows, merged_cells)
    assert filtered[0].links is None


def test_resolve_sheet_colors_map_empty() -> None:
    """Verify that a missing workbook colors map resolves to an empty sheet map."""

    assert _resolve_sheet_colors_map(None, "Sheet1") == {}


def test_resolve_sheet_formulas_map_empty() -> None:
    """Verify that a missing workbook formulas map resolves to an empty sheet map."""

    assert _resolve_sheet_formulas_map(None, "Sheet1") == {}


def test_merge_intervals_merges_adjacent() -> None:
    """Verify that adjacent intervals are coalesced into one range."""

    assert _merge_intervals([(1, 2), (3, 4)]) == [(1, 4)]


def test_col_in_intervals_fast_false() -> None:
    """Verify that `_col_in_intervals` short-circuits for columns before the range."""

    assert _col_in_intervals(1, [(3, 5)]) is False


def test_step_extract_colors_map_openpyxl_sets_data(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the openpyxl colors-map step stores extracted data."""

    def _fake(
        _backend: OpenpyxlBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        """Return an empty colors map for the test."""
        _ = _backend
        _ = include_default_background
        _ = ignore_colors
        return WorkbookColorsMap(sheets={})

    monkeypatch.setattr(OpenpyxlBackend, "extract_colors_map", _fake)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()
    step_extract_colors_map_openpyxl(inputs, artifacts)
    assert artifacts.colors_map_data is not None


def test_step_extract_colors_map_com_falls_back(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the COM colors-map step falls back to openpyxl on `None`."""

    def _fake_com(
        _backend: ComBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> None:
        """Simulate a COM extractor that returns no data."""
        _ = _backend
        _ = include_default_background
        _ = ignore_colors
        return None

    def _fake_openpyxl(
        _backend: OpenpyxlBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        """Return the fallback colors map for the test."""
        _ = _backend
        _ = include_default_background
        _ = ignore_colors
        return WorkbookColorsMap(sheets={})

    monkeypatch.setattr(ComBackend, "extract_colors_map", _fake_com)
    monkeypatch.setattr(OpenpyxlBackend, "extract_colors_map", _fake_openpyxl)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()
    step_extract_colors_map_com(inputs, artifacts, object())
    assert artifacts.colors_map_data is not None


def test_step_extract_auto_page_breaks_com_sets_data(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the COM auto-page-break step stores extracted ranges."""

    def _fake(_: ComBackend) -> dict[str, list[PrintArea]]:
        """Return a single-sheet auto page-break payload for the test."""
        return {"Sheet1": [PrintArea(r1=1, c1=0, r2=1, c2=0)]}

    monkeypatch.setattr(ComBackend, "extract_auto_page_breaks", _fake)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=True,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()
    step_extract_auto_page_breaks_com(inputs, artifacts, object())
    assert artifacts.auto_page_break_data


def test_build_cells_tables_workbook_fetches_missing_maps(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that workbook assembly fetches missing colors and formulas maps."""

    colors_map = WorkbookColorsMap(sheets={})
    formulas_map = WorkbookFormulasMap(sheets={})

    def _fake_colors(
        _backend: OpenpyxlBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        """Return the prebuilt colors map for the test."""
        _ = _backend
        _ = include_default_background
        _ = ignore_colors
        return colors_map

    def _fake_formulas(_: OpenpyxlBackend) -> object:
        """Return the prebuilt formulas map for the test."""
        return formulas_map

    monkeypatch.setattr(OpenpyxlBackend, "extract_colors_map", _fake_colors)
    monkeypatch.setattr(OpenpyxlBackend, "extract_formulas_map", _fake_formulas)

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=True,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts(
        cell_data={"Sheet1": [CellRow(r=1, c={"0": "A"})]},
        merged_cell_data={"Sheet1": []},
    )
    wb = build_cells_tables_workbook(inputs=inputs, artifacts=artifacts, reason="test")
    assert "Sheet1" in wb.sheets


def test_step_extract_formulas_map_openpyxl_skips_on_failure(
    tmp_path: Path, monkeypatch: MonkeyPatch, caplog: "pytest.LogCaptureFixture"
) -> None:
    """Verify that openpyxl formulas extraction logs and skips failures."""

    def _raise(_: OpenpyxlBackend) -> object:
        """Raise to simulate an openpyxl formulas extraction failure."""
        raise RuntimeError("boom")

    monkeypatch.setattr(OpenpyxlBackend, "extract_formulas_map", _raise)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=True,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()

    with caplog.at_level(logging.WARNING):
        step_extract_formulas_map_openpyxl(inputs, artifacts)

    assert artifacts.formulas_map_data is None
    assert "Failed to extract formulas_map via openpyxl" in caplog.text


def test_step_extract_formulas_map_com_skips_on_failure(
    tmp_path: Path, monkeypatch: MonkeyPatch, caplog: "pytest.LogCaptureFixture"
) -> None:
    """Verify that COM formulas extraction logs and skips failures."""

    def _raise(_: ComBackend) -> object:
        """Raise to simulate a COM formulas extraction failure."""
        raise RuntimeError("boom")

    monkeypatch.setattr(ComBackend, "extract_formulas_map", _raise)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=True,
        use_com_for_formulas=True,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()

    with caplog.at_level(logging.WARNING):
        step_extract_formulas_map_com(inputs, artifacts, object())

    assert artifacts.formulas_map_data is None
    assert "Failed to extract formulas_map via COM" in caplog.text


def test_filter_rows_excluding_merged_values_returns_rows_when_intervals_empty() -> (
    None
):
    """Verify that row filtering is a no-op when there are no merged intervals."""

    rows = [CellRow(r=1, c={"0": "A"})]
    merged_cells = [MergedCellRange(r1=3, c1=0, r2=4, c2=1, v="A")]
    assert _filter_rows_excluding_merged_values(rows, merged_cells) == rows


def test_resolve_sheet_colors_map_missing_sheet() -> None:
    """Verify that missing sheets resolve to an empty colors map."""

    colors_map = WorkbookColorsMap(
        sheets={"Other": SheetColorsMap(sheet_name="Other", colors_map={})}
    )
    assert _resolve_sheet_colors_map(colors_map, "Sheet1") == {}


def test_resolve_sheet_formulas_map_missing_sheet() -> None:
    """Verify that missing sheets resolve to an empty formulas map."""

    formulas_map = WorkbookFormulasMap(
        sheets={"Other": SheetFormulasMap(sheet_name="Other", formulas_map={})}
    )
    assert _resolve_sheet_formulas_map(formulas_map, "Sheet1") == {}


def test_merge_intervals_empty() -> None:
    """Verify that merging an empty interval list returns an empty list."""

    assert _merge_intervals([]) == []


def test_merge_intervals_keeps_non_overlapping() -> None:
    """Verify that non-overlapping intervals are preserved as-is."""

    assert _merge_intervals([(1, 2), (5, 6)]) == [(1, 2), (5, 6)]


def test_step_extract_shapes_com_sets_data(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the COM shapes step stores extracted shape data."""

    shapes_data = {"Sheet1": [object()]}

    def _fake(_: object, *, mode: str) -> dict[str, list[object]]:
        """Return the shape payload captured in the enclosing test."""
        _ = mode
        return shapes_data

    monkeypatch.setattr("exstruct.core.pipeline.get_shapes_with_position", _fake)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()
    step_extract_shapes_com(inputs, artifacts, object())
    assert artifacts.shape_data == shapes_data


def test_step_extract_charts_com_sets_data(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the COM charts step stores chart data by sheet."""

    charts = [object()]

    def _fake(_: object, *, mode: str) -> list[object]:
        """Return the chart payload captured in the enclosing test."""
        _ = mode
        return charts

    class _Sheet:
        """Minimal worksheet test double."""

        def __init__(self, name: str) -> None:
            """Store the worksheet display name used by the pipeline."""
            self.name = name

    class _Workbook:
        """Minimal workbook test double."""

        sheets = [_Sheet("Sheet1")]

    monkeypatch.setattr("exstruct.core.pipeline.get_charts", _fake)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()
    step_extract_charts_com(inputs, artifacts, _Workbook())
    assert artifacts.chart_data == {"Sheet1": charts}


def test_step_extract_print_areas_com_skips_when_present(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the COM print-area step does not overwrite existing data."""

    def _raise(_: ComBackend) -> object:
        """Raise if the existing-data fast path stops working."""
        raise RuntimeError("should not be called")

    monkeypatch.setattr(ComBackend, "extract_print_areas", _raise)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts(
        print_area_data={"Sheet1": [PrintArea(r1=1, c1=0, r2=1, c2=0)]}
    )
    step_extract_print_areas_com(inputs, artifacts, object())


def test_step_extract_print_areas_com_sets_data(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the COM print-area step stores the extracted ranges."""

    def _fake(_: ComBackend) -> dict[str, list[PrintArea]]:
        """Return a single-sheet print-area payload for the test."""
        return {"Sheet1": [PrintArea(r1=1, c1=0, r2=1, c2=0)]}

    monkeypatch.setattr(ComBackend, "extract_print_areas", _fake)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()
    step_extract_print_areas_com(inputs, artifacts, object())
    assert artifacts.print_area_data


def test_step_extract_colors_map_com_sets_data(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the COM colors-map step stores its direct result."""

    colors_map = WorkbookColorsMap(sheets={})

    def _fake_com(
        _backend: ComBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        """Return the prebuilt COM colors map for the test."""
        _ = _backend
        _ = include_default_background
        _ = ignore_colors
        return colors_map

    def _raise(
        _backend: OpenpyxlBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        """Raise if the fallback path runs unexpectedly."""
        _ = _backend
        _ = include_default_background
        _ = ignore_colors
        raise RuntimeError("should not be called")

    monkeypatch.setattr(ComBackend, "extract_colors_map", _fake_com)
    monkeypatch.setattr(OpenpyxlBackend, "extract_colors_map", _raise)
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()
    step_extract_colors_map_com(inputs, artifacts, object())
    assert artifacts.colors_map_data is colors_map


def test_run_com_pipeline_executes_steps(tmp_path: Path) -> None:
    """Verify that `run_com_pipeline` executes each planned step once."""

    calls: list[str] = []

    def _step(_: ExtractionInputs, artifacts: ExtractionArtifacts, __: object) -> None:
        """Record the step invocation and populate one shape payload."""
        calls.append("called")
        artifacts.shape_data = {"Sheet1": [Shape(id=1, text="", l=0, t=0)]}

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )
    artifacts = ExtractionArtifacts()
    run_com_pipeline([_step], inputs, artifacts, object())
    assert calls == ["called"]
    assert artifacts.shape_data


def test_run_extraction_pipeline_com_success(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that run extraction pipeline COM success."""

    class _Sheet:
        """Minimal worksheet test double."""

        def __init__(self, name: str) -> None:
            """
            Initialize the instance with a display name.

            Parameters:
                name (str): The name to assign to the instance.
            """
            self.name = name

    class _Sheets:
        """Minimal worksheet collection test double."""

        def __init__(self) -> None:
            """
            Initialize the object with a single default sheet named "Sheet1".

            Creates the internal mapping `self._sheets` and populates it with one `_Sheet` instance keyed by "Sheet1".
            """
            self._sheets = {"Sheet1": _Sheet("Sheet1")}

        def __getitem__(self, name: str) -> _Sheet:
            """
            Access a worksheet by its name.

            Parameters:
                name (str): The name of the sheet to retrieve.

            Returns:
                _Sheet: The sheet object associated with `name`.

            Raises:
                KeyError: If no sheet with the given name exists.
            """
            return self._sheets[name]

    class _Workbook:
        """Minimal workbook test double."""

        sheets = _Sheets()

    def _pre_step(_: ExtractionInputs, artifacts: ExtractionArtifacts) -> None:
        """
        Populate artifacts with default minimal cell and merged-cell data for a single sheet.

        Parameters:
            _ (ExtractionInputs): Unused extraction inputs placeholder.
            artifacts (ExtractionArtifacts): Mutable extraction artifacts that will be updated with
                `cell_data` set to a single row for "Sheet1" and `merged_cell_data` set to an empty list
                for "Sheet1".
        """
        artifacts.cell_data = {"Sheet1": [CellRow(r=1, c={"0": "A"})]}
        artifacts.merged_cell_data = {"Sheet1": []}

    def _fake_plan(_: ExtractionInputs) -> PipelinePlan:
        """
        Create a fixed PipelinePlan for tests that forces COM usage and provides a single pre-COM step.

        Parameters:
            _ (ExtractionInputs): Ignored input; present to match the PipelinePlan factory signature.

        Returns:
            PipelinePlan: A plan with `pre_com_steps` set to a list containing `_pre_step`, `com_steps` empty, and `use_com` set to `True`.
        """
        return PipelinePlan(pre_com_steps=[_pre_step], com_steps=[], use_com=True)

    def _fake_detect_tables(_: object, **_kwargs: object) -> list[str]:
        """
        Provide a detector that always reports no table ranges.

        The input workbook-like object is ignored.

        Returns:
            list[str]: An empty list of table range identifiers.
        """
        return []

    def _fake_workbook(_: Path) -> object:
        """
        Provide a context manager that yields a lightweight fake workbook for tests.

        Parameters:
            _ (Path): Ignored file path parameter retained to match the real backend signature.

        Returns:
            object: A context manager whose `__enter__` returns a new `_Workbook` instance and whose `__exit__` does not suppress exceptions (returns `None`).
        """

        class _Context:
            """Context-manager test double for workbook access."""

            def __enter__(self) -> _Workbook:
                """Return the test double as the context-manager result."""

                return _Workbook()

            def __exit__(
                self,
                exc_type: type[BaseException] | None,
                exc: BaseException | None,
                tb: object | None,
            ) -> bool | None:
                """Accept context-manager exit arguments without suppressing errors."""

                _ = exc_type
                _ = exc
                _ = tb
                return None

        return _Context()

    monkeypatch.delenv("SKIP_COM_TESTS", raising=False)
    monkeypatch.setattr("exstruct.core.pipeline.build_pipeline_plan", _fake_plan)
    monkeypatch.setattr("exstruct.core.pipeline.detect_tables", _fake_detect_tables)
    monkeypatch.setattr("exstruct.core.pipeline.xlwings_workbook", _fake_workbook)

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=False,
        include_merged_values_in_rows=True,
    )

    result = run_extraction_pipeline(inputs)
    assert result.state.com_attempted is True
    assert result.state.com_succeeded is True
    assert "Sheet1" in result.workbook.sheets


def test_run_extraction_pipeline_preserves_ooxml_baseline_when_libreoffice_unavailable(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that LibreOffice fallback keeps the OOXML rich baseline in the workbook."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "value"
    path = tmp_path / "book.xlsx"
    workbook.save(path)
    workbook.close()

    monkeypatch.setattr(
        "exstruct.core.backends.ooxml_backend.read_sheet_drawings",
        lambda _path: {
            "Sheet1": SheetDrawingData(
                shapes=[
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=1,
                            name="Shape 1",
                            kind="shape",
                            left=10,
                            top=20,
                            width=40,
                            height=30,
                        ),
                        text="shape",
                        shape_type="AutoShape-Rectangle",
                    )
                ],
                charts=[
                    OoxmlChartInfo(
                        name="Chart 1",
                        chart_type="Column",
                        title="title",
                        y_axis_title="Y",
                        y_axis_range=[],
                        series=[],
                        anchor_left=50,
                        anchor_top=60,
                        anchor_width=120,
                        anchor_height=90,
                    )
                ],
            )
        },
    )

    def _raise_backend(**_kwargs: object) -> object:
        raise LibreOfficeUnavailableError("missing runtime")

    monkeypatch.setattr("exstruct.core.pipeline.resolve_rich_backend", _raise_backend)

    inputs = ExtractionInputs(
        file_path=path,
        mode="libreoffice",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=True,
        include_merged_values_in_rows=True,
    )

    result = run_extraction_pipeline(inputs)

    assert result.state.fallback_reason == FallbackReason.LIBREOFFICE_UNAVAILABLE
    assert len(result.workbook.sheets["Sheet1"].shapes) == 1
    assert result.workbook.sheets["Sheet1"].shapes[0].provenance == "python_ooxml"
    assert len(result.workbook.sheets["Sheet1"].charts) == 1
    assert result.workbook.sheets["Sheet1"].charts[0].provenance == "python_ooxml"


def test_run_extraction_pipeline_keeps_baseline_charts_when_libreoffice_chart_step_fails(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that chart-step fallback preserves the seeded OOXML chart baseline."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "value"
    path = tmp_path / "book.xlsx"
    workbook.save(path)
    workbook.close()

    monkeypatch.setattr(
        "exstruct.core.backends.ooxml_backend.read_sheet_drawings",
        lambda _path: {
            "Sheet1": SheetDrawingData(
                charts=[
                    OoxmlChartInfo(
                        name="Chart 1",
                        chart_type="Column",
                        title="title",
                        y_axis_title="Y",
                        y_axis_range=[],
                        series=[],
                        anchor_left=50,
                        anchor_top=60,
                        anchor_width=120,
                        anchor_height=90,
                    )
                ]
            )
        },
    )

    class _Backend:
        def extract_shapes(self, *, mode: str) -> dict[str, list[Shape]]:
            _ = mode
            return {
                "Sheet1": [
                    Shape(id=1, text="uno", l=0, t=0, provenance="libreoffice_uno")
                ]
            }

        def extract_charts(self, *, mode: str) -> dict[str, list[Chart]]:
            _ = mode
            raise LibreOfficeUnavailableError("chart geometry failed")

    monkeypatch.setattr(
        "exstruct.core.pipeline.resolve_rich_backend",
        lambda **_kwargs: _Backend(),
    )

    inputs = ExtractionInputs(
        file_path=path,
        mode="libreoffice",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=True,
        include_merged_values_in_rows=True,
    )

    result = run_extraction_pipeline(inputs)

    assert result.state.fallback_reason == FallbackReason.LIBREOFFICE_UNAVAILABLE
    assert len(result.workbook.sheets["Sheet1"].shapes) == 1
    assert result.workbook.sheets["Sheet1"].shapes[0].provenance == "libreoffice_uno"
    assert len(result.workbook.sheets["Sheet1"].charts) == 1
    assert result.workbook.sheets["Sheet1"].charts[0].provenance == "python_ooxml"


def test_run_extraction_pipeline_continues_when_libreoffice_ooxml_seed_fails(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify LibreOffice enrichment still runs when OOXML baseline seeding fails."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "value"
    path = tmp_path / "book.xlsx"
    workbook.save(path)
    workbook.close()

    class _Backend:
        def extract_shapes(self, *, mode: str) -> dict[str, list[Shape]]:
            _ = mode
            return {
                "Sheet1": [
                    Shape(id=1, text="uno", l=1, t=2, provenance="libreoffice_uno")
                ]
            }

        def extract_charts(self, *, mode: str) -> dict[str, list[Chart]]:
            _ = mode
            return {
                "Sheet1": [
                    Chart(
                        name="Chart 1",
                        chart_type="Column",
                        title="title",
                        y_axis_title="Y",
                        y_axis_range=[],
                        series=[],
                        l=10,
                        t=20,
                        provenance="libreoffice_uno",
                    )
                ]
            }

    def _raise_seed_failure(self: object, *, mode: str) -> dict[str, list[Shape]]:
        raise RuntimeError(f"bad {mode} seed")

    monkeypatch.setattr(
        "exstruct.core.pipeline.OoxmlRichBackend.extract_shapes",
        _raise_seed_failure,
    )
    monkeypatch.setattr(
        "exstruct.core.pipeline.resolve_rich_backend",
        lambda **_kwargs: _Backend(),
    )

    inputs = ExtractionInputs(
        file_path=path,
        mode="libreoffice",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=True,
        include_merged_values_in_rows=True,
    )

    result = run_extraction_pipeline(inputs)

    assert result.state.fallback_reason is None
    assert len(result.workbook.sheets["Sheet1"].shapes) == 1
    assert result.workbook.sheets["Sheet1"].shapes[0].provenance == "libreoffice_uno"
    assert len(result.workbook.sheets["Sheet1"].charts) == 1
    assert result.workbook.sheets["Sheet1"].charts[0].provenance == "libreoffice_uno"


def test_run_extraction_pipeline_falls_back_when_light_chart_step_fails(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify light-mode rich extraction degrades to the fallback workbook on errors."""

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "value"
    path = tmp_path / "book.xlsx"
    workbook.save(path)
    workbook.close()

    class _Backend:
        def extract_shapes(self, *, mode: str) -> dict[str, list[Shape]]:
            _ = mode
            return {
                "Sheet1": [
                    Shape(id=1, text="shape", l=5, t=6, provenance="python_ooxml")
                ]
            }

        def extract_charts(self, *, mode: str) -> dict[str, list[Chart]]:
            _ = mode
            raise RuntimeError("bad chart payload")

    monkeypatch.setattr(
        "exstruct.core.pipeline.resolve_rich_backend",
        lambda **_kwargs: _Backend(),
    )

    inputs = ExtractionInputs(
        file_path=path,
        mode="light",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=False,
        use_com_for_formulas=False,
        include_merged_cells=True,
        include_merged_values_in_rows=True,
    )

    result = run_extraction_pipeline(inputs)

    assert result.state.fallback_reason == FallbackReason.LIGHT_PIPELINE_FAILED
    assert result.workbook.sheets["Sheet1"].rows
    assert len(result.workbook.sheets["Sheet1"].shapes) == 1
    assert result.workbook.sheets["Sheet1"].shapes[0].provenance == "python_ooxml"
    assert result.workbook.sheets["Sheet1"].charts == []
