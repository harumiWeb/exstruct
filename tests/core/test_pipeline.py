import logging
from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch
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
from exstruct.models import CellRow, PrintArea, Shape


def test_build_pre_com_pipeline_respects_flags(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
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


def test_resolve_extraction_inputs_forces_merged_cells_when_excluding_values(
    tmp_path: Path,
) -> None:
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
    calls: list[str] = []

    def _warn_once(key: str, message: str) -> None:
        """
        Record a warning key in the shared `calls` list while ignoring the message.

        Parameters:
            key (str): Identifier for the warning; appended to the module-level `calls` list.
            message (str): Ignored placeholder kept for compatibility with expected callback signature.
        """
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


def test_resolve_extraction_inputs_sets_ignore_colors(tmp_path: Path) -> None:
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
    def fake_detect_tables(_: Path, __: str, **_kwargs: object) -> list[str]:
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


def test_filter_rows_excluding_merged_values_updates_links() -> None:
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
    rows = [CellRow(r=1, c={"0": "A"}, links={"0": "L0"})]
    merged_cells = [MergedCellRange(r1=1, c1=0, r2=1, c2=0, v="A")]
    filtered = _filter_rows_excluding_merged_values(rows, merged_cells)
    assert filtered == []


def test_filter_rows_excluding_merged_values_returns_when_empty() -> None:
    assert _filter_rows_excluding_merged_values([], []) == []


def test_filter_rows_excluding_merged_values_keeps_rows_without_intervals() -> None:
    rows = [CellRow(r=1, c={"0": "A"})]
    merged_cells = [MergedCellRange(r1=2, c1=0, r2=2, c2=1, v="B")]
    filtered = _filter_rows_excluding_merged_values(rows, merged_cells)
    assert filtered == rows


def test_filter_rows_excluding_merged_values_drops_links_when_filtered() -> None:
    rows = [CellRow(r=1, c={"0": "A", "1": "B"}, links={"0": "L0"})]
    merged_cells = [MergedCellRange(r1=1, c1=0, r2=1, c2=0, v="A")]
    filtered = _filter_rows_excluding_merged_values(rows, merged_cells)
    assert filtered[0].links is None


def test_resolve_sheet_colors_map_empty() -> None:
    assert _resolve_sheet_colors_map(None, "Sheet1") == {}


def test_resolve_sheet_formulas_map_empty() -> None:
    assert _resolve_sheet_formulas_map(None, "Sheet1") == {}


def test_merge_intervals_merges_adjacent() -> None:
    assert _merge_intervals([(1, 2), (3, 4)]) == [(1, 4)]


def test_col_in_intervals_fast_false() -> None:
    assert _col_in_intervals(1, [(3, 5)]) is False


def test_step_extract_colors_map_openpyxl_sets_data(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    def _fake(
        _backend: OpenpyxlBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        """
        Provide a placeholder colors map for testing that is always empty.

        Parameters:
            include_default_background (bool): Accepted for signature compatibility; has no effect on the returned value.
            ignore_colors (set[str] | None): Accepted for signature compatibility; has no effect on the returned value.

        Returns:
            WorkbookColorsMap: An empty colors map with no sheets.
        """
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
    def _fake_com(
        _backend: ComBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> None:
        """
        No-op placeholder that simulates a COM backend extraction step without producing any side effects.

        This function accepts a COM backend and related flags but intentionally performs no operations; it is used in tests as a stub implementation.
        """
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
        """
        Return an empty WorkbookColorsMap regardless of inputs.

        Parameters:
            include_default_background (bool): Ignored; present for signature compatibility.
            ignore_colors (set[str] | None): Ignored; present for signature compatibility.

        Returns:
            WorkbookColorsMap: A colors map with no sheets.
        """
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
    def _fake(_: ComBackend) -> dict[str, list[PrintArea]]:
        """
        Return a stub mapping of sheet names to print areas containing a single 1x1 print area for "Sheet1".

        Returns:
            dict[str, list[PrintArea]]: Mapping where "Sheet1" maps to a list with one PrintArea covering row 1, column 0 to row 1, column 0.
        """
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
    colors_map = WorkbookColorsMap(sheets={})
    formulas_map = WorkbookFormulasMap(sheets={})

    def _fake_colors(
        _backend: OpenpyxlBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        """
        Return a fake workbook colors map used by tests.

        Parameters:
            _backend (OpenpyxlBackend): Ignored backend parameter retained for signature compatibility.
            include_default_background (bool): Whether the default background color would be included (ignored).
            ignore_colors (set[str] | None): Set of color names to ignore (ignored).

        Returns:
            object: A preconstructed colors map object used by tests.
        """
        _ = _backend
        _ = include_default_background
        _ = ignore_colors
        return colors_map

    def _fake_formulas(_: OpenpyxlBackend) -> object:
        """
        Return the pre-captured formulas_map object.

        Returns:
            The pre-captured `formulas_map` object.
        """
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
    def _raise(_: OpenpyxlBackend) -> object:
        """
        Always raises a RuntimeError with the message "boom".

        Raises:
            RuntimeError: always raised with message "boom".
        """
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
    def _raise(_: ComBackend) -> object:
        """
        Always raises a RuntimeError with message "boom".

        Raises:
            RuntimeError: Always raised by this helper.
        """
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
    rows = [CellRow(r=1, c={"0": "A"})]
    merged_cells = [MergedCellRange(r1=3, c1=0, r2=4, c2=1, v="A")]
    assert _filter_rows_excluding_merged_values(rows, merged_cells) == rows


def test_resolve_sheet_colors_map_missing_sheet() -> None:
    colors_map = WorkbookColorsMap(
        sheets={"Other": SheetColorsMap(sheet_name="Other", colors_map={})}
    )
    assert _resolve_sheet_colors_map(colors_map, "Sheet1") == {}


def test_resolve_sheet_formulas_map_missing_sheet() -> None:
    formulas_map = WorkbookFormulasMap(
        sheets={"Other": SheetFormulasMap(sheet_name="Other", formulas_map={})}
    )
    assert _resolve_sheet_formulas_map(formulas_map, "Sheet1") == {}


def test_merge_intervals_empty() -> None:
    assert _merge_intervals([]) == []


def test_merge_intervals_keeps_non_overlapping() -> None:
    assert _merge_intervals([(1, 2), (5, 6)]) == [(1, 2), (5, 6)]


def test_step_extract_shapes_com_sets_data(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    shapes_data = {"Sheet1": [object()]}

    def _fake(_: object, *, mode: str) -> dict[str, list[object]]:
        """
        Provide a stub that supplies the module-level `shapes_data` mapping.

        Parameters:
            _ (object): Placeholder positional argument; ignored.
            mode (str): Mode selector; ignored.

        Returns:
            dict[str, list[object]]: Mapping of sheet names to lists of shape objects from `shapes_data`.
        """
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
    charts = [object()]

    def _fake(_: object, *, mode: str) -> list[object]:
        """
        Return the captured charts list.

        Parameters:
            mode (str): Ignored; accepted for compatibility with callers.

        Returns:
            list[object]: The charts list captured from the enclosing scope.
        """
        _ = mode
        return charts

    class _Sheet:
        def __init__(self, name: str) -> None:
            """
            Initialize the instance with a display name.

            Parameters:
                name (str): The name to assign to the instance.
            """
            self.name = name

    class _Workbook:
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
    def _raise(_: ComBackend) -> object:
        """
        Raise a RuntimeError indicating this code path must not be invoked.

        This function always raises RuntimeError("should not be called").
        """
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
    def _fake(_: ComBackend) -> dict[str, list[PrintArea]]:
        """
        Return a stub mapping of sheet names to print areas containing a single 1x1 print area for "Sheet1".

        Returns:
            dict[str, list[PrintArea]]: Mapping where "Sheet1" maps to a list with one PrintArea covering row 1, column 0 to row 1, column 0.
        """
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
    colors_map = WorkbookColorsMap(sheets={})

    def _fake_com(
        _backend: ComBackend,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        """
        Return a colors map object suitable for use as a COM backend response.

        Parameters:
            include_default_background (bool): If true, the returned colors map should include the default background color.
            ignore_colors (set[str] | None): Optional set of color identifiers to exclude from the returned map; `None` means no colors are excluded.

        Returns:
            object: A colors map representing workbook-level color mappings.
        """
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
        """
        Placeholder backend sentinel that always raises a RuntimeError when invoked.

        Raises:
            RuntimeError: Always raised with message "should not be called".
        """
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
    calls: list[str] = []

    def _step(_: ExtractionInputs, artifacts: ExtractionArtifacts, __: object) -> None:
        """
        Test pipeline step that simulates shape extraction.

        Sets artifacts.shape_data to a mapping for "Sheet1" containing a single Shape and records invocation by appending "called" to the outer `calls` list.

        Parameters:
            _ (ExtractionInputs): Unused extraction inputs placeholder.
            artifacts (ExtractionArtifacts): Artifacts object to populate with shape data.
            __ (object): Unused context placeholder.
        """
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
    class _Sheet:
        def __init__(self, name: str) -> None:
            """
            Initialize the instance with a display name.

            Parameters:
                name (str): The name to assign to the instance.
            """
            self.name = name

    class _Sheets:
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
            def __enter__(self) -> _Workbook:
                return _Workbook()

            def __exit__(
                self,
                exc_type: type[BaseException] | None,
                exc: BaseException | None,
                tb: object | None,
            ) -> bool | None:
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
