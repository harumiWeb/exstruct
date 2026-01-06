from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
import logging
import os
from pathlib import Path
from typing import Literal

import xlwings as xw

from ..errors import FallbackReason
from ..models import (
    Arrow,
    CellRow,
    Chart,
    PrintArea,
    Shape,
    SmartArt,
    WorkbookData,
)
from .backends.com_backend import ComBackend
from .backends.openpyxl_backend import OpenpyxlBackend
from .cells import MergedCellRange, WorkbookColorsMap, detect_tables
from .charts import get_charts
from .logging_utils import log_fallback
from .modeling import SheetRawData, WorkbookRawData, build_workbook_data
from .shapes import get_shapes_with_position
from .workbook import xlwings_workbook

ExtractionMode = Literal["light", "standard", "verbose"]
CellData = dict[str, list[CellRow]]
PrintAreaData = dict[str, list[PrintArea]]
MergedCellData = dict[str, list[MergedCellRange]]
ShapeData = dict[str, list[Shape | Arrow | SmartArt]]
ChartData = dict[str, list[Chart]]

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ExtractionInputs:
    """Immutable inputs for pipeline steps.

    Attributes:
        file_path: Path to the Excel workbook.
        mode: Extraction mode (light/standard/verbose).
        include_cell_links: Whether to include cell hyperlinks.
        include_print_areas: Whether to include print areas.
        include_auto_page_breaks: Whether to include auto page breaks.
        include_colors_map: Whether to include background colors map.
        include_default_background: Whether to include default background color.
        ignore_colors: Optional set of color keys to ignore.
        include_merged_cells: Whether to include merged cell ranges.
        include_merged_values_in_rows: Whether to keep merged values in rows.
    """

    file_path: Path
    mode: ExtractionMode
    include_cell_links: bool
    include_print_areas: bool
    include_auto_page_breaks: bool
    include_colors_map: bool
    include_default_background: bool
    ignore_colors: set[str] | None
    include_merged_cells: bool
    include_merged_values_in_rows: bool


@dataclass
class ExtractionArtifacts:
    """Mutable artifacts collected by pipeline steps.

    Attributes:
        cell_data: Extracted cell rows per sheet.
        print_area_data: Extracted print areas per sheet.
        auto_page_break_data: Extracted auto page-break areas per sheet.
        colors_map_data: Extracted colors map for workbook sheets.
        shape_data: Extracted shapes per sheet.
        chart_data: Extracted charts per sheet.
        merged_cell_data: Extracted merged cell ranges per sheet.
    """

    cell_data: CellData = field(default_factory=dict)
    print_area_data: PrintAreaData = field(default_factory=dict)
    auto_page_break_data: PrintAreaData = field(default_factory=dict)
    colors_map_data: WorkbookColorsMap | None = None
    shape_data: ShapeData = field(default_factory=dict)
    chart_data: ChartData = field(default_factory=dict)
    merged_cell_data: MergedCellData = field(default_factory=dict)


ExtractionStep = Callable[[ExtractionInputs, ExtractionArtifacts], None]
ComExtractionStep = Callable[[ExtractionInputs, ExtractionArtifacts, xw.Book], None]


@dataclass(frozen=True)
class PipelinePlan:
    """Resolved pipeline plan for an extraction run.

    Attributes:
        pre_com_steps: Ordered list of steps to run before COM access.
        com_steps: Ordered list of steps to run with COM access.
        use_com: Whether COM-based extraction should be attempted.
    """

    pre_com_steps: list[ExtractionStep]
    com_steps: list[ComExtractionStep]
    use_com: bool


@dataclass(frozen=True)
class StepConfig:
    """Configuration for a pipeline step.

    Attributes:
        name: Step name for debugging.
        step: Callable to execute.
        enabled: Predicate to include the step in the pipeline.
    """

    name: str
    step: ExtractionStep
    enabled: Callable[[ExtractionInputs], bool]


@dataclass(frozen=True)
class ComStepConfig:
    """Configuration for a COM pipeline step.

    Attributes:
        name: Step name for debugging.
        step: Callable to execute with COM workbook.
        enabled: Predicate to include the step in the pipeline.
    """

    name: str
    step: ComExtractionStep
    enabled: Callable[[ExtractionInputs], bool]


@dataclass
class PipelineState:
    """Mutable execution state for a pipeline run.

    Attributes:
        com_attempted: Whether COM access was attempted.
        com_succeeded: Whether COM steps completed successfully.
        fallback_reason: Optional fallback reason code.
    """

    com_attempted: bool = False
    com_succeeded: bool = False
    fallback_reason: FallbackReason | None = None


@dataclass(frozen=True)
class PipelineResult:
    """Result of a pipeline run.

    Attributes:
        workbook: Extracted workbook data.
        artifacts: Collected extraction artifacts.
        state: Pipeline execution state.
    """

    workbook: WorkbookData
    artifacts: ExtractionArtifacts
    state: PipelineState


def resolve_extraction_inputs(
    file_path: str | Path,
    *,
    mode: ExtractionMode,
    include_cell_links: bool | None,
    include_print_areas: bool | None,
    include_auto_page_breaks: bool,
    include_colors_map: bool | None,
    include_default_background: bool,
    ignore_colors: set[str] | None,
    include_merged_cells: bool | None,
    include_merged_values_in_rows: bool,
) -> ExtractionInputs:
    """Resolve include flags and normalize inputs for the pipeline.

    Args:
        file_path: Workbook path (str or Path).
        mode: Extraction mode.
        include_cell_links: Whether to include hyperlinks; None uses mode defaults.
        include_print_areas: Whether to include print areas; None defaults to True.
        include_auto_page_breaks: Whether to include auto page breaks.
        include_colors_map: Whether to include background colors; None uses mode defaults.
        include_default_background: Include default background colors when colors_map is enabled.
        ignore_colors: Optional set of colors to ignore when colors_map is enabled.
        include_merged_cells: Whether to include merged cell ranges; None uses mode defaults.
        include_merged_values_in_rows: Whether to keep merged values in rows.

    Returns:
        Resolved ExtractionInputs.

    Raises:
        ValueError: If an unsupported mode is provided.
    """
    allowed_modes: set[str] = {"light", "standard", "verbose"}
    if mode not in allowed_modes:
        raise ValueError(f"Unsupported mode: {mode}")

    normalized_file_path = file_path if isinstance(file_path, Path) else Path(file_path)
    resolved_cell_links = (
        include_cell_links if include_cell_links is not None else mode == "verbose"
    )
    resolved_print_areas = (
        include_print_areas if include_print_areas is not None else True
    )
    resolved_colors_map = (
        include_colors_map if include_colors_map is not None else mode == "verbose"
    )
    resolved_default_background = (
        include_default_background if resolved_colors_map else False
    )
    resolved_ignore_colors = ignore_colors if resolved_colors_map else None
    if resolved_colors_map and resolved_ignore_colors is None:
        resolved_ignore_colors = set()
    resolved_merged_cells = (
        include_merged_cells if include_merged_cells is not None else mode != "light"
    )
    if not include_merged_values_in_rows:
        resolved_merged_cells = True

    return ExtractionInputs(
        file_path=normalized_file_path,
        mode=mode,
        include_cell_links=resolved_cell_links,
        include_print_areas=resolved_print_areas,
        include_auto_page_breaks=include_auto_page_breaks,
        include_colors_map=resolved_colors_map,
        include_default_background=resolved_default_background,
        ignore_colors=resolved_ignore_colors,
        include_merged_cells=resolved_merged_cells,
        include_merged_values_in_rows=include_merged_values_in_rows,
    )


def build_pipeline_plan(inputs: ExtractionInputs) -> PipelinePlan:
    """Build a pipeline plan based on resolved inputs.

    Args:
        inputs: Resolved pipeline inputs.

    Returns:
        PipelinePlan containing pre-COM/COM steps and COM usage flag.
    """
    return PipelinePlan(
        pre_com_steps=build_pre_com_pipeline(inputs),
        com_steps=build_com_pipeline(inputs),
        use_com=inputs.mode != "light",
    )


def build_pre_com_pipeline(inputs: ExtractionInputs) -> list[ExtractionStep]:
    """Build pipeline steps that run before COM/Excel access.

    Args:
        inputs: Pipeline inputs describing extraction flags.

    Returns:
        Ordered list of extraction steps to run before COM.
    """
    step_table: dict[ExtractionMode, Sequence[StepConfig]] = {
        "light": (
            StepConfig(
                name="cells",
                step=step_extract_cells,
                enabled=lambda _inputs: True,
            ),
            StepConfig(
                name="print_areas_openpyxl",
                step=step_extract_print_areas_openpyxl,
                enabled=lambda _inputs: _inputs.include_print_areas,
            ),
            StepConfig(
                name="colors_map_openpyxl",
                step=step_extract_colors_map_openpyxl,
                enabled=lambda _inputs: _inputs.include_colors_map,
            ),
            StepConfig(
                name="merged_cells_openpyxl",
                step=step_extract_merged_cells_openpyxl,
                enabled=lambda _inputs: _inputs.include_merged_cells,
            ),
        ),
        "standard": (
            StepConfig(
                name="cells",
                step=step_extract_cells,
                enabled=lambda _inputs: True,
            ),
            StepConfig(
                name="print_areas_openpyxl",
                step=step_extract_print_areas_openpyxl,
                enabled=lambda _inputs: _inputs.include_print_areas,
            ),
            StepConfig(
                name="colors_map_openpyxl_if_skip_com",
                step=step_extract_colors_map_openpyxl,
                enabled=lambda _inputs: _inputs.include_colors_map
                and bool(os.getenv("SKIP_COM_TESTS")),
            ),
            StepConfig(
                name="merged_cells_openpyxl",
                step=step_extract_merged_cells_openpyxl,
                enabled=lambda _inputs: _inputs.include_merged_cells,
            ),
        ),
        "verbose": (
            StepConfig(
                name="cells",
                step=step_extract_cells,
                enabled=lambda _inputs: True,
            ),
            StepConfig(
                name="print_areas_openpyxl",
                step=step_extract_print_areas_openpyxl,
                enabled=lambda _inputs: _inputs.include_print_areas,
            ),
            StepConfig(
                name="colors_map_openpyxl_if_skip_com",
                step=step_extract_colors_map_openpyxl,
                enabled=lambda _inputs: _inputs.include_colors_map
                and bool(os.getenv("SKIP_COM_TESTS")),
            ),
            StepConfig(
                name="merged_cells_openpyxl",
                step=step_extract_merged_cells_openpyxl,
                enabled=lambda _inputs: _inputs.include_merged_cells,
            ),
        ),
    }
    steps: list[ExtractionStep] = []
    for config in step_table[inputs.mode]:
        if config.enabled(inputs):
            steps.append(config.step)
    return steps


def build_com_pipeline(inputs: ExtractionInputs) -> list[ComExtractionStep]:
    """Build pipeline steps that require COM/Excel access.

    Args:
        inputs: Pipeline inputs describing extraction flags.

    Returns:
        Ordered list of COM extraction steps.
    """
    if inputs.mode == "light":
        return []
    step_table: Sequence[ComStepConfig] = (
        ComStepConfig(
            name="shapes_com",
            step=step_extract_shapes_com,
            enabled=lambda _inputs: True,
        ),
        ComStepConfig(
            name="charts_com",
            step=step_extract_charts_com,
            enabled=lambda _inputs: True,
        ),
        ComStepConfig(
            name="print_areas_com",
            step=step_extract_print_areas_com,
            enabled=lambda _inputs: _inputs.include_print_areas,
        ),
        ComStepConfig(
            name="auto_page_breaks_com",
            step=step_extract_auto_page_breaks_com,
            enabled=lambda _inputs: _inputs.include_auto_page_breaks,
        ),
        ComStepConfig(
            name="colors_map_com",
            step=step_extract_colors_map_com,
            enabled=lambda _inputs: _inputs.include_colors_map,
        ),
    )
    steps: list[ComExtractionStep] = []
    for config in step_table:
        if config.enabled(inputs):
            steps.append(config.step)
    return steps


def run_pipeline(
    steps: Sequence[ExtractionStep],
    inputs: ExtractionInputs,
    artifacts: ExtractionArtifacts,
) -> ExtractionArtifacts:
    """Run steps in order and return updated artifacts.

    Args:
        steps: Ordered extraction steps.
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.

    Returns:
        Updated artifacts after running all steps.
    """
    for step in steps:
        step(inputs, artifacts)
    return artifacts


def run_com_pipeline(
    steps: Sequence[ComExtractionStep],
    inputs: ExtractionInputs,
    artifacts: ExtractionArtifacts,
    workbook: xw.Book,
) -> ExtractionArtifacts:
    """Run COM steps in order and return updated artifacts.

    Args:
        steps: Ordered COM extraction steps.
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
        workbook: xlwings workbook instance.

    Returns:
        Updated artifacts after running all COM steps.
    """
    for step in steps:
        step(inputs, artifacts, workbook)
    return artifacts


def step_extract_cells(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts
) -> None:
    """Extract cell rows, optionally including hyperlinks.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
    """
    backend = OpenpyxlBackend(inputs.file_path)
    artifacts.cell_data = backend.extract_cells(include_links=inputs.include_cell_links)


def step_extract_print_areas_openpyxl(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts
) -> None:
    """Extract print areas via openpyxl.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
    """
    backend = OpenpyxlBackend(inputs.file_path)
    artifacts.print_area_data = backend.extract_print_areas()


def step_extract_colors_map_openpyxl(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts
) -> None:
    """Extract colors_map via openpyxl; logs and skips on failure.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
    """
    backend = OpenpyxlBackend(inputs.file_path)
    artifacts.colors_map_data = backend.extract_colors_map(
        include_default_background=inputs.include_default_background,
        ignore_colors=inputs.ignore_colors,
    )


def step_extract_merged_cells_openpyxl(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts
) -> None:
    """Extract merged cell ranges via openpyxl.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
    """
    backend = OpenpyxlBackend(inputs.file_path)
    artifacts.merged_cell_data = backend.extract_merged_cells()


def step_extract_shapes_com(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts, workbook: xw.Book
) -> None:
    """Extract shapes via COM.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
        workbook: xlwings workbook instance.
    """
    artifacts.shape_data = get_shapes_with_position(workbook, mode=inputs.mode)


def step_extract_charts_com(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts, workbook: xw.Book
) -> None:
    """Extract charts via COM.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
        workbook: xlwings workbook instance.
    """
    chart_data: ChartData = {}
    for sheet in workbook.sheets:
        chart_data[sheet.name] = get_charts(sheet, mode=inputs.mode)
    artifacts.chart_data = chart_data


def step_extract_print_areas_com(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts, workbook: xw.Book
) -> None:
    """Extract print areas via COM when openpyxl data is unavailable.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
        workbook: xlwings workbook instance.
    """
    if artifacts.print_area_data:
        return
    artifacts.print_area_data = ComBackend(workbook).extract_print_areas()


def step_extract_auto_page_breaks_com(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts, workbook: xw.Book
) -> None:
    """Extract auto page breaks via COM.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
        workbook: xlwings workbook instance.
    """
    artifacts.auto_page_break_data = ComBackend(workbook).extract_auto_page_breaks()


def step_extract_colors_map_com(
    inputs: ExtractionInputs, artifacts: ExtractionArtifacts, workbook: xw.Book
) -> None:
    """Extract colors_map via COM, falling back to openpyxl when needed.

    Args:
        inputs: Pipeline inputs.
        artifacts: Artifact container to update.
        workbook: xlwings workbook instance.
    """
    com_result = ComBackend(workbook).extract_colors_map(
        include_default_background=inputs.include_default_background,
        ignore_colors=inputs.ignore_colors,
    )
    if com_result is not None:
        artifacts.colors_map_data = com_result
        return
    if artifacts.colors_map_data is None:
        artifacts.colors_map_data = OpenpyxlBackend(
            inputs.file_path
        ).extract_colors_map(
            include_default_background=inputs.include_default_background,
            ignore_colors=inputs.ignore_colors,
        )


def _resolve_sheet_colors_map(
    colors_map_data: WorkbookColorsMap | None, sheet_name: str
) -> dict[str, list[tuple[int, int]]]:
    """Resolve colors_map for a single sheet.

    Args:
        colors_map_data: Optional workbook colors map container.
        sheet_name: Target sheet name.

    Returns:
        colors_map dictionary for the sheet, or empty dict if unavailable.
    """
    if not colors_map_data:
        return {}
    sheet_colors = colors_map_data.get_sheet(sheet_name)
    if sheet_colors is None:
        return {}
    return sheet_colors.colors_map


def _filter_rows_excluding_merged_values(
    rows: list[CellRow],
    merged_cells: list[MergedCellRange],
) -> list[CellRow]:
    """Remove merged-cell values from rows.

    Args:
        rows: Extracted rows.
        merged_cells: Merged cell ranges.

    Returns:
        Filtered rows with merged-cell values removed.
    """
    if not rows or not merged_cells:
        return rows
    intervals_by_row = _build_merged_row_intervals(merged_cells)
    if not intervals_by_row:
        return rows
    filtered_rows: list[CellRow] = []
    for row in rows:
        intervals = intervals_by_row.get(row.r)
        if not intervals:
            filtered_rows.append(row)
            continue
        filtered_cells: dict[str, int | float | str] = {}
        for col_key, value in row.c.items():
            col_index = _safe_col_index(col_key)
            if col_index is None:
                filtered_cells[col_key] = value
                continue
            if not _col_in_intervals(col_index, intervals):
                filtered_cells[col_key] = value
        if not filtered_cells:
            continue
        filtered_links = None
        if row.links:
            filtered_links = {
                col_key: link
                for col_key, link in row.links.items()
                if col_key in filtered_cells
            }
            if not filtered_links:
                filtered_links = None
        filtered_rows.append(CellRow(r=row.r, c=filtered_cells, links=filtered_links))
    return filtered_rows


def _build_merged_row_intervals(
    merged_cells: list[MergedCellRange],
) -> dict[int, list[tuple[int, int]]]:
    """Build row -> merged column intervals lookup.

    Args:
        merged_cells: Merged cell ranges.

    Returns:
        Mapping of row index to merged column intervals.
    """
    intervals_by_row: dict[int, list[tuple[int, int]]] = {}
    for cell in merged_cells:
        for row in range(cell.r1, cell.r2 + 1):
            intervals_by_row.setdefault(row, []).append((cell.c1, cell.c2))
    for row, intervals in intervals_by_row.items():
        intervals_by_row[row] = _merge_intervals(intervals)
    return intervals_by_row


def _merge_intervals(intervals: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Merge overlapping or adjacent intervals."""
    if not intervals:
        return []
    sorted_intervals = sorted(intervals)
    merged: list[tuple[int, int]] = []
    current_start, current_end = sorted_intervals[0]
    for start, end in sorted_intervals[1:]:
        if start <= current_end + 1:
            current_end = max(current_end, end)
            continue
        merged.append((current_start, current_end))
        current_start, current_end = start, end
    merged.append((current_start, current_end))
    return merged


def _col_in_intervals(col_index: int, intervals: list[tuple[int, int]]) -> bool:
    """Check whether a column index falls in any interval."""
    for start, end in intervals:
        if col_index < start:
            return False
        if start <= col_index <= end:
            return True
    return False


def _safe_col_index(col_key: str) -> int | None:
    """Parse a column key to int, returning None on failure."""
    try:
        return int(col_key)
    except ValueError:
        return None


def collect_sheet_raw_data(
    *,
    cell_data: CellData,
    shape_data: ShapeData,
    chart_data: ChartData,
    merged_cell_data: MergedCellData,
    workbook: xw.Book,
    mode: ExtractionMode = "standard",
    include_merged_values_in_rows: bool,
    print_area_data: PrintAreaData | None = None,
    auto_page_break_data: PrintAreaData | None = None,
    colors_map_data: WorkbookColorsMap | None = None,
) -> dict[str, SheetRawData]:
    """Collect per-sheet raw data from extraction artifacts.

    Args:
        cell_data: Extracted cell rows per sheet.
        shape_data: Extracted shapes per sheet.
        chart_data: Extracted charts per sheet.
        merged_cell_data: Extracted merged cells per sheet.
        workbook: xlwings workbook instance.
        mode: Extraction mode.
        print_area_data: Optional print area data per sheet.
        auto_page_break_data: Optional auto page-break data per sheet.
        colors_map_data: Optional colors map data.
        include_merged_values_in_rows: Whether to keep merged values in rows.

    Returns:
        Mapping of sheet name to raw sheet data.
    """
    result: dict[str, SheetRawData] = {}
    for sheet_name, rows in cell_data.items():
        sheet = workbook.sheets[sheet_name]
        merged_cells = merged_cell_data.get(sheet_name, [])
        filtered_rows = (
            rows
            if include_merged_values_in_rows
            else _filter_rows_excluding_merged_values(rows, merged_cells)
        )
        sheet_raw = SheetRawData(
            rows=filtered_rows,
            shapes=shape_data.get(sheet_name, []),
            charts=chart_data.get(sheet_name, []) if mode != "light" else [],
            table_candidates=detect_tables(sheet),
            print_areas=print_area_data.get(sheet_name, []) if print_area_data else [],
            auto_print_areas=auto_page_break_data.get(sheet_name, [])
            if auto_page_break_data
            else [],
            colors_map=_resolve_sheet_colors_map(colors_map_data, sheet_name),
            merged_cells=merged_cells,
        )
        result[sheet_name] = sheet_raw
    return result


def run_extraction_pipeline(inputs: ExtractionInputs) -> PipelineResult:
    """Run the full extraction pipeline and return the result.

    Args:
        inputs: Resolved pipeline inputs.

    Returns:
        PipelineResult with workbook data, artifacts, and execution state.
    """
    plan = build_pipeline_plan(inputs)
    artifacts = run_pipeline(plan.pre_com_steps, inputs, ExtractionArtifacts())
    state = PipelineState()

    def _fallback(message: str, reason: FallbackReason) -> PipelineResult:
        state.fallback_reason = reason
        log_fallback(logger, reason, message)
        workbook = build_cells_tables_workbook(
            inputs=inputs,
            artifacts=artifacts,
            reason=message,
        )
        return PipelineResult(workbook=workbook, artifacts=artifacts, state=state)

    if not plan.use_com:
        return _fallback("Light mode selected.", FallbackReason.LIGHT_MODE)

    if os.getenv("SKIP_COM_TESTS"):
        return _fallback(
            "SKIP_COM_TESTS is set; skipping COM/xlwings access.",
            FallbackReason.SKIP_COM_TESTS,
        )

    try:
        with xlwings_workbook(inputs.file_path) as workbook:
            state.com_attempted = True
            try:
                run_com_pipeline(plan.com_steps, inputs, artifacts, workbook)
                raw_sheets = collect_sheet_raw_data(
                    cell_data=artifacts.cell_data,
                    shape_data=artifacts.shape_data,
                    chart_data=artifacts.chart_data,
                    merged_cell_data=artifacts.merged_cell_data,
                    workbook=workbook,
                    mode=inputs.mode,
                    include_merged_values_in_rows=inputs.include_merged_values_in_rows,
                    print_area_data=artifacts.print_area_data
                    if inputs.include_print_areas
                    else None,
                    auto_page_break_data=artifacts.auto_page_break_data
                    if inputs.include_auto_page_breaks
                    else None,
                    colors_map_data=artifacts.colors_map_data,
                )
                raw_workbook = WorkbookRawData(
                    book_name=inputs.file_path.name, sheets=raw_sheets
                )
                state.com_succeeded = True
                return PipelineResult(
                    workbook=build_workbook_data(raw_workbook),
                    artifacts=artifacts,
                    state=state,
                )
            except Exception as exc:
                return _fallback(
                    f"COM pipeline failed ({exc!r}).",
                    FallbackReason.COM_PIPELINE_FAILED,
                )
    except Exception as exc:
        return _fallback(
            f"xlwings/Excel COM is unavailable. ({exc!r})",
            FallbackReason.COM_UNAVAILABLE,
        )


def build_cells_tables_workbook(
    *,
    inputs: ExtractionInputs,
    artifacts: ExtractionArtifacts,
    reason: str,
) -> WorkbookData:
    """Build a WorkbookData containing cells + table_candidates (fallback).

    Args:
        inputs: Pipeline inputs.
        artifacts: Collected artifacts from extraction steps.
        reason: Reason to log for fallback.

    Returns:
        WorkbookData constructed from cells and detected tables.
    """
    logger.debug("Building fallback workbook: %s", reason)
    backend = OpenpyxlBackend(inputs.file_path)
    colors_map_data = artifacts.colors_map_data
    if inputs.include_colors_map and colors_map_data is None:
        colors_map_data = backend.extract_colors_map(
            include_default_background=inputs.include_default_background,
            ignore_colors=inputs.ignore_colors,
        )
    sheets: dict[str, SheetRawData] = {}
    for sheet_name, rows in artifacts.cell_data.items():
        sheet_colors = (
            colors_map_data.get_sheet(sheet_name) if colors_map_data else None
        )
        tables = backend.detect_tables(sheet_name)
        merged_cells = artifacts.merged_cell_data.get(sheet_name, [])
        filtered_rows = (
            rows
            if inputs.include_merged_values_in_rows
            else _filter_rows_excluding_merged_values(rows, merged_cells)
        )
        sheets[sheet_name] = SheetRawData(
            rows=filtered_rows,
            shapes=[],
            charts=[],
            table_candidates=tables,
            print_areas=artifacts.print_area_data.get(sheet_name, [])
            if inputs.include_print_areas
            else [],
            auto_print_areas=[],
            colors_map=sheet_colors.colors_map if sheet_colors else {},
            merged_cells=merged_cells,
        )
    raw = WorkbookRawData(book_name=inputs.file_path.name, sheets=sheets)
    return build_workbook_data(raw)
