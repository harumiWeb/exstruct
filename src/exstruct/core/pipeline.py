from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
import os
from pathlib import Path
from typing import Literal

from ..models import CellRow, PrintArea, SheetData, WorkbookData
from .backends.openpyxl_backend import OpenpyxlBackend
from .cells import WorkbookColorsMap

ExtractionMode = Literal["light", "standard", "verbose"]
CellData = dict[str, list[CellRow]]
PrintAreaData = dict[str, list[PrintArea]]


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
    """

    file_path: Path
    mode: ExtractionMode
    include_cell_links: bool
    include_print_areas: bool
    include_auto_page_breaks: bool
    include_colors_map: bool
    include_default_background: bool
    ignore_colors: set[str] | None


@dataclass
class ExtractionArtifacts:
    """Mutable artifacts collected by pipeline steps.

    Attributes:
        cell_data: Extracted cell rows per sheet.
        print_area_data: Extracted print areas per sheet.
        auto_page_break_data: Extracted auto page-break areas per sheet.
        colors_map_data: Extracted colors map for workbook sheets.
    """

    cell_data: CellData = field(default_factory=dict)
    print_area_data: PrintAreaData = field(default_factory=dict)
    auto_page_break_data: PrintAreaData = field(default_factory=dict)
    colors_map_data: WorkbookColorsMap | None = None


ExtractionStep = Callable[[ExtractionInputs, ExtractionArtifacts], None]


@dataclass(frozen=True)
class PipelinePlan:
    """Resolved pipeline plan for an extraction run.

    Attributes:
        pre_com_steps: Ordered list of steps to run before COM access.
        use_com: Whether COM-based extraction should be attempted.
    """

    pre_com_steps: list[ExtractionStep]
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

    return ExtractionInputs(
        file_path=normalized_file_path,
        mode=mode,
        include_cell_links=resolved_cell_links,
        include_print_areas=resolved_print_areas,
        include_auto_page_breaks=include_auto_page_breaks,
        include_colors_map=resolved_colors_map,
        include_default_background=resolved_default_background,
        ignore_colors=resolved_ignore_colors,
    )


def build_pipeline_plan(inputs: ExtractionInputs) -> PipelinePlan:
    """Build a pipeline plan based on resolved inputs.

    Args:
        inputs: Resolved pipeline inputs.

    Returns:
        PipelinePlan containing pre-COM steps and COM usage flag.
    """
    return PipelinePlan(
        pre_com_steps=build_pre_com_pipeline(inputs),
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
        ),
    }
    steps: list[ExtractionStep] = []
    for config in step_table[inputs.mode]:
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
    backend = OpenpyxlBackend(inputs.file_path)
    colors_map_data = artifacts.colors_map_data
    if inputs.include_colors_map and colors_map_data is None:
        colors_map_data = backend.extract_colors_map(
            include_default_background=inputs.include_default_background,
            ignore_colors=inputs.ignore_colors,
        )
    sheets: dict[str, SheetData] = {}
    for sheet_name, rows in artifacts.cell_data.items():
        sheet_colors = (
            colors_map_data.get_sheet(sheet_name) if colors_map_data else None
        )
        tables = backend.detect_tables(sheet_name)
        sheets[sheet_name] = SheetData(
            rows=rows,
            shapes=[],
            charts=[],
            table_candidates=tables,
            print_areas=artifacts.print_area_data.get(sheet_name, [])
            if inputs.include_print_areas
            else [],
            auto_print_areas=[],
            colors_map=sheet_colors.colors_map if sheet_colors else {},
        )
    return WorkbookData(book_name=inputs.file_path.name, sheets=sheets)
