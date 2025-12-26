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


def build_pre_com_pipeline(inputs: ExtractionInputs) -> list[ExtractionStep]:
    """Build pipeline steps that run before COM/Excel access.

    Args:
        inputs: Pipeline inputs describing extraction flags.

    Returns:
        Ordered list of extraction steps to run before COM.
    """
    steps: list[ExtractionStep] = [step_extract_cells]
    if inputs.include_print_areas:
        steps.append(step_extract_print_areas_openpyxl)
    if inputs.include_colors_map and (
        inputs.mode == "light" or os.getenv("SKIP_COM_TESTS")
    ):
        steps.append(step_extract_colors_map_openpyxl)
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
