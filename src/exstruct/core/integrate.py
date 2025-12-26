from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Literal

import xlwings as xw

from ..errors import FallbackReason
from ..models import CellRow, PrintArea, Shape, SheetData, WorkbookData
from .backends.com_backend import ComBackend
from .cells import WorkbookColorsMap, detect_tables, extract_sheet_colors_map
from .charts import get_charts
from .logging_utils import log_fallback
from .pipeline import (
    ExtractionArtifacts,
    ExtractionInputs,
    build_cells_tables_workbook,
    build_pre_com_pipeline,
    run_pipeline,
)
from .shapes import get_shapes_with_position
from .workbook import xlwings_workbook

logger = logging.getLogger(__name__)
_ALLOWED_MODES: set[str] = {"light", "standard", "verbose"}


def integrate_sheet_content(
    cell_data: dict[str, list[CellRow]],
    shape_data: dict[str, list[Shape]],
    workbook: xw.Book,
    mode: Literal["light", "standard", "verbose"] = "standard",
    print_area_data: dict[str, list[PrintArea]] | None = None,
    auto_page_break_data: dict[str, list[PrintArea]] | None = None,
    colors_map_data: WorkbookColorsMap | None = None,
) -> dict[str, SheetData]:
    """Integrate cells, shapes, charts, and tables into SheetData per sheet.

    Args:
        cell_data: Extracted cell rows per sheet.
        shape_data: Extracted shapes per sheet.
        workbook: xlwings workbook instance.
        mode: Extraction mode.
        print_area_data: Optional print area data per sheet.
        auto_page_break_data: Optional auto page-break data per sheet.
        colors_map_data: Optional colors map data.

    Returns:
        Mapping of sheet name to SheetData.
    """
    result: dict[str, SheetData] = {}
    for sheet_name, rows in cell_data.items():
        sheet_shapes = shape_data.get(sheet_name, [])
        sheet = workbook.sheets[sheet_name]
        sheet_colors = (
            colors_map_data.get_sheet(sheet_name) if colors_map_data else None
        )

        sheet_model = SheetData(
            rows=rows,
            shapes=sheet_shapes,
            charts=[] if mode == "light" else get_charts(sheet, mode=mode),
            table_candidates=detect_tables(sheet),
            print_areas=print_area_data.get(sheet_name, []) if print_area_data else [],
            auto_print_areas=auto_page_break_data.get(sheet_name, [])
            if auto_page_break_data
            else [],
            colors_map=sheet_colors.colors_map if sheet_colors else {},
        )

        result[sheet_name] = sheet_model
    return result


def extract_workbook(  # noqa: C901
    file_path: str | Path,
    mode: Literal["light", "standard", "verbose"] = "standard",
    *,
    include_cell_links: bool = False,
    include_print_areas: bool = True,
    include_auto_page_breaks: bool = False,
    include_colors_map: bool = False,
    include_default_background: bool = False,
    ignore_colors: set[str] | None = None,
) -> WorkbookData:
    """Extract workbook and return WorkbookData.

    Falls back to cells+tables if Excel COM is unavailable.

    Args:
        file_path: Workbook path.
        mode: Extraction mode.
        include_cell_links: Whether to include cell hyperlinks.
        include_print_areas: Whether to include print areas.
        include_auto_page_breaks: Whether to include auto page breaks.
        include_colors_map: Whether to include colors map.
        include_default_background: Whether to include default background color.
        ignore_colors: Optional set of color keys to ignore.

    Returns:
        Extracted WorkbookData.

    Raises:
        ValueError: If mode is unsupported.
    """
    if mode not in _ALLOWED_MODES:
        raise ValueError(f"Unsupported mode: {mode}")

    normalized_file_path = file_path if isinstance(file_path, Path) else Path(file_path)

    inputs = ExtractionInputs(
        file_path=normalized_file_path,
        mode=mode,
        include_cell_links=include_cell_links,
        include_print_areas=include_print_areas,
        include_auto_page_breaks=include_auto_page_breaks,
        include_colors_map=include_colors_map,
        include_default_background=include_default_background,
        ignore_colors=ignore_colors,
    )
    artifacts = run_pipeline(
        build_pre_com_pipeline(inputs),
        inputs,
        ExtractionArtifacts(),
    )

    def _cells_and_tables_only(reason: str, code: FallbackReason) -> WorkbookData:
        log_fallback(
            logger,
            code,
            f"{reason} Falling back to cells+tables only; shapes and charts will be empty.",
        )
        return build_cells_tables_workbook(
            inputs=inputs,
            artifacts=artifacts,
            reason=reason,
        )

    if mode == "light":
        return _cells_and_tables_only("Light mode selected.", FallbackReason.LIGHT_MODE)

    if os.getenv("SKIP_COM_TESTS"):
        return _cells_and_tables_only(
            "SKIP_COM_TESTS is set; skipping COM/xlwings access.",
            FallbackReason.SKIP_COM_TESTS,
        )

    try:
        with xlwings_workbook(normalized_file_path) as wb:
            try:
                com_backend = ComBackend(wb)
                if include_colors_map and artifacts.colors_map_data is None:
                    artifacts.colors_map_data = com_backend.extract_colors_map(
                        include_default_background=include_default_background,
                        ignore_colors=ignore_colors,
                    )
                    if artifacts.colors_map_data is None:
                        try:
                            artifacts.colors_map_data = extract_sheet_colors_map(
                                normalized_file_path,
                                include_default_background=include_default_background,
                                ignore_colors=ignore_colors,
                            )
                        except Exception as fallback_exc:
                            logger.warning(
                                "Color map extraction failed; skipping colors_map. (%r)",
                                fallback_exc,
                            )
                            artifacts.colors_map_data = None
                shape_data = get_shapes_with_position(wb, mode=mode)
                if include_print_areas and not artifacts.print_area_data:
                    # openpyxl couldn't read (e.g., .xls). Try COM as a fallback.
                    try:
                        artifacts.print_area_data = com_backend.extract_print_areas()
                    except Exception:
                        artifacts.print_area_data = {}
                if include_auto_page_breaks:
                    try:
                        artifacts.auto_page_break_data = (
                            com_backend.extract_auto_page_breaks()
                        )
                    except Exception:
                        artifacts.auto_page_break_data = {}
                merged = integrate_sheet_content(
                    artifacts.cell_data,
                    shape_data,
                    wb,
                    mode=mode,
                    print_area_data=artifacts.print_area_data
                    if include_print_areas
                    else None,
                    auto_page_break_data=artifacts.auto_page_break_data
                    if include_auto_page_breaks
                    else None,
                    colors_map_data=artifacts.colors_map_data,
                )
                return WorkbookData(book_name=normalized_file_path.name, sheets=merged)
            except Exception as e:
                logger.warning(
                    "Shape extraction failed; falling back to cells+tables. (%r)", e
                )
                return _cells_and_tables_only(
                    f"Shape extraction failed ({e!r}).",
                    FallbackReason.SHAPE_EXTRACTION_FAILED,
                )
    except Exception as e:
        return _cells_and_tables_only(
            f"xlwings/Excel COM is unavailable. ({e!r})",
            FallbackReason.COM_UNAVAILABLE,
        )
