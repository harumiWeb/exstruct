"""Openpyxl backend for Excel workbook extraction."""

from __future__ import annotations

from dataclasses import dataclass
import logging
from pathlib import Path
from typing import Literal

from ...models import PrintArea
from ..cells import (
    WorkbookColorsMap,
    WorkbookFormulasMap,
    detect_tables_openpyxl,
    extract_sheet_cells,
    extract_sheet_cells_with_links,
    extract_sheet_colors_map,
    extract_sheet_formulas_map,
    extract_sheet_merged_cells,
)
from ..ranges import parse_range_zero_based
from ..workbook import openpyxl_workbook
from .base import CellData, MergedCellData, PrintAreaData

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class OpenpyxlBackend:
    """Openpyxl-based backend for extraction tasks.

    Attributes:
        file_path: Path to the workbook file.
    """

    file_path: Path

    def extract_cells(self, *, include_links: bool) -> CellData:
        """Extract cell rows from the workbook.

        Args:
            include_links: Whether to include hyperlinks.

        Returns:
            Mapping of sheet name to cell rows.
        """
        return (
            extract_sheet_cells_with_links(self.file_path)
            if include_links
            else extract_sheet_cells(self.file_path)
        )

    def extract_print_areas(self) -> PrintAreaData:
        """Extract print areas per sheet using openpyxl defined names.

        Returns:
            Mapping of sheet name to print area list.
        """
        try:
            with openpyxl_workbook(
                self.file_path, data_only=True, read_only=False
            ) as wb:
                areas = _extract_print_areas_from_defined_names(wb)
                if not areas:
                    areas = _extract_print_areas_from_sheet_props(wb)
                return areas
        except Exception:
            return {}

    def extract_colors_map(
        self, *, include_default_background: bool, ignore_colors: set[str] | None
    ) -> WorkbookColorsMap | None:
        """Extract colors_map using openpyxl.

        Args:
            include_default_background: Whether to include default background colors.
            ignore_colors: Optional set of color keys to ignore.

        Returns:
            WorkbookColorsMap or None when extraction fails.
        """
        try:
            return extract_sheet_colors_map(
                self.file_path,
                include_default_background=include_default_background,
                ignore_colors=ignore_colors,
            )
        except Exception as exc:
            logger.warning(
                "Color map extraction failed; skipping colors_map. (%r)", exc
            )
            return None

    def extract_merged_cells(self) -> MergedCellData:
        """Extract merged cell ranges per sheet.

        Returns:
            Mapping of sheet name to merged cell ranges.
        """
        try:
            return extract_sheet_merged_cells(self.file_path)
        except Exception:
            return {}

    def extract_formulas_map(self) -> WorkbookFormulasMap | None:
        """
        Extract a mapping of workbook formulas for each sheet.

        Returns:
            WorkbookFormulasMap | None: A mapping from sheet name to its formulas, or `None` if extraction fails.
        """
        try:
            return extract_sheet_formulas_map(self.file_path)
        except Exception as exc:
            logger.warning(
                "Formula map extraction failed; skipping formulas_map. (%r)", exc
            )
            return None

    def detect_tables(
        self,
        sheet_name: str,
        *,
        mode: Literal["light", "standard", "verbose"] = "standard",
    ) -> list[str]:
        """
        Detects table candidate ranges within the specified worksheet.

        Parameters:
            sheet_name (str): Name of the worksheet to analyze for table candidates.
            mode (Literal["light", "standard", "verbose"]): Extraction mode, used to
                adjust scan limits in openpyxl-based detection.

        Returns:
            list[str]: Detected table candidate ranges as A1-style range strings; empty list if none are found or detection fails.
        """
        try:
            return detect_tables_openpyxl(self.file_path, sheet_name, mode=mode)
        except Exception:
            return []


def _extract_print_areas_from_defined_names(workbook: object) -> PrintAreaData:
    """Extract print areas from defined names in an openpyxl workbook.

    Args:
        workbook: openpyxl workbook instance.

    Returns:
        Mapping of sheet name to print area list.
    """
    defined = getattr(workbook, "defined_names", None)
    if defined is None:
        return {}
    defined_area = defined.get("_xlnm.Print_Area")
    if not defined_area:
        return {}

    areas: PrintAreaData = {}
    sheetnames = set(getattr(workbook, "sheetnames", []))
    for sheet_name, range_str in defined_area.destinations:
        if sheet_name not in sheetnames:
            continue
        _append_print_areas(areas, sheet_name, str(range_str))
    return areas


def _extract_print_areas_from_sheet_props(workbook: object) -> PrintAreaData:
    """Extract print areas from sheet-level print area properties.

    Args:
        workbook: openpyxl workbook instance.

    Returns:
        Mapping of sheet name to print area list.
    """
    areas: PrintAreaData = {}
    worksheets = getattr(workbook, "worksheets", [])
    for ws in worksheets:
        pa = getattr(ws, "_print_area", None)
        if not pa:
            continue
        _append_print_areas(areas, str(getattr(ws, "title", "")), str(pa))
    return areas


def _append_print_areas(areas: PrintAreaData, sheet_name: str, range_str: str) -> None:
    """Append parsed print areas to the mapping.

    Args:
        areas: Mapping to update.
        sheet_name: Target sheet name.
        range_str: Raw range string, possibly comma-separated.
    """
    for part in str(range_str).split(","):
        parsed = _parse_print_area_range(part)
        if not parsed:
            continue
        r1, c1, r2, c2 = parsed
        areas.setdefault(sheet_name, []).append(
            PrintArea(r1=r1 + 1, c1=c1, r2=r2 + 1, c2=c2)
        )


def _parse_print_area_range(range_str: str) -> tuple[int, int, int, int] | None:
    """Parse an Excel range string into zero-based coordinates.

    Args:
        range_str: Excel range string.

    Returns:
        Zero-based (r1, c1, r2, c2) tuple or None on failure.
    """
    bounds = parse_range_zero_based(range_str)
    if bounds is None:
        return None
    return (bounds.r1, bounds.c1, bounds.r2, bounds.c2)
