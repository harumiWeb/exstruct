from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Literal

import logging

import xlwings as xw

from ..models import CellRow, Shape, SheetData, WorkbookData
from .cells import (
    detect_tables,
    detect_tables_openpyxl,
    extract_sheet_cells,
)
from .charts import get_charts
from .shapes import get_shapes_with_position

logger = logging.getLogger(__name__)


def integrate_sheet_content(
    cell_data: Dict[str, List[CellRow]],
    shape_data: Dict[str, List[Shape]],
    workbook: xw.Book,
    mode: Literal["light", "standard", "verbose"] = "standard",
) -> Dict[str, SheetData]:
    """Integrate cells, shapes, charts, and tables into SheetData per sheet."""
    result: Dict[str, SheetData] = {}
    for sheet_name, rows in cell_data.items():
        sheet_shapes = shape_data.get(sheet_name, [])
        sheet = workbook.sheets[sheet_name]

        sheet_model = SheetData(
            rows=rows,
            shapes=sheet_shapes,
            charts=[] if mode == "light" else get_charts(sheet),
            tables=detect_tables(sheet),
        )

        result[sheet_name] = sheet_model
    return result


def extract_workbook(
    file_path: Path, mode: Literal["light", "standard", "verbose"] = "standard"
) -> WorkbookData:
    """Extract workbook and return WorkbookData; fallback to cells+tables if Excel COM is unavailable."""
    cell_data = extract_sheet_cells(file_path)

    def _cells_and_tables_only(reason: str) -> WorkbookData:
        sheets: Dict[str, SheetData] = {}
        for sheet_name, rows in cell_data.items():
            try:
                tables = detect_tables_openpyxl(file_path, sheet_name)
            except Exception:
                tables = []
            sheets[sheet_name] = SheetData(
                rows=rows,
                shapes=[],
                charts=[],
                tables=tables,
            )
        logger.warning(
            "%s Falling back to cells+tables only; shapes and charts will be empty.",
            reason,
        )
        return WorkbookData(book_name=file_path.name, sheets=sheets)

    if mode == "light":
        return _cells_and_tables_only("Light mode selected.")

    try:
        wb = xw.Book(file_path)
    except Exception as e:
        return _cells_and_tables_only(
            f"xlwings/Excel COM is unavailable. ({e!r})"
        )

    try:
        try:
            shape_data = get_shapes_with_position(wb, mode=mode)
            merged = integrate_sheet_content(cell_data, shape_data, wb, mode=mode)
            return WorkbookData(book_name=file_path.name, sheets=merged)
        except Exception as e:
            return _cells_and_tables_only(
                f"Shape extraction failed ({e!r})."
            )
    finally:
        wb.close()
