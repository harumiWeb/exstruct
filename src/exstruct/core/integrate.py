from __future__ import annotations

from pathlib import Path
from typing import Dict, List

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
) -> Dict[str, SheetData]:
    """Integrate cells, shapes, charts, and tables into SheetData per sheet."""
    result: Dict[str, SheetData] = {}
    for sheet_name, rows in cell_data.items():
        sheet_shapes = shape_data.get(sheet_name, [])
        sheet = workbook.sheets[sheet_name]

        sheet_model = SheetData(
            rows=rows,
            shapes=sheet_shapes,
            charts=get_charts(sheet),
            tables=detect_tables(sheet),
        )

        result[sheet_name] = sheet_model
    return result


def extract_workbook(file_path: Path) -> WorkbookData:
    """Extract workbook and return WorkbookData; fallback to cells+tables if Excel COM is unavailable."""
    cell_data = extract_sheet_cells(file_path)
    try:
        wb = xw.Book(file_path)
    except Exception as e:
        logger.warning(
            "xlwings/Excel COM is unavailable. Falling back to cells+tables only. "
            "Shapes and charts will be empty. (%r)",
            e,
        )
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
        return WorkbookData(book_name=file_path.name, sheets=sheets)

    try:
        shape_data = get_shapes_with_position(wb)
        merged = integrate_sheet_content(cell_data, shape_data, wb)
    finally:
        wb.close()

    return WorkbookData(book_name=file_path.name, sheets=merged)
