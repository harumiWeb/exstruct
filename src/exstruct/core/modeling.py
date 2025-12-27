from __future__ import annotations

from dataclasses import dataclass

from ..models import CellRow, Chart, PrintArea, Shape, SheetData, WorkbookData


@dataclass(frozen=True)
class SheetRawData:
    """Raw, extracted sheet data before model conversion.

    Attributes:
        rows: Extracted cell rows.
        shapes: Extracted shapes.
        charts: Extracted charts.
        table_candidates: Detected table ranges.
        print_areas: Extracted print areas.
        auto_print_areas: Extracted auto page-break areas.
        colors_map: Mapping of color keys to (row, column) positions.
    """

    rows: list[CellRow]
    shapes: list[Shape]
    charts: list[Chart]
    table_candidates: list[str]
    print_areas: list[PrintArea]
    auto_print_areas: list[PrintArea]
    colors_map: dict[str, list[tuple[int, int]]]


@dataclass(frozen=True)
class WorkbookRawData:
    """Raw, extracted workbook data before model conversion.

    Attributes:
        book_name: Workbook file name.
        sheets: Mapping of sheet name to raw sheet data.
    """

    book_name: str
    sheets: dict[str, SheetRawData]


def build_sheet_data(raw: SheetRawData) -> SheetData:
    """Build a SheetData model from raw sheet data.

    Args:
        raw: Raw sheet data.

    Returns:
        SheetData model instance.
    """
    return SheetData(
        rows=raw.rows,
        shapes=raw.shapes,
        charts=raw.charts,
        table_candidates=raw.table_candidates,
        print_areas=raw.print_areas,
        auto_print_areas=raw.auto_print_areas,
        colors_map=raw.colors_map,
    )


def build_workbook_data(raw: WorkbookRawData) -> WorkbookData:
    """Build a WorkbookData model from raw workbook data.

    Args:
        raw: Raw workbook data.

    Returns:
        WorkbookData model instance.
    """
    sheets = {name: build_sheet_data(sheet) for name, sheet in raw.sheets.items()}
    return WorkbookData(book_name=raw.book_name, sheets=sheets)
