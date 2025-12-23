from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any, Literal, cast

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
import xlwings as xw

from ..models import CellRow, PrintArea, Shape, SheetData, WorkbookData
from .cells import (
    WorkbookColorsMap,
    detect_tables,
    detect_tables_openpyxl,
    extract_sheet_cells,
    extract_sheet_cells_with_links,
    extract_sheet_colors_map,
    extract_sheet_colors_map_com,
)
from .charts import get_charts
from .shapes import get_shapes_with_position

logger = logging.getLogger(__name__)
_ALLOWED_MODES: set[str] = {"light", "standard", "verbose"}


def _find_open_workbook(file_path: Path) -> xw.Book | None:
    """Return an existing workbook if already open in Excel; otherwise None."""
    try:
        for app in xw.apps:
            for wb in app.books:
                try:
                    if Path(wb.fullname).resolve() == file_path.resolve():
                        return wb
                except Exception:
                    continue
    except Exception:
        return None
    return None


def _open_workbook(file_path: Path) -> tuple[xw.Book, bool]:
    """
    Open workbook:
    - If already open, reuse and do not close Excel on exit.
    - Otherwise create invisible Excel (visible=False) and close when done.
    Returns (workbook, should_close_app).
    """
    existing = _find_open_workbook(file_path)
    if existing:
        return existing, False
    app = xw.App(add_book=False, visible=False)
    wb = app.books.open(str(file_path))
    return wb, True


def _parse_print_area_range(
    range_str: str, *, zero_based: bool = True
) -> tuple[int, int, int, int] | None:
    """
    Parse an Excel range string into (r1, c1, r2, c2). Returns None on failure.
    """
    cleaned = range_str.strip()
    if not cleaned:
        return None
    if "!" in cleaned:
        cleaned = cleaned.split("!", 1)[1]
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cleaned)
    except Exception:
        return None
    if zero_based:
        return (min_row - 1, min_col - 1, max_row - 1, max_col - 1)
    return (min_row, min_col, max_row, max_col)


def _extract_print_areas_openpyxl(  # noqa: C901
    file_path: Path,
) -> dict[str, list[PrintArea]]:
    """
    Extract print areas per sheet using openpyxl defined names.

    Returns {sheet_name: [PrintArea, ...]}.
    """
    try:
        wb = load_workbook(file_path, data_only=True, read_only=True)
    except Exception:
        return {}

    try:
        defined = wb.defined_names.get("_xlnm.Print_Area")
        areas: dict[str, list[PrintArea]] = {}
        if defined:
            for sheet_name, range_str in defined.destinations:
                if sheet_name not in wb.sheetnames:
                    continue
                # A single destination can contain multiple comma-separated ranges.
                for part in str(range_str).split(","):
                    parsed = _parse_print_area_range(part)
                    if not parsed:
                        continue
                    r1, c1, r2, c2 = parsed
                    areas.setdefault(sheet_name, []).append(
                        PrintArea(r1=r1, c1=c1, r2=r2, c2=c2)
                    )
        # Fallback: some files carry sheet-level print_area without defined name.
        if not areas:
            for ws in wb.worksheets:
                pa = getattr(ws, "_print_area", None)
                if not pa:
                    continue
                for part in str(pa).split(","):
                    parsed = _parse_print_area_range(part)
                    if not parsed:
                        continue
                    r1, c1, r2, c2 = parsed
                    areas.setdefault(ws.title, []).append(
                        PrintArea(r1=r1, c1=c1, r2=r2, c2=c2)
                    )
        return areas
    finally:
        try:
            wb.close()
        except Exception:
            pass


def _extract_print_areas_com(workbook: xw.Book) -> dict[str, list[PrintArea]]:
    """
    Extract print areas per sheet via xlwings/COM.

    Uses Sheet.PageSetup.PrintArea which may contain comma-separated ranges.
    """
    areas: dict[str, list[PrintArea]] = {}
    for sheet in workbook.sheets:
        try:
            raw = sheet.api.PageSetup.PrintArea or ""
        except Exception:
            continue
        if not raw:
            continue
        parts = str(raw).split(",")
        for part in parts:
            parsed = _parse_print_area_range(part, zero_based=True)
            if not parsed:
                continue
            r1, c1, r2, c2 = parsed
            areas.setdefault(sheet.name, []).append(
                PrintArea(r1=r1, c1=c1, r2=r2, c2=c2)
            )
    return areas


def _normalize_area_for_sheet(part: str, ws_name: str) -> str | None:
    """
    Strip sheet name from a range part when it matches the target sheet; otherwise None.
    """
    s = part.strip()
    if "!" not in s:
        return s
    sheet, rng = s.rsplit("!", 1)
    sheet = sheet.strip()
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return rng if sheet == ws_name else None


def _split_csv_respecting_quotes(raw: str) -> list[str]:
    """
    Split a CSV-like string while keeping commas inside single quotes intact.
    """
    parts: list[str] = []
    buf: list[str] = []
    in_quote = False
    i = 0
    while i < len(raw):
        ch = raw[i]
        if ch == "'":
            if in_quote and i + 1 < len(raw) and raw[i + 1] == "'":
                buf.append("''")
                i += 2
                continue
            in_quote = not in_quote
            buf.append(ch)
            i += 1
            continue
        if ch == "," and not in_quote:
            parts.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    if buf:
        parts.append("".join(buf).strip())
    return [p for p in parts if p]


def _compute_auto_page_break_areas(workbook: xw.Book) -> dict[str, list[PrintArea]]:
    """
    Compute auto page-break rectangles per sheet using Excel COM.
    Falls back to empty dict on failure.
    """
    results: dict[str, list[PrintArea]] = {}
    for sheet in workbook.sheets:
        try:
            ws_api = cast(Any, sheet.api)  # xlwings COM API; treated as Any
            original_display: bool | None = ws_api.DisplayPageBreaks
            ws_api.DisplayPageBreaks = True
            print_area = ws_api.PageSetup.PrintArea or ws_api.UsedRange.Address
            parts_raw = _split_csv_respecting_quotes(str(print_area))
            area_parts: list[str] = []
            for part in parts_raw:
                rng = _normalize_area_for_sheet(part, sheet.name)
                if rng:
                    area_parts.append(rng)
            hpb = cast(Any, ws_api.HPageBreaks)
            vpb = cast(Any, ws_api.VPageBreaks)
            h_break_rows = [
                hpb.Item(i).Location.Row for i in range(1, int(hpb.Count) + 1)
            ]
            v_break_cols = [
                vpb.Item(i).Location.Column for i in range(1, int(vpb.Count) + 1)
            ]
            for addr in area_parts:
                range_obj = cast(Any, ws_api.Range(addr))
                min_row = int(range_obj.Row)
                max_row = min_row + int(range_obj.Rows.Count) - 1
                min_col = int(range_obj.Column)
                max_col = min_col + int(range_obj.Columns.Count) - 1
                rows = (
                    [min_row]
                    + [r for r in h_break_rows if min_row < r <= max_row]
                    + [max_row + 1]
                )
                cols = (
                    [min_col]
                    + [c for c in v_break_cols if min_col < c <= max_col]
                    + [max_col + 1]
                )
                for i in range(len(rows) - 1):
                    r1, r2 = rows[i], rows[i + 1] - 1
                    for j in range(len(cols) - 1):
                        c1, c2 = cols[j], cols[j + 1] - 1
                        c1_0 = c1 - 1
                        c2_0 = c2 - 1
                        results.setdefault(sheet.name, []).append(
                            PrintArea(r1=r1, c1=c1_0, r2=r2, c2=c2_0)
                        )
            if original_display is not None:
                ws_api.DisplayPageBreaks = original_display
        except Exception:
            try:
                if original_display is not None:
                    ws_api.DisplayPageBreaks = original_display
            except Exception:
                pass
            continue
    return results


def integrate_sheet_content(
    cell_data: dict[str, list[CellRow]],
    shape_data: dict[str, list[Shape]],
    workbook: xw.Book,
    mode: Literal["light", "standard", "verbose"] = "standard",
    print_area_data: dict[str, list[PrintArea]] | None = None,
    auto_page_break_data: dict[str, list[PrintArea]] | None = None,
    colors_map_data: WorkbookColorsMap | None = None,
) -> dict[str, SheetData]:
    """Integrate cells, shapes, charts, and tables into SheetData per sheet."""
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
) -> WorkbookData:
    """Extract workbook and return WorkbookData; fallback to cells+tables if Excel COM is unavailable."""
    if mode not in _ALLOWED_MODES:
        raise ValueError(f"Unsupported mode: {mode}")

    normalized_file_path = file_path if isinstance(file_path, Path) else Path(file_path)

    colors_map_data: WorkbookColorsMap | None = None
    if include_colors_map and (mode == "light" or os.getenv("SKIP_COM_TESTS")):
        try:
            colors_map_data = extract_sheet_colors_map(
                normalized_file_path,
                include_default_background=include_default_background,
            )
        except Exception as exc:
            logger.warning(
                "Color map extraction failed; skipping colors_map. (%r)", exc
            )
            colors_map_data = None
    cell_data = (
        extract_sheet_cells_with_links(normalized_file_path)
        if include_cell_links
        else extract_sheet_cells(normalized_file_path)
    )
    print_area_data: dict[str, list[PrintArea]] = {}
    if include_print_areas:
        print_area_data = _extract_print_areas_openpyxl(normalized_file_path)
    auto_page_break_data: dict[str, list[PrintArea]] = {}

    def _cells_and_tables_only(reason: str) -> WorkbookData:
        sheets: dict[str, SheetData] = {}
        fallback_colors = colors_map_data
        if include_colors_map and fallback_colors is None:
            try:
                fallback_colors = extract_sheet_colors_map(
                    normalized_file_path,
                    include_default_background=include_default_background,
                )
            except Exception as exc:
                logger.warning(
                    "Color map extraction failed; skipping colors_map. (%r)", exc
                )
                fallback_colors = None
        for sheet_name, rows in cell_data.items():
            sheet_colors = (
                fallback_colors.get_sheet(sheet_name) if fallback_colors else None
            )
            try:
                tables = detect_tables_openpyxl(normalized_file_path, sheet_name)
            except Exception:
                tables = []
            sheets[sheet_name] = SheetData(
                rows=rows,
                shapes=[],
                charts=[],
                table_candidates=tables,
                print_areas=print_area_data.get(sheet_name, [])
                if include_print_areas
                else [],
                auto_print_areas=[],
                colors_map=sheet_colors.colors_map if sheet_colors else {},
            )
        logger.warning(
            "%s Falling back to cells+tables only; shapes and charts will be empty.",
            reason,
        )
        return WorkbookData(book_name=normalized_file_path.name, sheets=sheets)

    if mode == "light":
        return _cells_and_tables_only("Light mode selected.")

    if os.getenv("SKIP_COM_TESTS"):
        return _cells_and_tables_only(
            "SKIP_COM_TESTS is set; skipping COM/xlwings access."
        )

    try:
        wb, close_app = _open_workbook(normalized_file_path)
    except Exception as e:
        return _cells_and_tables_only(f"xlwings/Excel COM is unavailable. ({e!r})")

    try:
        try:
            if include_colors_map and colors_map_data is None:
                try:
                    colors_map_data = extract_sheet_colors_map_com(
                        wb, include_default_background=include_default_background
                    )
                except Exception as exc:
                    logger.warning(
                        "COM color map extraction failed; falling back to openpyxl. (%r)",
                        exc,
                    )
                    try:
                        colors_map_data = extract_sheet_colors_map(
                            normalized_file_path,
                            include_default_background=include_default_background,
                        )
                    except Exception as fallback_exc:
                        logger.warning(
                            "Color map extraction failed; skipping colors_map. (%r)",
                            fallback_exc,
                        )
                        colors_map_data = None
            shape_data = get_shapes_with_position(wb, mode=mode)
            if include_print_areas and not print_area_data:
                # openpyxl couldn't read (e.g., .xls). Try COM as a fallback.
                try:
                    print_area_data = _extract_print_areas_com(wb)
                except Exception:
                    print_area_data = {}
            if include_auto_page_breaks:
                try:
                    auto_page_break_data = _compute_auto_page_break_areas(wb)
                except Exception:
                    auto_page_break_data = {}
            merged = integrate_sheet_content(
                cell_data,
                shape_data,
                wb,
                mode=mode,
                print_area_data=print_area_data if include_print_areas else None,
                auto_page_break_data=auto_page_break_data
                if include_auto_page_breaks
                else None,
                colors_map_data=colors_map_data,
            )
            return WorkbookData(book_name=normalized_file_path.name, sheets=merged)
        except Exception as e:
            logger.warning(
                "Shape extraction failed; falling back to cells+tables. (%r)", e
            )
            return _cells_and_tables_only(f"Shape extraction failed ({e!r}).")
    finally:
        # Close only if we created the app to avoid shutting user sessions.
        try:
            if close_app:
                app = wb.app
                wb.close()
                app.quit()
        except Exception:
            pass
