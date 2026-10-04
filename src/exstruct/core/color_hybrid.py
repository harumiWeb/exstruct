"""Conservative saved-fill extraction with Excel as the rendered-color oracle."""

from __future__ import annotations

from contextlib import ExitStack
import logging
from pathlib import Path
import posixpath
from time import perf_counter
from typing import TYPE_CHECKING
from zipfile import ZipFile

from defusedxml import ElementTree

from . import cells
from .cell_types import SheetColorsMap, WorkbookColorsMap
from .openpyxl_session import OpenpyxlExtractionSession

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet
    import xlwings as xw

logger = logging.getLogger(__name__)
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_RULES = {
    "expression",
    "cellIs",
    "colorScale",
    "dataBar",
    "iconSet",
    "top10",
    "uniqueValues",
    "duplicateValues",
    "containsText",
    "notContainsText",
    "beginsWith",
    "endsWith",
    "containsBlanks",
    "notContainsBlanks",
    "containsErrors",
    "notContainsErrors",
    "timePeriod",
    "aboveAverage",
}
Bounds = tuple[int, int, int, int]


def _worksheet_parts(archive: ZipFile) -> dict[str, str]:
    """Resolve worksheet names through package relationships, never sheet order."""
    rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {}
    for rel in rels:
        if rel.attrib.get("TargetMode") == "External":
            continue
        target = rel.attrib["Target"]
        targets[rel.attrib["Id"]] = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join("xl", target))
        )
    root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    return {
        sheet.attrib["name"]: targets[sheet.attrib[f"{{{_REL}}}id"]]
        for sheet in root.findall(f"{{{_MAIN}}}sheets/{{{_MAIN}}}sheet")
    }


def _conditional_ranges(raw: bytes) -> list[Bounds]:
    """Reject lossy extension parsing before trusting saved conditional ranges."""
    root = ElementTree.fromstring(raw)
    if root.tag != f"{{{_MAIN}}}worksheet":
        raise ValueError("unsupported worksheet namespace")
    ranges = []
    for element in root.iter():
        name = element.tag.rsplit("}", 1)[-1]
        if name in {"extLst", "AlternateContent", "pivotTableParts"}:
            raise ValueError("worksheet extensions may affect rendered fills")
        if name not in {"conditionalFormatting", "cfRule"}:
            continue
        if element.tag != f"{{{_MAIN}}}{name}":
            raise ValueError("unsupported conditional formatting namespace")
        if name == "cfRule":
            if element.attrib.get("type") not in _RULES:
                raise ValueError("unsupported conditional formatting rule")
        else:
            refs = element.attrib.get("sqref", "").split()
            if not refs:
                raise ValueError("missing conditional formatting range")
            for ref in refs:
                c1, r1, c2, r2 = cells.range_boundaries(ref)
                if not (1 <= r1 <= r2 <= 1048576 and 1 <= c1 <= c2 <= 16384):
                    raise ValueError("invalid conditional formatting range")
                ranges.append((r1, c1, r2, c2))
    return ranges


def _candidate_cells(ranges: list[Bounds], bounds: Bounds) -> set[tuple[int, int]]:
    """Clip first so entire-column rules cannot expand beyond UsedRange."""
    r1, c1, r2, c2 = bounds
    return {
        (row, col)
        for a, b, z, y in ranges
        for row in range(max(r1, a), min(r2, z) + 1)
        for col in range(max(c1, b), min(c2, y) + 1)
    }


def _static_colors(
    ws: Worksheet,
    bounds: Bounds,
    candidates: set[tuple[int, int]],
    include_default: bool,
) -> dict[tuple[int, int], str | None]:
    """Only trust fills whose RGB interpretation is identical to Excel's."""
    if ws.merged_cells.ranges or ws.tables:
        raise ValueError("merged cells or table styles require rendered scan")
    if any(d.has_style for d in ws.row_dimensions.values()) or any(
        d.has_style for d in ws.column_dimensions.values()
    ):
        raise ValueError("row/column styles require rendered scan")
    # An absent cell style inherits Normal; it may have a custom workbook fill.
    for style in ws.parent._named_styles:
        if style.builtinId == 0 and style.fill.patternType not in (None, "none"):
            raise ValueError("custom Normal fill requires rendered scan")
    r1, c1, r2, c2 = bounds
    result = {}
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            coord = (cell.row, cell.column)
            if coord in candidates:
                continue
            fill = cell.fill
            pattern = getattr(fill, "patternType", "unsupported")
            if pattern in (None, "none"):
                result[coord] = "FFFFFF" if include_default else None
            elif (
                pattern == "solid"
                and fill.fgColor.type == "rgb"
                and not fill.fgColor.tint
            ):
                result[coord] = cells._resolve_cell_background(cell, include_default)
            else:
                candidates.add(coord)
    return result


def _extract_sheet(
    sheet: xw.Sheet,
    session: OpenpyxlExtractionSession | None,
    archive: ZipFile | None,
    parts: dict[str, str],
    reason: str,
    include_default: bool,
    ignore_colors: set[str] | None,
) -> SheetColorsMap:
    started = perf_counter()
    token = cells._display_format_calls.set(0)
    static_seconds = 0.0
    rendered_seconds = 0.0
    candidates: set[tuple[int, int]] = set()
    conditional_count = 0
    used_cells = 0
    fallback = True
    try:
        used = sheet.used_range
        last = used.last_cell
        bounds = (
            int(getattr(used, "row", 1)),
            int(getattr(used, "column", 1)),
            int(last.row),
            int(last.column),
        )
        r1, c1, r2, c2 = bounds
        used_cells = max(0, r2 - r1 + 1) * max(0, c2 - c1 + 1)
        static_started = perf_counter()
        try:
            if session is None or archive is None:
                raise ValueError(reason)
            ranges = _conditional_ranges(archive.read(parts[sheet.name]))
            candidates = _candidate_cells(ranges, bounds)
            conditional_count = len(candidates)
            ws = session.workbook()[sheet.name]
            colors = _static_colors(ws, bounds, candidates, include_default)
            fallback = False
        except Exception as exc:
            reason = str(exc)
        finally:
            static_seconds = perf_counter() - static_started
        rendered_started = perf_counter()
        if fallback:
            logger.debug(
                "Color extraction legacy fallback sheet=%s reason=%s",
                sheet.name,
                reason,
            )
            return cells._extract_sheet_colors_com(
                sheet, include_default, ignore_colors
            )
        strict_token = cells._strict_display_format.set(True)
        try:
            for row, col in sorted(candidates):
                colors[row, col] = cells._resolve_cell_background_com(
                    sheet, row, col, include_default
                )
        except Exception as exc:
            fallback = True
            reason = f"rendered candidate extraction failed: {exc}"
        finally:
            cells._strict_display_format.reset(strict_token)
        if fallback:
            logger.debug(
                "Color extraction legacy fallback sheet=%s reason=%s",
                sheet.name,
                reason,
            )
            return cells._extract_sheet_colors_com(
                sheet, include_default, ignore_colors
            )
        rendered_seconds = perf_counter() - rendered_started
        ignore = cells._normalize_ignore_colors(ignore_colors)
        mapping: dict[str, list[tuple[int, int]]] = {}
        for (row, col), color in sorted(colors.items()):
            if color is not None:
                key = cells._normalize_color_key(color)
                if not cells._should_ignore_color(key, ignore):
                    mapping.setdefault(key, []).append((row, col - 1))
        return SheetColorsMap(sheet_name=sheet.name, colors_map=mapping)
    finally:
        if fallback and "rendered_started" in locals():
            rendered_seconds = perf_counter() - rendered_started
        logger.debug(
            "Color extraction sheet=%s used_cells=%d conditional_candidates=%d "
            "rendered_candidates=%d display_format_calls=%d fallback=%s "
            "static_duration_ms=%.3f com_duration_ms=%.3f duration_ms=%.3f",
            sheet.name,
            used_cells,
            conditional_count,
            len(candidates),
            cells._display_format_calls.get(),
            fallback,
            static_seconds * 1000,
            rendered_seconds * 1000,
            (perf_counter() - started) * 1000,
        )
        cells._display_format_calls.reset(token)


def extract_colors(
    workbook: xw.Book,
    include_default: bool,
    ignore_colors: set[str] | None,
    session: OpenpyxlExtractionSession | None,
) -> WorkbookColorsMap:
    """Own standalone resources and preserve legacy preparation and fallback."""
    cells._prepare_workbook_for_display_format(workbook)
    with ExitStack() as stack:
        archive = None
        parts: dict[str, str] = {}
        reason = "saved workbook unavailable"
        try:
            path = Path(workbook.fullname)
            if path.suffix.lower() not in {".xlsx", ".xlsm"}:
                raise ValueError("unsupported saved workbook format")
            if not workbook.api.Saved:
                raise ValueError("workbook has unsaved changes")
            if session is not None and session.file_path.resolve() != path.resolve():
                raise ValueError("saved workbook session path mismatch")
            archive = stack.enter_context(ZipFile(path))
            parts = _worksheet_parts(archive)
            if session is None:
                session = stack.enter_context(OpenpyxlExtractionSession(path))
        except Exception as exc:
            reason = str(exc)
            archive = None
        sheets = {}
        for sheet in workbook.sheets:
            cells._prepare_sheet_for_display_format(sheet)
            sheets[sheet.name] = _extract_sheet(
                sheet,
                session,
                archive,
                parts,
                reason,
                include_default,
                ignore_colors,
            )
        return WorkbookColorsMap(sheets=sheets)
