"""COM backend for Excel workbook extraction via xlwings."""

from __future__ import annotations

from dataclasses import dataclass
import logging
from typing import Any, cast

import xlwings as xw

from ...models import PrintArea
from ..cells import WorkbookColorsMap, extract_sheet_colors_map_com
from ..ranges import parse_range_zero_based
from .base import MergedCellData, PrintAreaData

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ComBackend:
    """COM-based backend for extraction tasks.

    Attributes:
        workbook: xlwings workbook instance.
    """

    workbook: xw.Book

    def extract_print_areas(self) -> PrintAreaData:
        """Extract print areas per sheet via xlwings/COM.

        Returns:
            Mapping of sheet name to print area list.
        """
        areas: PrintAreaData = {}
        for sheet in self.workbook.sheets:
            raw = ""
            try:
                raw = sheet.api.PageSetup.PrintArea or ""
            except Exception as exc:
                logger.warning(
                    "Failed to read print area via COM for sheet '%s'. (%r)",
                    sheet.name,
                    exc,
                )
            if not raw:
                continue
            for part in str(raw).split(","):
                parsed = _parse_print_area_range(part)
                if not parsed:
                    continue
                r1, c1, r2, c2 = parsed
                areas.setdefault(sheet.name, []).append(
                    PrintArea(r1=r1 + 1, c1=c1, r2=r2 + 1, c2=c2)
                )
        return areas

    def extract_colors_map(
        self, *, include_default_background: bool, ignore_colors: set[str] | None
    ) -> WorkbookColorsMap | None:
        """Extract colors_map via COM; logs and skips on failure.

        Args:
            include_default_background: Whether to include default backgrounds.
            ignore_colors: Optional set of color keys to ignore.

        Returns:
            WorkbookColorsMap or None when extraction fails.
        """
        try:
            return extract_sheet_colors_map_com(
                self.workbook,
                include_default_background=include_default_background,
                ignore_colors=ignore_colors,
            )
        except Exception as exc:
            logger.warning(
                "COM color map extraction failed; falling back to openpyxl. (%r)",
                exc,
            )
            return None

    def extract_auto_page_breaks(self) -> PrintAreaData:
        """Compute auto page-break rectangles per sheet using Excel COM.

        Returns:
            Mapping of sheet name to auto page-break areas.
        """
        results: PrintAreaData = {}
        for sheet in self.workbook.sheets:
            ws_api: Any | None = None
            original_display: bool | None = None
            failed = False
            try:
                ws_api = cast(Any, sheet.api)
                original_display = ws_api.DisplayPageBreaks
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
                            results.setdefault(sheet.name, []).append(
                                PrintArea(r1=r1, c1=c1 - 1, r2=r2, c2=c2 - 1)
                            )
            except Exception as exc:
                logger.warning(
                    "Failed to extract auto page breaks via COM for sheet '%s'. (%r)",
                    sheet.name,
                    exc,
                )
                failed = True
            finally:
                if ws_api is not None and original_display is not None:
                    try:
                        ws_api.DisplayPageBreaks = original_display
                    except Exception as exc:
                        logger.debug(
                            "Failed to restore DisplayPageBreaks for sheet '%s'. (%r)",
                            sheet.name,
                            exc,
                        )
            if failed:
                continue
        return results

    def extract_merged_cells(self) -> MergedCellData:
        """Extract merged cell ranges via COM (not implemented)."""
        raise NotImplementedError("COM merged cell extraction is not implemented.")


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


def _normalize_area_for_sheet(part: str, ws_name: str) -> str | None:
    """Strip sheet name from a range part when it matches the target sheet.

    Args:
        part: Raw range string part.
        ws_name: Target worksheet name.

    Returns:
        Range without sheet prefix, or None if not matching.
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
    """Split a CSV-like string while keeping commas inside single quotes intact.

    Args:
        raw: Raw CSV-like string.

    Returns:
        List of split parts.
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
