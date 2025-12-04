from __future__ import annotations

import logging
from collections import deque
from decimal import Decimal, InvalidOperation
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import xlwings as xw
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

from ..models import CellRow

logger = logging.getLogger(__name__)
_warned_keys: set[str] = set()


def warn_once(key: str, message: str) -> None:
    if key not in _warned_keys:
        logger.warning(message)
        _warned_keys.add(key)


def extract_sheet_cells(file_path: Path) -> Dict[str, List[CellRow]]:
    """Read all sheets via pandas and convert to CellRow list while skipping empty cells."""
    dfs = pd.read_excel(file_path, header=None, sheet_name=None, dtype=str)
    result: Dict[str, List[CellRow]] = {}
    for sheet_name, df in dfs.items():
        df = df.fillna("")
        rows: List[CellRow] = []
        for excel_row, row in enumerate(df.itertuples(index=False, name=None), start=1):
            filtered: Dict[str, int | float | str] = {}
            for j, v in enumerate(row):
                s = "" if v is None else str(v)
                if s.strip() == "":
                    continue
                filtered[str(j)] = _coerce_numeric_preserve_format(s)
            if not filtered:
                continue
            rows.append(CellRow(r=excel_row, c=filtered))
        result[sheet_name] = rows
    return result


def shrink_to_content(
    sheet: xw.Sheet,
    top: int,
    left: int,
    bottom: int,
    right: int,
    require_inside_border: bool = False,
    min_nonempty_ratio: float = 0.0,
) -> Tuple[int, int, int, int]:
    """Trim a rectangle based on cell contents and optional border heuristics."""
    rng = sheet.range((top, left), (bottom, right))
    vals = rng.value
    if vals is None:
        vals = []
    if not isinstance(vals, list):
        vals = [[vals]]
    elif vals and not isinstance(vals[0], list):
        vals = [vals]
    rows_n = len(vals)
    cols_n = len(vals[0]) if rows_n else 0

    def to_str(x):
        return "" if x is None else str(x)

    def is_empty_value(x):
        return to_str(x).strip() == ""

    def row_empty(i: int) -> bool:
        return cols_n == 0 or all(is_empty_value(vals[i][j]) for j in range(cols_n))

    def col_empty(j: int) -> bool:
        return rows_n == 0 or all(is_empty_value(vals[i][j]) for i in range(rows_n))

    def row_nonempty_ratio(i: int) -> float:
        if cols_n == 0:
            return 0.0
        cnt = sum(1 for j in range(cols_n) if not is_empty_value(vals[i][j]))
        return cnt / cols_n

    def col_nonempty_ratio(j: int) -> float:
        if rows_n == 0:
            return 0.0
        cnt = sum(1 for i in range(rows_n) if not is_empty_value(vals[i][j]))
        return cnt / rows_n

    XL_LINESTYLE_NONE = -4142
    XL_INSIDE_VERTICAL = 11
    XL_INSIDE_HORIZONTAL = 12

    def column_has_inside_border(col_idx: int) -> bool:
        if not require_inside_border:
            return False
        try:
            for r in range(top, bottom + 1):
                ls = (
                    sheet.api.Cells(r, left + col_idx)
                    .Borders(XL_INSIDE_VERTICAL)
                    .LineStyle
                )
                if ls is not None and ls != XL_LINESTYLE_NONE:
                    return True
        except Exception:
            pass
        return False

    def row_has_inside_border(row_idx: int) -> bool:
        if not require_inside_border:
            return False
        try:
            for c in range(left, right + 1):
                ls = (
                    sheet.api.Cells(top + row_idx, c)
                    .Borders(XL_INSIDE_HORIZONTAL)
                    .LineStyle
                )
                if ls is not None and ls != XL_LINESTYLE_NONE:
                    return True
        except Exception:
            pass
        return False

    def should_trim_col(j: int) -> bool:
        if col_empty(j):
            return True
        if require_inside_border and not column_has_inside_border(j):
            return True
        if min_nonempty_ratio > 0.0 and col_nonempty_ratio(j) < min_nonempty_ratio:
            return True
        return False

    def should_trim_row(i: int) -> bool:
        if row_empty(i):
            return True
        if require_inside_border and not row_has_inside_border(i):
            return True
        if min_nonempty_ratio > 0.0 and row_nonempty_ratio(i) < min_nonempty_ratio:
            return True
        return False

    while left <= right and cols_n > 0:
        if should_trim_col(0):
            for i in range(rows_n):
                if cols_n > 0:
                    vals[i].pop(0)
            cols_n = len(vals[0]) if rows_n else 0
            left += 1
        else:
            break
    while top <= bottom and rows_n > 0:
        if should_trim_row(0):
            vals.pop(0)
            rows_n = len(vals)
            top += 1
        else:
            break
    while left <= right and cols_n > 0:
        if should_trim_col(cols_n - 1):
            for i in range(rows_n):
                if cols_n > 0:
                    vals[i].pop(cols_n - 1)
            cols_n = len(vals[0]) if rows_n else 0
            right -= 1
        else:
            break
    while top <= bottom and rows_n > 0:
        if should_trim_row(rows_n - 1):
            vals.pop(rows_n - 1)
            rows_n = len(vals)
            bottom -= 1
        else:
            break
    return top, left, bottom, right


def load_border_maps_xlsx(xlsx_path: Path, sheet_name: str):
    wb = load_workbook(xlsx_path, data_only=True, read_only=False)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise KeyError(f"Sheet '{sheet_name}' not found in {xlsx_path}")

    ws = wb[sheet_name]
    try:
        min_col, min_row, max_col, max_row = range_boundaries(ws.calculate_dimension())
    except Exception:
        min_col, min_row, max_col, max_row = 1, 1, ws.max_column or 1, ws.max_row or 1

    shape = (max_row + 1, max_col + 1)
    has_border = np.zeros(shape, dtype=bool)
    top_edge = np.zeros(shape, dtype=bool)
    bottom_edge = np.zeros(shape, dtype=bool)
    left_edge = np.zeros(shape, dtype=bool)
    right_edge = np.zeros(shape, dtype=bool)

    def edge_has_style(edge) -> bool:
        if edge is None:
            return False
        style = getattr(edge, "style", None)
        return style is not None and style != "none"

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            b = getattr(cell, "border", None)
            if b is None:
                continue

            t = edge_has_style(b.top)
            btm = edge_has_style(b.bottom)
            l = edge_has_style(b.left)
            rgt = edge_has_style(b.right)

            if t or btm or l or rgt:
                has_border[r, c] = True
                if t:
                    top_edge[r, c] = True
                if btm:
                    bottom_edge[r, c] = True
                if l:
                    left_edge[r, c] = True
                if rgt:
                    right_edge[r, c] = True

    wb.close()
    return has_border, top_edge, bottom_edge, left_edge, right_edge, max_row, max_col


def detect_border_clusters(has_border: np.ndarray, min_size: int = 4):
    try:
        from scipy.ndimage import label

        structure = np.array([[0, 1, 0], [1, 1, 1], [0, 1, 0]], dtype=np.uint8)
        lbl, num = label(has_border.astype(np.uint8), structure=structure)  # type: ignore
        rects: List[Tuple[int, int, int, int]] = []
        for k in range(1, num + 1):
            ys, xs = np.where(lbl == k)
            if len(ys) < min_size:
                continue
            rects.append((int(ys.min()), int(xs.min()), int(ys.max()), int(xs.max())))
        return rects
    except Exception:
        warn_once(
            "scipy-missing",
            "scipy is not available. Falling back to pure-Python BFS for connected components, which may be significantly slower.",
        )
        h, w = has_border.shape
        visited = np.zeros_like(has_border, dtype=bool)
        rects: List[Tuple[int, int, int, int]] = []
        for r in range(h):
            for c in range(w):
                if not has_border[r, c] or visited[r, c]:
                    continue
                q = deque([(r, c)])
                visited[r, c] = True
                ys = [r]
                xs = [c]
                while q:
                    yy, xx = q.popleft()
                    for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                        ny, nx = yy + dy, xx + dx
                        if (
                            0 <= ny < h
                            and 0 <= nx < w
                            and has_border[ny, nx]
                            and not visited[ny, nx]
                        ):
                            visited[ny, nx] = True
                            q.append((ny, nx))
                            ys.append(ny)
                            xs.append(nx)
                if len(ys) >= min_size:
                    rects.append((min(ys), min(xs), max(ys), max(xs)))
        return rects


def _get_values_block(ws, top, left, bottom, right):
    vals = []
    for row in ws.iter_rows(
        min_row=top, max_row=bottom, min_col=left, max_col=right, values_only=True
    ):
        vals.append(list(row))
    return vals


def _table_density_metrics(matrix) -> tuple[float, float]:
    """
    Given a 2D matrix (list of rows), return (density, coverage).
    density: nonempty / total cells.
    coverage: area of tight bounding box of nonempty cells divided by total area.
    """
    if not matrix:
        return 0.0, 0.0
    rows = len(matrix)
    cols = len(matrix[0]) if rows else 0
    if rows == 0 or cols == 0:
        return 0.0, 0.0

    nonempty_coords = []
    for i, row in enumerate(matrix):
        if not isinstance(row, list):
            row = [row]
        for j, v in enumerate(row):
            if not (v is None or str(v).strip() == ""):
                nonempty_coords.append((i, j))

    total = rows * cols
    if not nonempty_coords:
        return 0.0, 0.0

    nonempty = len(nonempty_coords)
    density = nonempty / total

    ys = [p[0] for p in nonempty_coords]
    xs = [p[1] for p in nonempty_coords]
    bbox_h = (max(ys) - min(ys) + 1)
    bbox_w = (max(xs) - min(xs) + 1)
    coverage = (bbox_h * bbox_w) / total if total > 0 else 0.0
    return density, coverage


def _is_plausible_table(matrix) -> bool:
    """
    Heuristic: require at least 2 rows and 2 cols with meaningful data.
    - At least 2 rows have 2 以上の非空セル
    - At least 2 columns have 2 以上の非空セル
    """
    if not matrix:
        return False
    if not isinstance(matrix[0], list):
        # normalize to 2D
        matrix = [matrix]

    rows = len(matrix)
    cols = max((len(r) if isinstance(r, list) else 1) for r in matrix) if rows else 0
    if rows < 2 or cols < 2:
        return False

    row_counts = []
    col_counts = [0] * cols
    for r in matrix:
        if not isinstance(r, list):
            r = [r]
        cnt = 0
        for j in range(cols):
            v = r[j] if j < len(r) else None
            if not (v is None or str(v).strip() == ""):
                cnt += 1
                col_counts[j] += 1
        row_counts.append(cnt)

    rows_with_two = sum(1 for c in row_counts if c >= 2)
    cols_with_two = sum(1 for c in col_counts if c >= 2)
    return rows_with_two >= 2 and cols_with_two >= 2


def shrink_to_content_openpyxl(
    ws,
    top: int,
    left: int,
    bottom: int,
    right: int,
    require_inside_border: bool,
    top_edge,
    bottom_edge,
    left_edge,
    right_edge,
    min_nonempty_ratio: float = 0.0,
) -> Tuple[int, int, int, int]:
    vals = _get_values_block(ws, top, left, bottom, right)
    rows_n = bottom - top + 1
    cols_n = right - left + 1

    def to_str(x):
        return "" if x is None else str(x)

    def is_empty_value(x):
        return to_str(x).strip() == ""

    def row_nonempty_ratio_local(i: int) -> float:
        if cols_n <= 0:
            return 0.0
        row = vals[i]
        cnt = sum(1 for v in row if not is_empty_value(v))
        return cnt / cols_n

    def col_nonempty_ratio_local(j: int) -> float:
        if rows_n <= 0:
            return 0.0
        cnt = 0
        for i in range(rows_n):
            if not is_empty_value(vals[i][j]):
                cnt += 1
        return cnt / rows_n

    def col_has_inside_border(j_abs: int) -> bool:
        if not require_inside_border:
            return False
        count_pairs = 0
        for r_abs in range(top, bottom + 1):
            if (
                j_abs > left
                and right_edge[r_abs, j_abs - 1]
                and left_edge[r_abs, j_abs]
            ):
                count_pairs += 1
        return count_pairs > 0

    def row_has_inside_border(i_abs: int) -> bool:
        if not require_inside_border:
            return False
        count_pairs = 0
        for c_abs in range(left, right + 1):
            if i_abs > top and bottom_edge[i_abs - 1, c_abs] and top_edge[i_abs, c_abs]:
                count_pairs += 1
        return count_pairs > 0

    while left <= right and cols_n > 0:
        empty_col = all(
            not (
                top_edge[i, left]
                or bottom_edge[i, left]
                or left_edge[i, left]
                or right_edge[i, left]
            )
            for i in range(top, bottom + 1)
        )
        if (
            empty_col
            or (require_inside_border and not col_has_inside_border(left))
            or (
                min_nonempty_ratio > 0.0
                and col_nonempty_ratio_local(0) < min_nonempty_ratio
            )
        ):
            for i in range(rows_n):
                if cols_n > 0:
                    vals[i].pop(0)
            cols_n -= 1
            left += 1
        else:
            break
    while top <= bottom and rows_n > 0:
        empty_row = all(
            not (
                top_edge[top, j]
                or bottom_edge[top, j]
                or left_edge[top, j]
                or right_edge[top, j]
            )
            for j in range(left, right + 1)
        )
        if (
            empty_row
            or (require_inside_border and not row_has_inside_border(top))
            or (
                min_nonempty_ratio > 0.0
                and row_nonempty_ratio_local(0) < min_nonempty_ratio
            )
        ):
            vals.pop(0)
            rows_n -= 1
            top += 1
        else:
            break
    while left <= right and cols_n > 0:
        empty_col = all(
            not (
                top_edge[i, right]
                or bottom_edge[i, right]
                or left_edge[i, right]
                or right_edge[i, right]
            )
            for i in range(top, bottom + 1)
        )
        if (
            empty_col
            or (require_inside_border and not col_has_inside_border(right))
            or (
                min_nonempty_ratio > 0.0
                and col_nonempty_ratio_local(cols_n - 1) < min_nonempty_ratio
            )
        ):
            for i in range(rows_n):
                if cols_n > 0:
                    vals[i].pop(cols_n - 1)
            cols_n -= 1
            right -= 1
        else:
            break
    while top <= bottom and rows_n > 0:
        empty_row = all(
            not (
                top_edge[bottom, j]
                or bottom_edge[bottom, j]
                or left_edge[bottom, j]
                or right_edge[bottom, j]
            )
            for j in range(left, right + 1)
        )
        if (
            empty_row
            or (require_inside_border and not row_has_inside_border(bottom))
            or (
                min_nonempty_ratio > 0.0
                and row_nonempty_ratio_local(rows_n - 1) < min_nonempty_ratio
            )
        ):
            vals.pop(rows_n - 1)
            rows_n -= 1
            bottom -= 1
        else:
            break
    return top, left, bottom, right


def detect_tables_xlwings(sheet: xw.Sheet) -> List[str]:
    """Detect table-like ranges via COM: ListObjects first, then border clusters."""
    tables: List[str] = []
    try:
        for lo in sheet.api.ListObjects:
            rng = lo.Range
            top_row = int(rng.Row)
            left_col = int(rng.Column)
            bottom_row = top_row + int(rng.Rows.Count) - 1
            right_col = left_col + int(rng.Columns.Count) - 1
            addr = rng.Address(RowAbsolute=False, ColumnAbsolute=False)
            tables.append(addr)
    except Exception:
        pass

    used = sheet.used_range
    max_row = used.last_cell.row
    max_col = used.last_cell.column
    XL_EDGE_LEFT = 7
    XL_EDGE_TOP = 8
    XL_EDGE_BOTTOM = 9
    XL_EDGE_RIGHT = 10
    XL_INSIDE_VERTICAL = 11
    XL_INSIDE_HORIZONTAL = 12
    XL_LINESTYLE_NONE = -4142

    def cell_has_any_border(r: int, c: int) -> bool:
        try:
            b = sheet.api.Cells(r, c).Borders
            for idx in (
                XL_EDGE_LEFT,
                XL_EDGE_TOP,
                XL_EDGE_RIGHT,
                XL_EDGE_BOTTOM,
                XL_INSIDE_VERTICAL,
                XL_INSIDE_HORIZONTAL,
            ):
                ls = b(idx).LineStyle
                if ls is not None and ls != XL_LINESTYLE_NONE:
                    try:
                        if getattr(b(idx), "Weight", 0) == 0:
                            continue
                    except Exception:
                        pass
                    return True
            return False
        except Exception:
            return False

    grid = [[False] * (max_col + 1) for _ in range(max_row + 1)]
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if cell_has_any_border(r, c):
                grid[r][c] = True
    visited = [[False] * (max_col + 1) for _ in range(max_row + 1)]

    def dfs(sr: int, sc: int, acc: List[Tuple[int, int]]):
        stack = [(sr, sc)]
        while stack:
            rr, cc = stack.pop()
            if not (1 <= rr <= max_row and 1 <= cc <= max_col):
                continue
            if visited[rr][cc] or not grid[rr][cc]:
                continue
            visited[rr][cc] = True
            acc.append((rr, cc))
            for dr, dc in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                stack.append((rr + dr, cc + dc))

    clusters: List[Tuple[int, int, int, int]] = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if grid[r][c] and not visited[r][c]:
                cluster: List[Tuple[int, int]] = []
                dfs(r, c, cluster)
                if len(cluster) < 4:
                    continue
                rows = [rc[0] for rc in cluster]
                cols = [rc[1] for rc in cluster]
                top_row = min(rows)
                bottom_row = max(rows)
                left_col = min(cols)
                right_col = max(cols)
                clusters.append((top_row, left_col, bottom_row, right_col))

    def overlaps_for_merge(a, b):
        # Do not merge if one rect fully contains the other (separate clusters like big frame vs small table)
        contains = (
            (a[0] <= b[0] and a[1] <= b[1] and a[2] >= b[2] and a[3] >= b[3])
            or (b[0] <= a[0] and b[1] <= a[1] and b[2] >= a[2] and b[3] >= a[3])
        )
        if contains:
            return False
        return not (a[1] > b[3] or a[3] < b[1] or a[0] > b[2] or a[2] < b[0])

    merged_rects: List[Tuple[int, int, int, int]] = []
    for rect in sorted(clusters):
        merged = False
        for i, ex in enumerate(merged_rects):
            if overlaps_for_merge(rect, ex):
                merged_rects[i] = (
                    min(rect[0], ex[0]),
                    min(rect[1], ex[1]),
                    max(rect[2], ex[2]),
                    max(rect[3], ex[3]),
                )
                merged = True
                break
        if not merged:
            merged_rects.append(rect)

    for top_row, left_col, bottom_row, right_col in merged_rects:
        top_row, left_col, bottom_row, right_col = shrink_to_content(
            sheet, top_row, left_col, bottom_row, right_col, require_inside_border=False
        )
        try:
            rng_vals = sheet.range((top_row, left_col), (bottom_row, right_col)).value
            if not isinstance(rng_vals, list):
                rng_vals = [[rng_vals]]
            nonempty = sum(
                1
                for row in rng_vals
                for v in (row if isinstance(row, list) else [row])
                if not (v is None or str(v).strip() == "")
            )
        except Exception:
            nonempty = 0
        if nonempty < 3:
            continue
        density, coverage = _table_density_metrics(rng_vals)
        if density < 0.05 and coverage < 0.2:
            continue
        if not _is_plausible_table(rng_vals):
            continue
        addr = f"{xw.utils.col_name(left_col)}{top_row}:{xw.utils.col_name(right_col)}{bottom_row}"
        tables.append(addr)
    return tables


def detect_tables_openpyxl(xlsx_path: Path, sheet_name: str) -> List[str]:
    wb = load_workbook(
        xlsx_path,
        data_only=True,
        read_only=False,
    )
    ws = wb[sheet_name]
    tables: List[str] = []
    try:
        openpyxl_tables = []
        if hasattr(ws, "tables") and ws.tables:
            if isinstance(ws.tables, dict):
                openpyxl_tables = list(ws.tables.values())
            else:
                openpyxl_tables = list(ws.tables)
        elif hasattr(ws, "_tables") and ws._tables:  # type: ignore
            openpyxl_tables = list(ws._tables)  # type: ignore
        for t in openpyxl_tables:
            addr = t.ref
            tables.append(addr)
    except Exception:
        pass

    has_border, top_edge, bottom_edge, left_edge, right_edge, max_row, max_col = (
        load_border_maps_xlsx(xlsx_path, sheet_name)
    )
    rects = detect_border_clusters(has_border, min_size=4)

    def overlaps_for_merge(a, b):
        contains = (
            (a[0] <= b[0] and a[1] <= b[1] and a[2] >= b[2] and a[3] >= b[3])
            or (b[0] <= a[0] and b[1] <= a[1] and b[2] >= a[2] and b[3] >= a[3])
        )
        if contains:
            return False
        return not (a[1] > b[3] or a[3] < b[1] or a[0] > b[2] or a[2] < b[0])

    merged_rects: List[Tuple[int, int, int, int]] = []
    for rect in sorted(rects):
        merged = False
        for i, ex in enumerate(merged_rects):
            if overlaps_for_merge(rect, ex):
                merged_rects[i] = (
                    min(rect[0], ex[0]),
                    min(rect[1], ex[1]),
                    max(rect[2], ex[2]),
                    max(rect[3], ex[3]),
                )
                merged = True
                break
        if not merged:
            merged_rects.append(rect)

    for top_row, left_col, bottom_row, right_col in merged_rects:
        top_row, left_col, bottom_row, right_col = shrink_to_content_openpyxl(
            ws,
            top_row,
            left_col,
            bottom_row,
            right_col,
            require_inside_border=False,
            top_edge=top_edge,
            bottom_edge=bottom_edge,
            left_edge=left_edge,
            right_edge=right_edge,
            min_nonempty_ratio=0.0,
        )
        vals_block = _get_values_block(ws, top_row, left_col, bottom_row, right_col)
        nonempty = sum(
            1 for row in vals_block for v in row if not (v is None or str(v).strip() == "")
        )
        if nonempty < 3:
            continue
        density, coverage = _table_density_metrics(vals_block)
        if density < 0.05 and coverage < 0.2:
            continue
        if not _is_plausible_table(vals_block):
            continue
        addr = f"{get_column_letter(left_col)}{top_row}:{get_column_letter(right_col)}{bottom_row}"
        tables.append(addr)
    wb.close()
    return tables


def detect_tables(sheet: xw.Sheet) -> List[str]:
    excel_path: Optional[Path] = None
    try:
        excel_path = Path(sheet.book.fullname)
    except Exception:
        excel_path = None

    if excel_path and excel_path.suffix.lower() == ".xls":
        warn_once(
            f"xls-fallback::{excel_path}",
            f"File '{excel_path.name}' is .xls (BIFF); openpyxl cannot read it. Falling back to COM-based detection (slower). Consider converting to .xlsx.",
        )
        return detect_tables_xlwings(sheet)

    if excel_path and excel_path.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            import openpyxl  # noqa: F401
        except Exception:
            warn_once(
                "openpyxl-missing",
                "openpyxl is not installed. Falling back to COM-based detection (slower).",
            )
            return detect_tables_xlwings(sheet)

        try:
            return detect_tables_openpyxl(excel_path, sheet.name)
        except Exception as e:
            warn_once(
                f"openpyxl-parse-fallback::{excel_path}::{sheet.name}",
                f"openpyxl failed to parse '{excel_path.name}' (sheet '{sheet.name}'): {e!r}. Falling back to COM-based detection (slower).",
            )
            return detect_tables_xlwings(sheet)

    warn_once(
        "unknown-ext-fallback",
        "Workbook path or extension is unavailable; falling back to COM-based detection (slower).",
    )
    return detect_tables_xlwings(sheet)


_INT_RE = re.compile(r"^[+-]?\d+$")
_FLOAT_RE = re.compile(r"^[+-]?\d*\.\d+$")


def _coerce_numeric_preserve_format(val: str) -> int | float | str:
    """
    Convert numeric-looking strings to int/float while keeping precision.
    Integers stay int; decimals keep scale via Decimal before casting to float.
    """
    if _INT_RE.match(val):
        try:
            return int(val)
        except Exception:
            return val
    if _FLOAT_RE.match(val):
        try:
            dec = Decimal(val)
            scale = max(1, -dec.as_tuple().exponent)
            quantized = dec.quantize(Decimal("1." + "0" * scale))
            return float(quantized)
        except (InvalidOperation, Exception):
            return val
    return val
