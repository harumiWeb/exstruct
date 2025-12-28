from __future__ import annotations

import logging
from typing import Literal

import xlwings as xw

from ..models import Chart, ChartSeries
from ..models.maps import XL_CHART_TYPE_MAP

logger = logging.getLogger(__name__)


def _extract_series_args_text(formula: str) -> str | None:  # noqa: C901
    """Extract the outer argument text from '=SERIES(...)'; return None if unmatched."""
    if not formula:
        return None
    s = formula.strip()
    if not s.upper().startswith("=SERIES"):
        return None
    try:
        open_idx = s.index("(", s.upper().index("=SERIES"))
    except ValueError:
        return None
    depth_paren = 0
    depth_brace = 0
    in_str = False
    i = open_idx + 1
    start = i
    while i < len(s):
        ch = s[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(s) and s[i + 1] == '"':
                    i += 2
                    continue
                else:
                    in_str = False
                    i += 1
                    continue
            else:
                i += 1
                continue
        else:
            if ch == '"':
                in_str = True
                i += 1
                continue
            elif ch == "(":
                depth_paren += 1
            elif ch == ")":
                if depth_paren == 0:
                    return s[start:i].strip()
                depth_paren -= 1
            elif ch == "{":
                depth_brace += 1
            elif ch == "}":
                if depth_brace > 0:
                    depth_brace -= 1
            i += 1
    return None


def _split_top_level_args(args_text: str) -> list[str]:  # noqa: C901
    """Split SERIES arguments at top-level separators (',' or ';')."""
    if args_text is None:
        return []
    use_semicolon = (";" in args_text) and ("," not in args_text.split('"')[0])
    sep_chars = (";",) if use_semicolon else (",",)
    args: list[str] = []
    buf: list[str] = []
    depth_paren = 0
    depth_brace = 0
    in_str = False
    i = 0
    while i < len(args_text):
        ch = args_text[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(args_text) and args_text[i + 1] == '"':
                    buf.append('""')
                    i += 2
                    continue
                else:
                    in_str = False
                    buf.append('"')
                    i += 1
                    continue
            else:
                buf.append(ch)
                i += 1
                continue
        else:
            if ch == '"':
                in_str = True
                buf.append('"')
                i += 1
                continue
            elif ch == "(":
                depth_paren += 1
                buf.append(ch)
                i += 1
                continue
            elif ch == ")":
                depth_paren = max(0, depth_paren - 1)
                buf.append(ch)
                i += 1
                continue
            elif ch == "{":
                depth_brace += 1
                buf.append(ch)
                i += 1
                continue
            elif ch == "}":
                depth_brace = max(0, depth_brace - 1)
                buf.append(ch)
                i += 1
                continue
            elif (ch in sep_chars) and depth_paren == 0 and depth_brace == 0:
                args.append("".join(buf).strip())
                buf = []
                i += 1
                continue
            else:
                buf.append(ch)
                i += 1
                continue
    if buf or (args and args_text.endswith(sep_chars)):
        args.append("".join(buf).strip())
    return args


def _unquote_excel_string(s: str | None) -> str | None:
    """Decode Excel-style quoted string; return None if not quoted."""
    if s is None:
        return None
    st = s.strip()
    if len(st) >= 2 and st[0] == '"' and st[-1] == '"':
        inner = st[1:-1]
        return inner.replace('""', '"')
    return None


def parse_series_formula(formula: str) -> dict[str, str | None] | None:
    """Parse =SERIES into a dict of references; return None on failure."""
    args_text = _extract_series_args_text(formula)
    if args_text is None:
        return None
    parts = _split_top_level_args(args_text)
    name_part = parts[0].strip() if len(parts) >= 1 and parts[0].strip() != "" else None
    x_part = parts[1].strip() if len(parts) >= 2 and parts[1].strip() != "" else None
    y_part = parts[2].strip() if len(parts) >= 3 and parts[2].strip() != "" else None
    plot_order_part = (
        parts[3].strip() if len(parts) >= 4 and parts[3].strip() != "" else None
    )
    bubble_part = (
        parts[4].strip() if len(parts) >= 5 and parts[4].strip() != "" else None
    )
    name_literal = _unquote_excel_string(name_part)
    name_range = None if name_literal is not None else name_part
    return {
        "name_range": name_range,
        "x_range": x_part,
        "y_range": y_part,
        "plot_order": plot_order_part,
        "bubble_size_range": bubble_part,
        "name_literal": name_literal,
    }


def get_charts(
    sheet: xw.Sheet, mode: Literal["light", "standard", "verbose"] = "standard"
) -> list[Chart]:
    """Parse charts in a sheet into Chart models; failed charts carry an error field."""
    charts: list[Chart] = []
    for ch in sheet.charts:
        series_list: list[ChartSeries] = []
        y_axis_title: str = ""
        y_axis_range: list[int] = []
        chart_type_label: str = "unknown"
        error: str | None = None
        chart_width: int | None = None
        chart_height: int | None = None

        try:
            chart_com = sheet.api.ChartObjects(ch.name).Chart
            chart_type_num = chart_com.ChartType
            chart_type_label = XL_CHART_TYPE_MAP.get(
                chart_type_num, f"unknown_{chart_type_num}"
            )
            try:
                chart_width = int(ch.width)
                chart_height = int(ch.height)
            except Exception:
                chart_width = None
                chart_height = None

            for s in chart_com.SeriesCollection():
                parsed = parse_series_formula(getattr(s, "Formula", ""))
                name_range = parsed["name_range"] if parsed else None
                x_range = parsed["x_range"] if parsed else None
                y_range = parsed["y_range"] if parsed else None

                series_list.append(
                    ChartSeries(
                        name=s.Name,
                        name_range=name_range,
                        x_range=x_range,
                        y_range=y_range,
                    )
                )

            try:
                y_axis = chart_com.Axes(2, 1)
                if y_axis.HasTitle:
                    y_axis_title = y_axis.AxisTitle.Text
                y_axis_range = [y_axis.MinimumScale, y_axis.MaximumScale]
            except Exception:
                y_axis_title = ""
                y_axis_range = []

            title = chart_com.ChartTitle.Text if chart_com.HasTitle else None
        except Exception:
            logger.warning("Failed to parse chart; returning with error string.")
            title = None
            error = "Failed to build chart JSON structure"

        charts.append(
            Chart(
                name=ch.name,
                chart_type=chart_type_label,
                title=title,
                y_axis_title=y_axis_title,
                y_axis_range=[float(v) for v in y_axis_range],
                w=chart_width,
                h=chart_height,
                series=series_list,
                l=int(ch.left),
                t=int(ch.top),
                error=error,
            )
        )
    return charts
