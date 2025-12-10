from __future__ import annotations

import importlib
import json
from pathlib import Path
import re
from types import ModuleType
from typing import Literal, cast

from openpyxl.utils import range_boundaries

from ..models import CellRow, Chart, PrintArea, PrintAreaView, Shape, WorkbookData
from ..models.types import JsonStructure


def dict_without_empty_values(obj: object) -> JsonStructure:
    """Recursively drop empty values from nested structures."""
    if isinstance(obj, dict):
        return {
            k: dict_without_empty_values(v)
            for k, v in obj.items()
            if v not in [None, "", [], {}]
        }
    if isinstance(obj, list):
        return [
            dict_without_empty_values(v) for v in obj if v not in [None, "", [], {}]
        ]
    if isinstance(obj, (WorkbookData, CellRow, Chart, PrintArea, PrintAreaView, Shape)):
        return dict_without_empty_values(obj.model_dump(exclude_none=True))
    return cast(JsonStructure, obj)


def save_as_json(
    model: WorkbookData, path: Path, *, pretty: bool = False, indent: int | None = None
) -> None:
    text = serialize_workbook(model, fmt="json", pretty=pretty, indent=indent)
    path.write_text(text, encoding="utf-8")


def save_as_yaml(model: WorkbookData, path: Path) -> None:
    text = serialize_workbook(model, fmt="yaml")
    path.write_text(text, encoding="utf-8")


def save_as_toon(model: WorkbookData, path: Path) -> None:
    text = serialize_workbook(model, fmt="toon")
    path.write_text(text, encoding="utf-8")


def _sanitize_sheet_filename(name: str) -> str:
    """Make a sheet name safe for filesystem usage."""
    safe = re.sub(r"[\\/:*?\"<>|]", "_", name)
    return safe or "sheet"


def _parse_range_zero_based(range_str: str) -> tuple[int, int, int, int] | None:
    """
    Parse an Excel range string into zero-based (r1, c1, r2, c2) bounds.
    Returns None on failure.
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
    return (min_row - 1, min_col - 1, max_row - 1, max_col - 1)


def _row_in_area(row: CellRow, area: PrintArea) -> bool:
    return area.r1 <= row.r <= area.r2


def _filter_row_to_area(
    row: CellRow, area: PrintArea, *, normalize: bool = False
) -> CellRow | None:
    if not _row_in_area(row, area):
        return None

    filtered_cells: dict[str, int | float | str] = {}
    filtered_links: dict[str, str] = {}

    for col_idx_str, value in row.c.items():
        try:
            col_idx = int(col_idx_str)
        except Exception:
            continue
        if area.c1 <= col_idx <= area.c2:
            key = str(col_idx - area.c1) if normalize else col_idx_str
            filtered_cells[key] = value

    if row.links:
        for col_idx_str, url in row.links.items():
            try:
                col_idx = int(col_idx_str)
            except Exception:
                continue
            if area.c1 <= col_idx <= area.c2:
                key = str(col_idx - area.c1) if normalize else col_idx_str
                filtered_links[key] = url

    if not filtered_cells and not filtered_links:
        return None

    new_row_idx = row.r - area.r1 if normalize else row.r
    return CellRow(r=new_row_idx, c=filtered_cells, links=filtered_links or None)


def _filter_table_candidates_to_area(
    table_candidates: list[str], area: PrintArea
) -> list[str]:
    filtered: list[str] = []
    for candidate in table_candidates:
        bounds = _parse_range_zero_based(candidate)
        if not bounds:
            continue
        r1, c1, r2, c2 = bounds
        if r1 >= area.r1 and r2 <= area.r2 and c1 >= area.c1 and c2 <= area.c2:
            filtered.append(candidate)
    return filtered


def _area_to_px_rect(
    area: PrintArea, *, col_px: int = 64, row_px: int = 20
) -> tuple[int, int, int, int]:
    """
    Convert a cell-based print area to an approximate pixel rectangle (l, t, r, b).
    Uses default Excel-like cell sizes; accuracy is highest when shapes/charts are COM-extracted.
    """
    left = area.c1 * col_px
    top = area.r1 * row_px
    right = (area.c2 + 1) * col_px
    bottom = (area.r2 + 1) * row_px
    return left, top, right, bottom


def _rects_overlap(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    """Return True if rectangles (l, t, r, b) overlap."""
    return not (a[2] <= b[0] or a[0] >= b[2] or a[3] <= b[1] or a[1] >= b[3])


def _filter_shapes_to_area(shapes: list[Shape], area: PrintArea) -> list[Shape]:
    area_rect = _area_to_px_rect(area)
    filtered: list[Shape] = []
    for shp in shapes:
        if shp.w is None or shp.h is None:
            # Fallback: treat shape as a point if size is unknown (standard mode).
            if (
                area_rect[0] <= shp.l <= area_rect[2]
                and area_rect[1] <= shp.t <= area_rect[3]
            ):
                filtered.append(shp)
            continue
        shp_rect = (shp.l, shp.t, shp.l + shp.w, shp.t + shp.h)
        if _rects_overlap(area_rect, shp_rect):
            filtered.append(shp)
    return filtered


def _filter_charts_to_area(charts: list[Chart], area: PrintArea) -> list[Chart]:
    area_rect = _area_to_px_rect(area)
    filtered: list[Chart] = []
    for ch in charts:
        if ch.w is None or ch.h is None:
            continue
        ch_rect = (ch.l, ch.t, ch.l + ch.w, ch.t + ch.h)
        if _rects_overlap(area_rect, ch_rect):
            filtered.append(ch)
    return filtered


def build_print_area_views(
    workbook: WorkbookData,
    *,
    normalize: bool = False,
    include_shapes: bool = True,
    include_charts: bool = True,
    include_shape_size: bool = True,
    include_chart_size: bool = True,
) -> dict[str, list[PrintAreaView]]:
    """
    Construct PrintAreaView instances for all print areas in the workbook.
    Returns a mapping of sheet name to ordered list of PrintAreaView.
    """
    views: dict[str, list[PrintAreaView]] = {}
    for sheet_name, sheet in workbook.sheets.items():
        if not sheet.print_areas:
            continue
        sheet_views: list[PrintAreaView] = []
        for area in sheet.print_areas:
            rows_in_area: list[CellRow] = []
            for row in sheet.rows:
                filtered_row = _filter_row_to_area(row, area, normalize=normalize)
                if filtered_row:
                    rows_in_area.append(filtered_row)
            area_tables = _filter_table_candidates_to_area(sheet.table_candidates, area)
            area_shapes = (
                _filter_shapes_to_area(sheet.shapes, area) if include_shapes else []
            )
            if not include_shape_size:
                area_shapes = [
                    s.model_copy(update={"w": None, "h": None}) for s in area_shapes
                ]
            area_charts = (
                _filter_charts_to_area(sheet.charts, area) if include_charts else []
            )
            if not include_chart_size:
                area_charts = [
                    c.model_copy(update={"w": None, "h": None}) for c in area_charts
                ]
            sheet_views.append(
                PrintAreaView(
                    book_name=workbook.book_name,
                    sheet_name=sheet_name,
                    area=area,
                    shapes=area_shapes,
                    charts=area_charts,
                    rows=rows_in_area,
                    table_candidates=area_tables,
                )
            )
        if sheet_views:
            views[sheet_name] = sheet_views
    return views


def save_print_area_views(
    workbook: WorkbookData,
    output_dir: Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
    normalize: bool = False,
    include_shapes: bool = True,
    include_charts: bool = True,
    include_shape_size: bool = True,
    include_chart_size: bool = True,
) -> dict[str, Path]:
    """
    Save each print area as an individual file in the specified format.
    Returns a map of area key (e.g., 'Sheet1#1') to written path.
    """
    format_hint = fmt.lower()
    if format_hint == "yml":
        format_hint = "yaml"
    if format_hint not in ("json", "yaml", "toon"):
        raise ValueError(f"Unsupported print-area export format: {fmt}")

    views = build_print_area_views(
        workbook,
        normalize=normalize,
        include_shapes=include_shapes,
        include_charts=include_charts,
        include_shape_size=include_shape_size,
        include_chart_size=include_chart_size,
    )
    if not views:
        return {}

    output_dir.mkdir(parents=True, exist_ok=True)
    written: dict[str, Path] = {}
    suffix = {"json": ".json", "yaml": ".yaml", "toon": ".toon"}[format_hint]

    for sheet_name, sheet_views in views.items():
        for idx, view in enumerate(sheet_views):
            key = f"{sheet_name}#{idx + 1}"
            area = view.area
            file_name = (
                f"{_sanitize_sheet_filename(sheet_name)}"
                f"_area{idx + 1}_r{area.r1}-{area.r2}_c{area.c1}-{area.c2}{suffix}"
            )
            path = output_dir / file_name
            match format_hint:
                case "json":
                    indent_val = 2 if pretty and indent is None else indent
                    text = view.to_json(pretty=pretty, indent=indent_val)
                case "yaml":
                    text = view.to_yaml()
                case "toon":
                    text = view.to_toon()
                case _:
                    raise ValueError(f"Unsupported print-area export format: {fmt}")
            path.write_text(text, encoding="utf-8")
            written[key] = path
    return written


def serialize_workbook(
    model: WorkbookData,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> str:
    """
    Convert WorkbookData to string in the requested format without writing to disk.
    """
    format_hint = fmt.lower()
    if format_hint == "yml":
        format_hint = "yaml"
    filtered_dict = dict_without_empty_values(model.model_dump(exclude_none=True))

    match format_hint:
        case "json":
            indent_val = 2 if pretty and indent is None else indent
            return json.dumps(filtered_dict, ensure_ascii=False, indent=indent_val)
        case "yaml":
            yaml = _require_yaml()
            return str(
                yaml.safe_dump(
                    filtered_dict,
                    allow_unicode=True,
                    sort_keys=False,
                    indent=2,
                )
            )
        case "toon":
            toon = _require_toon()
            return str(toon.encode(filtered_dict))
        case _:
            raise ValueError(f"Unsupported export format: {fmt}")


def save_sheets_as_json(
    workbook: WorkbookData,
    output_dir: Path,
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> dict[str, Path]:
    """
    Save each sheet as an individual JSON file.
    Contents include book_name and the sheet's SheetData.
    Returns a map of sheet name -> written path.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    written: dict[str, Path] = {}
    for sheet_name, sheet_data in workbook.sheets.items():
        payload = dict_without_empty_values(
            {
                "book_name": workbook.book_name,
                "sheet_name": sheet_name,
                "sheet": sheet_data.model_dump(exclude_none=True),
            }
        )
        file_name = f"{_sanitize_sheet_filename(sheet_name)}.json"
        path = output_dir / file_name
        indent_val = 2 if pretty and indent is None else indent
        path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=indent_val), encoding="utf-8"
        )
        written[sheet_name] = path
    return written


def save_sheets(
    workbook: WorkbookData,
    output_dir: Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> dict[str, Path]:
    """
    Save each sheet as an individual file in the specified format (json/yaml/toon).
    Payload includes book_name and the sheet's SheetData.
    """
    format_hint = fmt.lower()
    if format_hint == "yml":
        format_hint = "yaml"
    if format_hint not in ("json", "yaml", "toon"):
        raise ValueError(f"Unsupported sheet export format: {fmt}")

    output_dir.mkdir(parents=True, exist_ok=True)
    written: dict[str, Path] = {}
    for sheet_name, sheet_data in workbook.sheets.items():
        payload = dict_without_empty_values(
            {
                "book_name": workbook.book_name,
                "sheet_name": sheet_name,
                "sheet": sheet_data.model_dump(exclude_none=True),
            }
        )
        suffix = {"json": ".json", "yaml": ".yaml", "toon": ".toon"}[format_hint]
        file_name = f"{_sanitize_sheet_filename(sheet_name)}{suffix}"
        path = output_dir / file_name
        match format_hint:
            case "json":
                indent_val = 2 if pretty and indent is None else indent
                text = json.dumps(payload, ensure_ascii=False, indent=indent_val)
            case "yaml":
                yaml = _require_yaml()
                text = str(
                    yaml.safe_dump(
                        payload, allow_unicode=True, sort_keys=False, indent=2
                    )
                )
            case "toon":
                toon = _require_toon()
                text = str(toon.encode(payload))
            case _:
                raise ValueError(f"Unsupported sheet export format: {format_hint}")
        path.write_text(text, encoding="utf-8")
        written[sheet_name] = path
    return written


def _require_yaml() -> ModuleType:
    try:
        module = importlib.import_module("yaml")
    except ImportError as e:
        raise RuntimeError(
            "YAML export requires pyyaml. Install it via `pip install pyyaml` "
            "or add the 'yaml' extra."
        ) from e
    return module


def _require_toon() -> ModuleType:
    try:
        module = importlib.import_module("toon")
    except ImportError as e:
        raise RuntimeError(
            "TOON export requires python-toon. Install it via `pip install python-toon` "
            "or add the 'toon' extra."
        ) from e
    return module


__all__ = [
    "dict_without_empty_values",
    "save_as_json",
    "save_as_yaml",
    "save_as_toon",
    "save_sheets",
    "save_sheets_as_json",
    "build_print_area_views",
    "save_print_area_views",
    "serialize_workbook",
]
