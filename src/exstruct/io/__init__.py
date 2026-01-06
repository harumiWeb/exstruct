from __future__ import annotations

import logging
from pathlib import Path
import re
from typing import Literal, cast

from ..core.ranges import RangeBounds, parse_range_zero_based
from ..errors import OutputError, SerializationError
from ..models import (
    Arrow,
    CellRow,
    Chart,
    PrintArea,
    PrintAreaView,
    Shape,
    SmartArt,
    WorkbookData,
)
from ..models.types import JsonStructure
from .serialize import (
    _FORMAT_HINTS,
    _ensure_format_hint,
    _require_toon,
    _require_yaml,
    _serialize_payload_from_hint,
)

logger = logging.getLogger(__name__)


def dict_without_empty_values(obj: object) -> JsonStructure:
    """
    Remove None, empty string, empty list, and empty dict values from a nested structure or supported model object.

    Recursively processes dicts, lists, and supported model types (WorkbookData, CellRow, Chart, PrintArea, PrintAreaView, Shape, Arrow, SmartArt). Model instances are converted to dictionaries with None fields excluded before recursive cleaning. Values considered empty and removed are: `None`, `""` (empty string), `[]` (empty list), and `{}` (empty dict).

    Parameters:
        obj (object): A value to clean; may be a dict, list, scalar, or one of the supported model instances.

    Returns:
        JsonStructure: The input structure with empty values removed, preserving other values and nesting.
    """
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
    if isinstance(
        obj,
        WorkbookData
        | CellRow
        | Chart
        | PrintArea
        | PrintAreaView
        | Shape
        | Arrow
        | SmartArt,
    ):
        return dict_without_empty_values(
            obj.model_dump(exclude_none=True, by_alias=True)
        )
    return cast(JsonStructure, obj)


def _write_text(path: Path, text: str) -> None:
    """Write UTF-8 text to disk, wrapping IO errors."""
    try:
        path.write_text(text, encoding="utf-8")
    except Exception as exc:
        raise OutputError(f"Failed to write output to '{path}'.") from exc


def save_as_json(
    model: WorkbookData, path: Path, *, pretty: bool = False, indent: int | None = None
) -> None:
    text = serialize_workbook(model, fmt="json", pretty=pretty, indent=indent)
    _write_text(path, text)


def save_as_yaml(model: WorkbookData, path: Path) -> None:
    text = serialize_workbook(model, fmt="yaml")
    _write_text(path, text)


def save_as_toon(model: WorkbookData, path: Path) -> None:
    text = serialize_workbook(model, fmt="toon")
    _write_text(path, text)


def _sanitize_sheet_filename(name: str) -> str:
    """Make a sheet name safe for filesystem usage."""
    safe = re.sub(r"[\\/:*?\"<>|]", "_", name)
    return safe or "sheet"


def _parse_range_zero_based(range_str: str) -> RangeBounds | None:
    """Parse an Excel range string into zero-based bounds.

    Args:
        range_str: Excel range string (e.g., "Sheet1!A1:B2").

    Returns:
        RangeBounds in zero-based coordinates, or None on failure.
    """
    return parse_range_zero_based(range_str)


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
        r1 = bounds.r1 + 1
        r2 = bounds.r2 + 1
        if (
            r1 >= area.r1
            and r2 <= area.r2
            and bounds.c1 >= area.c1
            and bounds.c2 <= area.c2
        ):
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
    top = (area.r1 - 1) * row_px
    right = (area.c2 + 1) * col_px
    bottom = area.r2 * row_px
    return left, top, right, bottom


def _rects_overlap(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    """
    Determine whether two axis-aligned rectangles intersect (overlap in area).

    Parameters:
        a (tuple[int, int, int, int]): Rectangle A as (left, top, right, bottom).
        b (tuple[int, int, int, int]): Rectangle B as (left, top, right, bottom).

    Notes:
        Rectangles are treated as half-open in this context: if they only touch at edges or corners, they do not count as overlapping.

    Returns:
        bool: `True` if the rectangles have a non-zero-area intersection, `False` otherwise.
    """
    return not (a[2] <= b[0] or a[0] >= b[2] or a[3] <= b[1] or a[1] >= b[3])


def _filter_shapes_to_area(
    shapes: list[Shape | Arrow | SmartArt], area: PrintArea
) -> list[Shape | Arrow | SmartArt]:
    """
    Filter drawable shapes to those that intersect the given print area.

    Shapes and the print area are compared in approximate pixel coordinates. Shapes that have both width and height are included when their bounding rectangle overlaps the area. Shapes with unknown size (width or height is None) are treated as a point at their left/top coordinates and included only if that point lies inside the area.

    Parameters:
        shapes (list[Shape | Arrow | SmartArt]): Drawable objects with `l`, `t`, `w`, `h` coordinates.
        area (PrintArea): Cell-based print area that will be converted to an approximate pixel rectangle.

    Returns:
        list[Shape | Arrow | SmartArt]: Subset of `shapes` whose geometry intersects the print area.
    """
    area_rect = _area_to_px_rect(area)
    filtered: list[Shape | Arrow | SmartArt] = []
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


def _iter_area_views(
    workbook: WorkbookData,
    *,
    area_attr: Literal["print_areas", "auto_print_areas"],
    normalize: bool,
    include_shapes: bool,
    include_charts: bool,
    include_shape_size: bool,
    include_chart_size: bool,
) -> dict[str, list[PrintAreaView]]:
    views: dict[str, list[PrintAreaView]] = {}
    for sheet_name, sheet in workbook.sheets.items():
        areas: list[PrintArea] = getattr(sheet, area_attr)
        if not areas:
            continue
        sheet_views: list[PrintAreaView] = []
        for area in areas:
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
    return _iter_area_views(
        workbook,
        area_attr="print_areas",
        normalize=normalize,
        include_shapes=include_shapes,
        include_charts=include_charts,
        include_shape_size=include_shape_size,
        include_chart_size=include_chart_size,
    )


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
    format_hint = _ensure_format_hint(
        fmt,
        allowed=_FORMAT_HINTS,
        error_type=SerializationError,
        error_message="Unsupported print-area export format '{fmt}'. Allowed: json, yaml, yml, toon.",
    )

    views = build_print_area_views(
        workbook,
        normalize=normalize,
        include_shapes=include_shapes,
        include_charts=include_charts,
        include_shape_size=include_shape_size,
        include_chart_size=include_chart_size,
    )
    if not views:
        logger.info("No print areas found; skipping export to %s", output_dir)
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
            payload = dict_without_empty_values(
                view.model_dump(exclude_none=True, by_alias=True)
            )
            text = _serialize_payload_from_hint(
                payload, format_hint, pretty=pretty, indent=indent
            )
            _write_text(path, text)
            written[key] = path
    return written


def save_auto_page_break_views(
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
    Save auto page-break areas (computed via Excel COM) per sheet in the specified format.
    Returns a map of area key (e.g., 'Sheet1#auto#1') to written path.
    """
    format_hint = _ensure_format_hint(
        fmt,
        allowed=_FORMAT_HINTS,
        error_type=SerializationError,
        error_message="Unsupported auto page-break export format '{fmt}'. Allowed: json, yaml, yml, toon.",
    )

    views = _iter_area_views(
        workbook,
        area_attr="auto_print_areas",
        normalize=normalize,
        include_shapes=include_shapes,
        include_charts=include_charts,
        include_shape_size=include_shape_size,
        include_chart_size=include_chart_size,
    )
    if not views:
        logger.info("No auto page-break areas found; skipping export to %s", output_dir)
        return {}

    output_dir.mkdir(parents=True, exist_ok=True)
    written: dict[str, Path] = {}
    suffix = {"json": ".json", "yaml": ".yaml", "toon": ".toon"}[format_hint]

    for sheet_name, sheet_views in views.items():
        for idx, view in enumerate(sheet_views):
            key = f"{sheet_name}#auto#{idx + 1}"
            area = view.area
            file_name = (
                f"{_sanitize_sheet_filename(sheet_name)}"
                f"_auto_page{idx + 1}_r{area.r1}-{area.r2}_c{area.c1}-{area.c2}{suffix}"
            )
            path = output_dir / file_name
            payload = dict_without_empty_values(
                view.model_dump(exclude_none=True, by_alias=True)
            )
            text = _serialize_payload_from_hint(
                payload, format_hint, pretty=pretty, indent=indent
            )
            _write_text(path, text)
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
    format_hint = _ensure_format_hint(
        fmt,
        allowed=_FORMAT_HINTS,
        error_type=SerializationError,
        error_message="Unsupported export format '{fmt}'. Allowed: json, yaml, yml, toon.",
    )
    filtered_dict = dict_without_empty_values(
        model.model_dump(exclude_none=True, by_alias=True)
    )
    return _serialize_payload_from_hint(
        filtered_dict, format_hint, pretty=pretty, indent=indent
    )


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
                "sheet": sheet_data.model_dump(exclude_none=True, by_alias=True),
            }
        )
        file_name = f"{_sanitize_sheet_filename(sheet_name)}.json"
        path = output_dir / file_name
        text = _serialize_payload_from_hint(
            payload, "json", pretty=pretty, indent=indent
        )
        _write_text(path, text)
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
    format_hint = _ensure_format_hint(
        fmt,
        allowed=_FORMAT_HINTS,
        error_type=SerializationError,
        error_message="Unsupported sheet export format: {fmt}",
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    written: dict[str, Path] = {}
    for sheet_name, sheet_data in workbook.sheets.items():
        payload = dict_without_empty_values(
            {
                "book_name": workbook.book_name,
                "sheet_name": sheet_name,
                "sheet": sheet_data.model_dump(exclude_none=True, by_alias=True),
            }
        )
        suffix = {"json": ".json", "yaml": ".yaml", "toon": ".toon"}[format_hint]
        file_name = f"{_sanitize_sheet_filename(sheet_name)}{suffix}"
        path = output_dir / file_name
        text = _serialize_payload_from_hint(
            payload, format_hint, pretty=pretty, indent=indent
        )
        _write_text(path, text)
        written[sheet_name] = path
    return written


__all__ = [
    "dict_without_empty_values",
    "save_as_json",
    "save_as_yaml",
    "save_as_toon",
    "save_sheets",
    "save_sheets_as_json",
    "build_print_area_views",
    "save_print_area_views",
    "save_auto_page_break_views",
    "serialize_workbook",
    "_require_yaml",
    "_require_toon",
]
