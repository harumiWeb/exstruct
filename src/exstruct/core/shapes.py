from __future__ import annotations

import math
from typing import Dict, List, Optional

import xlwings as xw
from xlwings import Book

from ..models import Shape
from ..models.maps import MSO_AUTO_SHAPE_TYPE_MAP, MSO_SHAPE_TYPE_MAP


def compute_line_angle_deg(w: float, h: float) -> float:
    """Compute clockwise angle in Excel coordinates where 0 deg points East."""
    return math.degrees(math.atan2(h, w)) % 360.0


def angle_to_compass(angle: float) -> str:
    """Convert angle to 8-point compass direction (0deg=E, 45deg=NE, 90deg=N, etc)."""
    dirs = ["E", "NE", "N", "NW", "W", "SW", "S", "SE"]
    idx = int(((angle + 22.5) % 360) // 45)
    return dirs[idx]


def coord_to_cell_by_edges(
    row_edges: List[float], col_edges: List[float], x: float, y: float
) -> Optional[str]:
    """Estimate cell address from coordinates and cumulative edges; return None if out of range."""

    def find_index(edges, pos):
        for i in range(1, len(edges)):
            if edges[i - 1] <= pos < edges[i]:
                return i
        return None

    r = find_index(row_edges, y)
    c = find_index(col_edges, x)
    if r is None or c is None:
        return None
    return f"{xw.utils.col_name(c)}{r}"


def has_arrow(style_val) -> bool:
    """Return True if Excel arrow style value indicates an arrowhead."""
    try:
        v = int(style_val)
        return v != 0
    except Exception:
        return False


def iter_shapes_recursive(shp):
    """Yield shapes recursively, including group children."""
    yield shp
    try:
        if shp.api.Type == 6:
            items = shp.api.GroupItems
            for i in range(1, items.Count + 1):
                inner = items.Item(i)
                try:
                    name = inner.Name
                    xl_shape = shp.parent.shapes[name]
                except Exception:
                    xl_shape = None

                if xl_shape is not None:
                    for s in iter_shapes_recursive(xl_shape):
                        yield s
    except Exception:
        pass


def _should_include_shape(
    *,
    text: str,
    shape_type_num: Optional[int],
    shape_type_str: Optional[str],
    autoshape_type_str: Optional[str],
    output_mode: str = "standard",
) -> bool:
    """
    Decide whether to emit a shape given output mode.
    - standard: emit if text exists OR the shape is an arrow/line/connector.
    - light: suppress shapes entirely (handled upstream, but guard defensively).
    - verbose: include all (except already-filtered chart/comment/picture/form controls).
    """
    if output_mode == "light":
        return False

    is_arrow = False
    if shape_type_num in (3, 9):  # line/connector
        is_arrow = True
    if autoshape_type_str and "Arrow" in autoshape_type_str:
        is_arrow = True
    if shape_type_str and "Connector" in shape_type_str:
        is_arrow = True
    if shape_type_str and shape_type_str == "Line":
        is_arrow = True

    if output_mode == "standard":
        return bool(text) or is_arrow
    # verbose
    return True


def get_shapes_with_position(workbook: Book, mode: str = "standard") -> Dict[str, List[Shape]]:
    """Scan shapes in a workbook and return per-sheet Shape lists with position info."""
    shape_data: Dict[str, List[Shape]] = {}
    for sheet in workbook.sheets:
        shapes: List[Shape] = []
        seen_names: set[str] = set()
        for root in sheet.shapes:
            for shp in iter_shapes_recursive(root):
                try:
                    shp_name = getattr(shp, "name", None)
                    if shp_name and shp_name in seen_names:
                        continue
                    if shp_name:
                        seen_names.add(shp_name)
                except Exception:
                    pass

                try:
                    type_num = shp.api.Type
                    shape_type_str = MSO_SHAPE_TYPE_MAP.get(
                        type_num, f"Unknown({type_num})"
                    )
                    if shape_type_str in ["Chart", "Comment", "Picture", "FormControl"]:
                        continue
                    autoshape_type_str = None
                    if type_num == 1:
                        astype_num = shp.api.AutoShapeType
                        autoshape_type_str = MSO_AUTO_SHAPE_TYPE_MAP.get(
                            astype_num, f"Unknown({astype_num})"
                        )
                except Exception:
                    type_num = None
                    shape_type_str = None
                    autoshape_type_str = None
                try:
                    text = shp.text.strip() if shp.text else ""
                except Exception:
                    text = ""

                if not _should_include_shape(
                    text=text,
                    shape_type_num=type_num,
                    shape_type_str=shape_type_str,
                    autoshape_type_str=autoshape_type_str,
                    output_mode=mode,
                ):
                    continue

                shape_obj = Shape(
                    text=text,
                    l=int(shp.left),
                    t=int(shp.top),
                    w=int(shp.width) if mode == "verbose" or shape_type_str == "Group" else None,
                    h=int(shp.height) if mode == "verbose" or shape_type_str == "Group" else None,
                    type=f"{shape_type_str}{f'-{autoshape_type_str}' if autoshape_type_str else ''}",
                )
                try:
                    if type_num in (9, 3):
                        angle = compute_line_angle_deg(float(shp.width), float(shp.height))
                        shape_obj.angle_deg = angle
                        shape_obj.direction = angle_to_compass(angle)  # type: ignore
                        try:
                            rot = float(shp.api.Rotation)
                            if abs(rot) > 1e-6:
                                shape_obj.rotation = rot
                        except Exception:
                            pass
                        try:
                            begin_style = int(shp.api.Line.BeginArrowheadStyle)
                            end_style = int(shp.api.Line.EndArrowheadStyle)
                            shape_obj.begin_arrow_style = begin_style
                            shape_obj.end_arrow_style = end_style
                        except Exception:
                            pass
                    elif type_num == 1 and (autoshape_type_str and "Arrow" in autoshape_type_str):
                        try:
                            rot = float(shp.api.Rotation)
                            if abs(rot) > 1e-6:
                                shape_obj.rotation = rot
                        except Exception:
                            pass
                except Exception:
                    pass
                shapes.append(shape_obj)
        shape_data[sheet.name] = shapes
    return shape_data
