from __future__ import annotations

from collections.abc import Iterator
import math
from typing import SupportsInt, cast

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
    row_edges: list[float], col_edges: list[float], x: float, y: float
) -> str | None:
    """Estimate cell address from coordinates and cumulative edges; return None if out of range."""

    def find_index(edges: list[float], pos: float) -> int | None:
        for i in range(1, len(edges)):
            if edges[i - 1] <= pos < edges[i]:
                return i
        return None

    r = find_index(row_edges, y)
    c = find_index(col_edges, x)
    if r is None or c is None:
        return None
    return f"{xw.utils.col_name(c)}{r}"


def has_arrow(style_val: object) -> bool:
    """Return True if Excel arrow style value indicates an arrowhead."""
    try:
        v = int(cast(SupportsInt, style_val))
        return v != 0
    except Exception:
        return False


def iter_shapes_recursive(shp: xw.Shape) -> Iterator[xw.Shape]:
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
                    yield from iter_shapes_recursive(xl_shape)
    except Exception:
        pass


def _should_include_shape(
    *,
    text: str,
    shape_type_num: int | None,
    shape_type_str: str | None,
    autoshape_type_str: str | None,
    shape_name: str | None,
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

    is_relationship = False
    if shape_type_num in (3, 9):  # line/connector
        is_relationship = True
    if autoshape_type_str and (
        "Arrow" in autoshape_type_str or "Connector" in autoshape_type_str
    ):
        is_relationship = True
    if shape_type_str and (
        "Connector" in shape_type_str or shape_type_str in ("Line", "ConnectLine")
    ):
        is_relationship = True
    if shape_name and ("Connector" in shape_name or "Line" in shape_name):
        is_relationship = True

    if output_mode == "standard":
        return bool(text) or is_relationship
    # verbose
    return True


def get_shapes_with_position(  # noqa: C901
    workbook: Book, mode: str = "standard"
) -> dict[str, list[Shape]]:
    """Scan shapes in a workbook and return per-sheet Shape lists with position info."""
    shape_data: dict[str, list[Shape]] = {}
    for sheet in workbook.sheets:
        shapes: list[Shape] = []
        excel_names: list[tuple[str, int]] = []
        node_index = 0
        pending_connections: list[tuple[Shape, str | None, str | None]] = []
        for root in sheet.shapes:
            for shp in iter_shapes_recursive(root):
                try:
                    shape_name = getattr(shp, "name", None)
                except Exception:
                    shape_name = None
                try:
                    type_num = shp.api.Type
                    shape_type_str = MSO_SHAPE_TYPE_MAP.get(
                        type_num, f"Unknown({type_num})"
                    )
                    if shape_type_str in ["Chart", "Comment", "Picture", "FormControl"]:
                        continue
                    autoshape_type_str = None
                    try:
                        astype_num = shp.api.AutoShapeType
                        autoshape_type_str = MSO_AUTO_SHAPE_TYPE_MAP.get(
                            astype_num, f"Unknown({astype_num})"
                        )
                    except Exception:
                        autoshape_type_str = None
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
                    shape_name=shape_name,
                    output_mode=mode,
                ):
                    continue

                if (
                    autoshape_type_str
                    and autoshape_type_str == "NotPrimitive"
                    and shape_name
                ):
                    type_label = shape_name
                else:
                    type_label = (
                        f"{shape_type_str}-{autoshape_type_str}"
                        if autoshape_type_str
                        else (shape_type_str or shape_name or "Unknown")
                    )

                is_relationship_geom = False
                if type_num in (3, 9):
                    is_relationship_geom = True
                if autoshape_type_str and (
                    "Arrow" in autoshape_type_str or "Connector" in autoshape_type_str
                ):
                    is_relationship_geom = True
                if shape_type_str and (
                    "Connector" in shape_type_str or shape_type_str in ("Line", "ConnectLine")
                ):
                    is_relationship_geom = True
                if shape_name and ("Connector" in shape_name or "Line" in shape_name):
                    is_relationship_geom = True

                shape_id = None
                if not is_relationship_geom:
                    node_index += 1
                    shape_id = node_index

                excel_name = shape_name if isinstance(shape_name, str) else None

                shape_obj = Shape(
                    id=shape_id,
                    text=text,
                    l=int(shp.left),
                    t=int(shp.top),
                    w=int(shp.width)
                    if mode == "verbose" or shape_type_str == "Group"
                    else None,
                    h=int(shp.height)
                    if mode == "verbose" or shape_type_str == "Group"
                    else None,
                    type=type_label,
                )
                if excel_name:
                    if shape_id is not None:
                        excel_names.append((excel_name, shape_id))
                try:
                    begin_name: str | None = None
                    end_name: str | None = None
                    if is_relationship_geom:
                        angle = compute_line_angle_deg(
                            float(shp.width), float(shp.height)
                        )
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
                        # Connector begin/end connected shapes (if this shape is a connector).
                        try:
                            connector = shp.api.ConnectorFormat
                            try:
                                begin_shape = connector.BeginConnectedShape
                                if begin_shape is not None:
                                    name = getattr(begin_shape, "Name", None)
                                    if isinstance(name, str):
                                        begin_name = name
                            except Exception:
                                pass
                            try:
                                end_shape = connector.EndConnectedShape
                                if end_shape is not None:
                                    name = getattr(end_shape, "Name", None)
                                    if isinstance(name, str):
                                        end_name = name
                            except Exception:
                                pass
                        except Exception:
                            # Not a connector or ConnectorFormat is unavailable.
                            pass
                    elif type_num == 1 and (
                        autoshape_type_str and "Arrow" in autoshape_type_str
                    ):
                        try:
                            rot = float(shp.api.Rotation)
                            if abs(rot) > 1e-6:
                                shape_obj.rotation = rot
                        except Exception:
                            pass
                except Exception:
                    pass
                pending_connections.append((shape_obj, begin_name, end_name))
                shapes.append(shape_obj)
        if pending_connections:
            name_to_id = {name: sid for name, sid in excel_names}
            for shape_obj, begin_name, end_name in pending_connections:
                if begin_name:
                    shape_obj.begin_id = name_to_id.get(begin_name)
                if end_name:
                    shape_obj.end_id = name_to_id.get(end_name)
        shape_data[sheet.name] = shapes
    return shape_data
