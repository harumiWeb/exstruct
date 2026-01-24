from __future__ import annotations

from collections.abc import Iterable, Iterator
import math
from typing import Literal, Protocol, SupportsInt, cast, runtime_checkable

import xlwings as xw
from xlwings import Book

from ..models import Arrow, Shape, SmartArt, SmartArtNode
from ..models.maps import MSO_AUTO_SHAPE_TYPE_MAP, MSO_SHAPE_TYPE_MAP


def compute_line_angle_deg(w: float, h: float) -> float:
    """
    Compute the clockwise angle (in degrees) in Excel coordinates where 0° points East.

    Parameters:
        w (float): Horizontal delta (width, positive to the right).
        h (float): Vertical delta (height, positive downward).

    Returns:
        float: Angle in degrees measured clockwise from East (e.g., 0° = East, 90° = South).
    """
    return math.degrees(math.atan2(h, w)) % 360.0


def angle_to_compass(
    angle: float,
) -> Literal["E", "SE", "S", "SW", "W", "NW", "N", "NE"]:
    """
    Map an angle in degrees to one of eight compass directions.

    The angle is interpreted with 0 degrees at East and increasing values rotating counterclockwise (45 -> NE, 90 -> N).

    Parameters:
        angle (float): Angle in degrees.

    Returns:
        str: One of `"E"`, `"SE"`, `"S"`, `"SW"`, `"W"`, `"NW"`, `"N"`, or `"NE"` corresponding to the nearest 8-point compass direction.
    """
    dirs = ["E", "NE", "N", "NW", "W", "SW", "S", "SE"]
    idx = int(((angle + 22.5) % 360) // 45)
    return cast(Literal["E", "SE", "S", "SW", "W", "NW", "N", "NE"], dirs[idx])


def coord_to_cell_by_edges(
    row_edges: list[float], col_edges: list[float], x: float, y: float
) -> str | None:
    """
    Estimate the Excel A1-style cell that contains a point given cumulative row and column edge coordinates.

    Parameters:
        row_edges (list[float]): Monotonic list of cumulative vertical edges (top-to-bottom). Consecutive entries define row spans.
        col_edges (list[float]): Monotonic list of cumulative horizontal edges (left-to-right). Consecutive entries define column spans.
        x (float): Horizontal coordinate (same coordinate system as col_edges).
        y (float): Vertical coordinate (same coordinate system as row_edges).

    Returns:
        str | None: A1-style cell address (e.g., "B3") if the point falls inside the grid; `None` if the point is outside the provided edge ranges. Intervals are treated as left-inclusive and right-exclusive: [edge_i, edge_{i+1}).
    """

    def find_index(edges: list[float], pos: float) -> int | None:
        """Return the 1-based index for a position inside edge intervals."""
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
    Determine whether a shape should be included in the output based on its properties and the selected output mode.

    Modes:
    - "light": always exclude shapes.
    - "standard": include when the shape has text or represents a relationship (line/connector).
    - "verbose": include all shapes (other global exclusions are handled elsewhere).

    Parameters:
        output_mode (str): One of "light", "standard", or "verbose"; controls inclusion rules.

    Returns:
        bool: `True` if the shape should be emitted, `False` otherwise.
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


@runtime_checkable
class _TextRangeLike(Protocol):
    """Text range interface for SmartArt nodes."""

    Text: str | None


@runtime_checkable
class _TextFrameLike(Protocol):
    """Text frame interface for SmartArt nodes."""

    HasText: bool
    TextRange: _TextRangeLike


@runtime_checkable
class _SmartArtNodeLike(Protocol):
    """SmartArt node interface."""

    Level: int
    TextFrame2: _TextFrameLike


@runtime_checkable
class _SmartArtLike(Protocol):
    """SmartArt interface."""

    Layout: object
    AllNodes: Iterable[_SmartArtNodeLike]


def _shape_has_smartart(shp: xw.Shape) -> bool:
    """
    Determine whether a shape exposes SmartArt content.

    Returns:
        bool: `True` if the shape exposes SmartArt (i.e., has an accessible `HasSmartArt` attribute), `False` otherwise.
    """
    try:
        api = shp.api
    except Exception:
        return False
    try:
        return bool(api.HasSmartArt)
    except Exception:
        return False


def _get_smartart_layout_name(smartart: _SmartArtLike | None) -> str:
    """
    Get the SmartArt layout name or "Unknown" if it cannot be determined.

    Returns:
        layout_name (str): The layout name from `smartart.Layout.Name`, or "Unknown" when `smartart` is None or the name cannot be retrieved.
    """
    if smartart is None:
        return "Unknown"
    try:
        layout = getattr(smartart, "Layout", None)
        name = getattr(layout, "Name", None)
        return str(name) if name is not None else "Unknown"
    except Exception:
        return "Unknown"


def _collect_smartart_node_info(
    smartart: _SmartArtLike | None,
) -> list[tuple[int, str]]:
    """
    Extract a list of (level, text) tuples for each node present in the given SmartArt.

    Parameters:
        smartart (_SmartArtLike | None): A SmartArt-like COM object or `None`. If `None` or inaccessible, no nodes are collected.

    Returns:
        list[tuple[int, str]]: A list of tuples where each tuple is (node level, node text). Returns an empty list if the SmartArt is `None`, inaccessible, or if nodes lack a numeric level.
    """
    nodes_info: list[tuple[int, str]] = []
    if smartart is None:
        return nodes_info
    try:
        all_nodes = smartart.AllNodes
    except Exception:
        return nodes_info

    for node in all_nodes:
        level = _get_smartart_node_level(node)
        if level is None:
            continue
        text = ""
        try:
            text_frame = node.TextFrame2
            if text_frame.HasText:
                text_value = text_frame.TextRange.Text
                text = str(text_value) if text_value is not None else ""
        except Exception:
            text = ""
        nodes_info.append((level, text))
    return nodes_info


def _get_smartart_node_level(node: _SmartArtNodeLike) -> int | None:
    """
    Get the numerical level of a SmartArt node.

    Returns:
        int | None: The node's level as an integer, or `None` if the level is missing or cannot be converted to an integer.
    """
    try:
        return int(node.Level)
    except Exception:
        return None


def _build_smartart_tree(nodes_info: list[tuple[int, str]]) -> list[SmartArtNode]:
    """
    Build a nested tree of SmartArtNode objects from a flat list of (level, text) tuples.

    Parameters:
        nodes_info (list[tuple[int, str]]): Ordered tuples where each tuple is (level, text);
            `level` is the hierarchical depth (integer) and `text` is the node label.

    Returns:
        roots (list[SmartArtNode]): Top-level SmartArtNode instances whose `kids` lists
            contain their nested child nodes according to the provided levels.
    """
    roots: list[SmartArtNode] = []
    stack: list[tuple[int, SmartArtNode]] = []
    for level, text in nodes_info:
        node = SmartArtNode(text=text, kids=[])
        while stack and stack[-1][0] >= level:
            stack.pop()
        if stack:
            stack[-1][1].kids.append(node)
        else:
            roots.append(node)
        stack.append((level, node))
    return roots


def _extract_smartart_nodes(smartart: _SmartArtLike | None) -> list[SmartArtNode]:
    """
    Convert a SmartArt COM object into a list of root SmartArtNode trees.

    Parameters:
        smartart (_SmartArtLike | None): SmartArt-like COM object to extract nodes from; pass `None` to produce an empty list.

    Returns:
        list[SmartArtNode]: Root nodes representing the hierarchical SmartArt structure (each node contains its text and children).
    """
    nodes_info = _collect_smartart_node_info(smartart)
    return _build_smartart_tree(nodes_info)


def get_shapes_with_position(  # noqa: C901
    workbook: Book, mode: str = "standard"
) -> dict[str, list[Shape | Arrow | SmartArt]]:
    """
    Scan all shapes in each worksheet and collect their positional and metadata information.

    Parameters:
        workbook (Book): The xlwings workbook to scan.
        mode (str): Output detail level; "light" skips most shapes, "standard" includes shapes with text or relationships, and "verbose" includes full size/rotation details.

    Returns:
        dict[str, list[Shape | Arrow | SmartArt]]: Mapping of sheet name to a list of collected shape objects (Shape, Arrow, or SmartArt) containing position (left/top), optional size (width/height), textual content, and other captured metadata (ids, directions, connections, layout/nodes for SmartArt).
    """
    shape_data: dict[str, list[Shape | Arrow | SmartArt]] = {}
    for sheet in workbook.sheets:
        shapes: list[Shape | Arrow | SmartArt] = []
        excel_names: list[tuple[str, int]] = []
        node_index = 0
        pending_connections: list[tuple[Arrow, str | None, str | None]] = []
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

                if mode == "light":
                    continue

                has_smartart = _shape_has_smartart(shp)
                if not has_smartart and not _should_include_shape(
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
                    "Connector" in shape_type_str
                    or shape_type_str in ("Line", "ConnectLine")
                ):
                    is_relationship_geom = True
                if shape_name and ("Connector" in shape_name or "Line" in shape_name):
                    is_relationship_geom = True

                shape_id = None
                if not is_relationship_geom:
                    node_index += 1
                    shape_id = node_index

                excel_name = shape_name if isinstance(shape_name, str) else None

                shape_obj: Shape | Arrow | SmartArt
                if has_smartart:
                    smartart_obj: _SmartArtLike | None = None
                    try:
                        smartart_obj = shp.api.SmartArt
                    except Exception:
                        smartart_obj = None
                    shape_obj = SmartArt(
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
                        layout=_get_smartart_layout_name(smartart_obj),
                        nodes=_extract_smartart_nodes(smartart_obj),
                    )
                elif is_relationship_geom:
                    shape_obj = Arrow(
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
                    )
                else:
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
                        if isinstance(shape_obj, Arrow):
                            shape_obj.direction = angle_to_compass(angle)
                        try:
                            rot = float(shp.api.Rotation)
                            if abs(rot) > 1e-6:
                                shape_obj.rotation = rot
                        except Exception:
                            pass
                        try:
                            begin_style = int(shp.api.Line.BeginArrowheadStyle)
                            end_style = int(shp.api.Line.EndArrowheadStyle)
                            if isinstance(shape_obj, Arrow):
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
                if isinstance(shape_obj, Arrow):
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
