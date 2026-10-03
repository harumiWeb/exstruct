"""Shared OOXML shape construction and connector geometry; no COM imports."""

from __future__ import annotations

from collections.abc import Sequence
import math
from typing import Protocol

from ...models import Arrow, Shape, SmartArt
from ..drawing_geometry import angle_to_compass, compute_line_angle_deg
from ..ooxml_drawing import OoxmlConnectorInfo, OoxmlShapeInfo


class ConnectorSnapshot(Protocol):
    """Geometry and direct references needed from optional enrichment snapshots."""

    @property
    def left(self) -> int | None: ...

    @property
    def top(self) -> int | None: ...

    @property
    def width(self) -> int | None: ...

    @property
    def height(self) -> int | None: ...

    @property
    def start_shape_name(self) -> str | None: ...

    @property
    def end_shape_name(self) -> str | None: ...


def _build_shapes_from_ooxml(
    shapes: Sequence[OoxmlShapeInfo],
    connectors: Sequence[OoxmlConnectorInfo],
    *,
    provenance: str = "libreoffice_uno",
) -> list[Shape | Arrow | SmartArt]:
    """Build emitted shape models directly from OOXML drawing metadata.

    Args:
        shapes: Parsed OOXML shape candidates.
        connectors: Parsed OOXML connector candidates.

    Returns:
        Emitted shape and arrow models derived from the OOXML drawing anchors.
    """

    emitted: list[Shape | Arrow | SmartArt] = []
    drawing_to_shape_id: dict[int, int] = {}
    shape_boxes: dict[int, _ShapeBox] = {}
    next_shape_id = 0
    for shape_info in shapes:
        next_shape_id += 1
        shape_id = next_shape_id
        drawing_to_shape_id[shape_info.ref.drawing_id] = shape_id
        box = _to_shape_box(
            shape_id=shape_id,
            left=shape_info.ref.left,
            top=shape_info.ref.top,
            width=shape_info.ref.width,
            height=shape_info.ref.height,
        )
        if box is not None:
            shape_boxes[shape_id] = box
        emitted.append(
            Shape(
                id=shape_id,
                text=shape_info.text,
                l=shape_info.ref.left or 0,
                t=shape_info.ref.top or 0,
                w=shape_info.ref.width,
                h=shape_info.ref.height,
                rotation=shape_info.rotation,
                type=shape_info.shape_type,
                provenance=provenance,
                approximation_level="partial",
                confidence=0.75,
            )
        )
    for connector_info in connectors:
        begin_id, end_id, approximation_level, confidence = _resolve_connector(
            connector_info,
            uno_connector=None,
            drawing_to_shape_id=drawing_to_shape_id,
            shape_name_to_id={},
            shape_boxes=shape_boxes,
        )
        emitted.append(
            Arrow(
                id=None,
                text=connector_info.text,
                l=connector_info.ref.left or 0,
                t=connector_info.ref.top or 0,
                w=connector_info.ref.width,
                h=connector_info.ref.height,
                rotation=connector_info.rotation,
                begin_arrow_style=connector_info.begin_arrow_style,
                end_arrow_style=connector_info.end_arrow_style,
                begin_id=begin_id,
                end_id=end_id,
                direction=_resolve_direction(
                    connector_info=connector_info,
                    uno_connector=None,
                    begin_id=begin_id,
                    end_id=end_id,
                    shape_boxes=shape_boxes,
                ),
                provenance=provenance,
                approximation_level=approximation_level,
                confidence=confidence,
            )
        )
    return emitted


def _resolve_connector(
    connector_info: OoxmlConnectorInfo | None,
    *,
    uno_connector: ConnectorSnapshot | None,
    drawing_to_shape_id: dict[int, int],
    shape_name_to_id: dict[str, int],
    shape_boxes: dict[int, _ShapeBox],
) -> tuple[int | None, int | None, str, float]:
    """Resolve connector endpoints using OOXML refs, UNO refs, or geometry heuristics.

    Args:
        connector_info: OOXML connector metadata when available.
        uno_connector: LibreOffice draw-page connector snapshot when available.
        drawing_to_shape_id: Mapping from OOXML drawing ids to emitted shape ids.
        shape_name_to_id: Mapping from UNO shape names to emitted shape ids.
        shape_boxes: Bounding boxes used for heuristic endpoint matching.

    Returns:
        Tuple of begin id, end id, approximation label, and confidence score.
    """

    begin_id, end_id_resolved, used_ooxml_direct, used_uno_direct = (
        _resolve_direct_connector_ids(
            connector_info=connector_info,
            uno_connector=uno_connector,
            drawing_to_shape_id=drawing_to_shape_id,
            shape_name_to_id=shape_name_to_id,
        )
    )

    if begin_id is not None and end_id_resolved is not None:
        return _classify_connector_resolution(
            begin_id=begin_id,
            end_id=end_id_resolved,
            used_ooxml_direct=used_ooxml_direct,
            used_uno_direct=used_uno_direct,
            used_heuristic=False,
        )

    start_point, end_point = _connector_endpoints(
        connector_info=connector_info,
        uno_connector=uno_connector,
    )
    if begin_id is None:
        begin_id = _nearest_shape_id(start_point, shape_boxes)
    if end_id_resolved is None:
        end_id_resolved = _nearest_shape_id(end_point, shape_boxes)
    return _classify_connector_resolution(
        begin_id=begin_id,
        end_id=end_id_resolved,
        used_ooxml_direct=used_ooxml_direct,
        used_uno_direct=used_uno_direct,
        used_heuristic=True,
    )


def _resolve_direction(
    *,
    connector_info: OoxmlConnectorInfo | None,
    uno_connector: ConnectorSnapshot | None,
    begin_id: int | None = None,
    end_id: int | None = None,
    shape_boxes: dict[int, _ShapeBox] | None = None,
) -> str | None:
    """Infer connector direction from OOXML deltas or resolved endpoint geometry."""

    if connector_info is None:
        return _direction_from_shape_boxes(
            begin_id=begin_id,
            end_id=end_id,
            shape_boxes=shape_boxes,
        )
    dx = connector_info.direction_dx
    dy = connector_info.direction_dy
    if dx is None or dy is None:
        return _direction_from_shape_boxes(
            begin_id=begin_id,
            end_id=end_id,
            shape_boxes=shape_boxes,
        )
    if dx == 0 and dy == 0:
        return _direction_from_shape_boxes(
            begin_id=begin_id,
            end_id=end_id,
            shape_boxes=shape_boxes,
        )
    rotated_dx, rotated_dy = _rotate_connector_delta(
        float(dx),
        float(dy),
        connector_info.rotation,
    )
    angle = compute_line_angle_deg(rotated_dx, rotated_dy)
    return angle_to_compass(angle)


def _connector_endpoints(
    *,
    connector_info: OoxmlConnectorInfo | None,
    uno_connector: ConnectorSnapshot | None,
) -> tuple[tuple[float, float] | None, tuple[float, float] | None]:
    """Return connector endpoints for heuristic endpoint matching."""

    if connector_info is not None:
        left = connector_info.ref.left
        top = connector_info.ref.top
        dx = connector_info.direction_dx
        dy = connector_info.direction_dy
        if (
            left is not None
            and top is not None
            and dx is not None
            and dy is not None
            and (dx != 0 or dy != 0)
        ):
            rotated_dx, rotated_dy = _rotate_connector_delta(
                float(dx),
                float(dy),
                connector_info.rotation,
            )
            start = (float(left), float(top))
            end = (float(left) + rotated_dx, float(top) + rotated_dy)
            return (start, end)

    if uno_connector is None:
        return (None, None)
    left = uno_connector.left
    top = uno_connector.top
    width = uno_connector.width
    height = uno_connector.height
    if left is None or top is None or width is None or height is None:
        return (None, None)
    start = (float(left), float(top))
    end = (float(left + width), float(top + height))
    return (start, end)


def _nearest_shape_id(
    point: tuple[float, float] | None, shape_boxes: dict[int, _ShapeBox]
) -> int | None:
    """Return the closest emitted shape id to a point."""

    if point is None or not shape_boxes:
        return None
    x, y = point
    best_shape_id: int | None = None
    best_distance: float | None = None
    for shape_id, box in shape_boxes.items():
        distance = _distance_to_box(x, y, box)
        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_shape_id = shape_id
    return best_shape_id


def _distance_to_box(x: float, y: float, box: _ShapeBox) -> float:
    """Compute the Euclidean distance from a point to a shape box."""

    dx = max(box.left - x, 0.0, x - box.right)
    dy = max(box.top - y, 0.0, y - box.bottom)
    return math.hypot(dx, dy)


def _rotate_connector_delta(
    dx: float,
    dy: float,
    rotation_deg: float | None,
) -> tuple[float, float]:
    """Rotate an OOXML connector delta into sheet coordinates when needed."""

    if rotation_deg is None:
        return (dx, dy)
    if math.isclose(rotation_deg % 360.0, 0.0, abs_tol=1e-9):
        return (dx, dy)
    length = math.hypot(dx, dy)
    if length == 0.0:
        return (dx, dy)
    angle_rad = math.radians(compute_line_angle_deg(dx, dy) + rotation_deg)
    return (length * math.cos(angle_rad), length * math.sin(angle_rad))


def _to_shape_box(
    *,
    shape_id: int,
    left: int | None,
    top: int | None,
    width: int | None,
    height: int | None,
) -> _ShapeBox | None:
    """Build a shape box when complete rectangle geometry is available."""

    if left is None or top is None or width is None or height is None:
        return None
    return _ShapeBox(
        shape_id=shape_id,
        left=float(left),
        top=float(top),
        right=float(left + width),
        bottom=float(top + height),
    )


def _direction_from_shape_boxes(
    *,
    begin_id: int | None,
    end_id: int | None,
    shape_boxes: dict[int, _ShapeBox] | None,
) -> str | None:
    """Infer a connector direction from resolved endpoint shape centers."""

    if begin_id is None or end_id is None or shape_boxes is None:
        return None
    begin_box = shape_boxes.get(begin_id)
    end_box = shape_boxes.get(end_id)
    if begin_box is None or end_box is None:
        return None
    begin_center = _shape_box_center(begin_box)
    end_center = _shape_box_center(end_box)
    dx = end_center[0] - begin_center[0]
    dy = end_center[1] - begin_center[1]
    if dx == 0 and dy == 0:
        return None
    angle = compute_line_angle_deg(dx, dy)
    return angle_to_compass(angle)


def _shape_box_center(box: _ShapeBox) -> tuple[float, float]:
    """Return the center point of a matched emitted shape box."""

    return ((box.left + box.right) / 2.0, (box.top + box.bottom) / 2.0)


def _resolve_direct_connector_ids(
    *,
    connector_info: OoxmlConnectorInfo | None,
    uno_connector: ConnectorSnapshot | None,
    drawing_to_shape_id: dict[int, int],
    shape_name_to_id: dict[str, int],
) -> tuple[int | None, int | None, bool, bool]:
    """Resolve direct connector endpoints from OOXML refs and UNO shape refs."""

    begin_id: int | None = None
    end_id: int | None = None
    used_ooxml_direct = False
    used_uno_direct = False
    if connector_info is not None:
        start_id = connector_info.connection.start_drawing_id
        target_id = connector_info.connection.end_drawing_id
        begin_id = drawing_to_shape_id.get(start_id) if start_id is not None else None
        end_id = drawing_to_shape_id.get(target_id) if target_id is not None else None
        used_ooxml_direct = begin_id is not None or end_id is not None
    if uno_connector is not None:
        if begin_id is None and uno_connector.start_shape_name is not None:
            begin_id = shape_name_to_id.get(uno_connector.start_shape_name)
            used_uno_direct = used_uno_direct or begin_id is not None
        if end_id is None and uno_connector.end_shape_name is not None:
            end_id = shape_name_to_id.get(uno_connector.end_shape_name)
            used_uno_direct = used_uno_direct or end_id is not None
    return (begin_id, end_id, used_ooxml_direct, used_uno_direct)


def _classify_connector_resolution(
    *,
    begin_id: int | None,
    end_id: int | None,
    used_ooxml_direct: bool,
    used_uno_direct: bool,
    used_heuristic: bool,
) -> tuple[int | None, int | None, str, float]:
    """Classify connector provenance once endpoint resolution is complete."""

    if used_heuristic:
        return (begin_id, end_id, "heuristic", 0.6)
    if used_ooxml_direct and used_uno_direct:
        return (begin_id, end_id, "partial", 0.9)
    if used_ooxml_direct:
        return (begin_id, end_id, "direct", 1.0)
    if used_uno_direct:
        return (begin_id, end_id, "direct", 0.9)
    return (begin_id, end_id, "heuristic", 0.6)


class _ShapeBox:
    """Axis-aligned bounding box for emitted shapes."""

    def __init__(
        self, *, shape_id: int, left: float, top: float, right: float, bottom: float
    ) -> None:
        """Store bounding-box coordinates for emitted shape matching."""

        self.shape_id = shape_id
        self.left = left
        self.top = top
        self.right = right
        self.bottom = bottom
