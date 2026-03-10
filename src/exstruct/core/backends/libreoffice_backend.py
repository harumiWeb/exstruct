"""LibreOffice-backed rich shape and chart extraction helpers."""

from __future__ import annotations

from collections.abc import Callable, Sequence
import logging
import math
from pathlib import Path

from ...models import Arrow, Chart, Shape, SmartArt
from ..libreoffice import (
    LibreOfficeChartGeometry,
    LibreOfficeDrawPageShape,
    LibreOfficeSession,
)
from ..ooxml_drawing import OoxmlConnectorInfo, OoxmlShapeInfo, read_sheet_drawings
from ..shapes import angle_to_compass, compute_line_angle_deg
from .base import ChartData, RichBackend, ShapeData

logger = logging.getLogger(__name__)


class LibreOfficeRichBackend(RichBackend):
    """Best-effort rich extraction backend gated by LibreOffice runtime availability."""

    def __init__(
        self,
        file_path: Path,
        *,
        session_factory: Callable[[], LibreOfficeSession] = LibreOfficeSession.from_env,
    ) -> None:
        """Store the workbook path and session factory used for lazy LibreOffice extraction."""

        self.file_path = file_path
        self._session_factory = session_factory
        self._chart_geometries: dict[str, list[LibreOfficeChartGeometry]] | None = None
        self._draw_page_shapes: dict[str, list[LibreOfficeDrawPageShape]] | None = None

    def extract_shapes(self, *, mode: str) -> dict[str, list[Shape | Arrow | SmartArt]]:
        """Extract LibreOffice-mode shapes and connectors for each worksheet.

        Args:
            mode: Requested extraction mode. Only ``"libreoffice"`` is supported.

        Returns:
            Mapping of sheet names to emitted shape, arrow, and SmartArt models.

        Raises:
            ValueError: If a non-LibreOffice mode is requested.
        """

        if mode != "libreoffice":
            raise ValueError("LibreOfficeRichBackend only supports libreoffice mode.")
        drawings = read_sheet_drawings(self.file_path)
        draw_page_shapes = self._read_draw_page_shapes()
        shape_data: ShapeData = {}
        sheet_names = list(dict.fromkeys([*drawings.keys(), *draw_page_shapes.keys()]))
        for sheet_name in sheet_names:
            drawing = drawings.get(sheet_name)
            snapshots = draw_page_shapes.get(sheet_name, [])
            if snapshots:
                _log_unmatched_ooxml_candidates(
                    sheet_name=sheet_name,
                    snapshots=snapshots,
                    drawing_shapes=drawing.shapes if drawing is not None else [],
                    drawing_connectors=drawing.connectors
                    if drawing is not None
                    else [],
                )
                # UNO draw-page snapshots define the canonical emitted order for v1.
                # Unmatched OOXML-only shapes/connectors remain supplemental metadata
                # and are intentionally not appended to the emitted list.
                shape_data[sheet_name] = _build_shapes_from_draw_page(
                    snapshots,
                    drawing_shapes=drawing.shapes if drawing is not None else [],
                    drawing_connectors=drawing.connectors
                    if drawing is not None
                    else [],
                )
                continue
            if drawing is not None:
                shape_data[sheet_name] = _build_shapes_from_ooxml(
                    drawing.shapes,
                    drawing.connectors,
                )
        return shape_data

    def extract_charts(self, *, mode: str) -> dict[str, list[Chart]]:
        """Extract LibreOffice-mode charts for each worksheet.

        Args:
            mode: Requested extraction mode. Only ``"libreoffice"`` is supported.

        Returns:
            Mapping of sheet names to emitted chart models.

        Raises:
            ValueError: If a non-LibreOffice mode is requested.
        """

        if mode != "libreoffice":
            raise ValueError("LibreOfficeRichBackend only supports libreoffice mode.")
        drawings = read_sheet_drawings(self.file_path)
        chart_geometries = self._read_chart_geometries()
        chart_data: ChartData = {}
        for sheet_name, drawing in drawings.items():
            charts: list[Chart] = []
            geometry_matches = _match_chart_geometries(
                drawing.charts,
                chart_geometries.get(sheet_name, []),
            )
            for chart_info, geometry in zip(
                drawing.charts, geometry_matches, strict=False
            ):
                left = chart_info.anchor_left or 0
                top = chart_info.anchor_top or 0
                width = chart_info.anchor_width
                height = chart_info.anchor_height
                confidence = 0.5
                if geometry is not None:
                    left = geometry.left if geometry.left is not None else left
                    top = geometry.top if geometry.top is not None else top
                    width = geometry.width if geometry.width is not None else width
                    height = geometry.height if geometry.height is not None else height
                    confidence = 0.8
                charts.append(
                    Chart(
                        name=chart_info.name,
                        chart_type=chart_info.chart_type,
                        title=chart_info.title,
                        y_axis_title=chart_info.y_axis_title,
                        y_axis_range=chart_info.y_axis_range,
                        w=width,
                        h=height,
                        series=chart_info.series,
                        l=left,
                        t=top,
                        provenance="libreoffice_uno",
                        approximation_level="partial",
                        confidence=confidence,
                    )
                )
            chart_data[sheet_name] = charts
        return chart_data

    def _read_chart_geometries(self) -> dict[str, list[LibreOfficeChartGeometry]]:
        """Load and cache chart geometry snapshots from the LibreOffice session."""

        if self._chart_geometries is not None:
            return self._chart_geometries
        with self._session_factory() as session:
            if hasattr(session, "extract_chart_geometries"):
                self._chart_geometries = session.extract_chart_geometries(
                    self.file_path
                )
            else:
                self._chart_geometries = {}
        return self._chart_geometries

    def _read_draw_page_shapes(self) -> dict[str, list[LibreOfficeDrawPageShape]]:
        """Load and cache draw-page shape snapshots from the LibreOffice session."""

        if self._draw_page_shapes is not None:
            return self._draw_page_shapes
        with self._session_factory() as session:
            if hasattr(session, "extract_draw_page_shapes"):
                self._draw_page_shapes = session.extract_draw_page_shapes(
                    self.file_path
                )
            else:
                self._draw_page_shapes = {}
        return self._draw_page_shapes


def _build_shapes_from_ooxml(
    shapes: Sequence[OoxmlShapeInfo],
    connectors: Sequence[OoxmlConnectorInfo],
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
                provenance="libreoffice_uno",
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
                provenance="libreoffice_uno",
                approximation_level=approximation_level,
                confidence=confidence,
            )
        )
    return emitted


def _build_shapes_from_draw_page(
    snapshots: Sequence[LibreOfficeDrawPageShape],
    *,
    drawing_shapes: Sequence[OoxmlShapeInfo],
    drawing_connectors: Sequence[OoxmlConnectorInfo],
) -> list[Shape | Arrow | SmartArt]:
    """Merge UNO draw-page snapshots with OOXML drawing metadata.

    Args:
        snapshots: LibreOffice draw-page snapshots for one worksheet.
        drawing_shapes: OOXML shape candidates for the worksheet.
        drawing_connectors: OOXML connector candidates for the worksheet.

    Returns:
        Emitted shape and arrow models built from the combined metadata.
    """

    emitted: list[Shape | Arrow | SmartArt] = []
    snapshot_shapes = [snapshot for snapshot in snapshots if not snapshot.is_connector]
    snapshot_connectors = [snapshot for snapshot in snapshots if snapshot.is_connector]
    matched_shapes = _match_shape_infos(snapshot_shapes, drawing_shapes)
    matched_connectors = _match_connector_infos(snapshot_connectors, drawing_connectors)
    drawing_to_shape_id: dict[int, int] = {}
    shape_name_to_id: dict[str, int] = {}
    shape_boxes: dict[int, _ShapeBox] = {}
    next_shape_id = 0
    assigned_shapes: list[
        tuple[LibreOfficeDrawPageShape, OoxmlShapeInfo | None, int]
    ] = []

    for snapshot, shape_info in zip(snapshot_shapes, matched_shapes, strict=False):
        next_shape_id += 1
        shape_id = next_shape_id
        assigned_shapes.append((snapshot, shape_info, shape_id))
        if shape_info is not None:
            drawing_to_shape_id[shape_info.ref.drawing_id] = shape_id
        shape_name_to_id[snapshot.name] = shape_id
        box = _to_shape_box(
            shape_id=shape_id,
            left=snapshot.left,
            top=snapshot.top,
            width=snapshot.width,
            height=snapshot.height,
        )
        if box is not None:
            shape_boxes[shape_id] = box

    shape_index = 0
    connector_index = 0

    for snapshot in snapshots:
        if snapshot.is_connector:
            connector_info = matched_connectors[connector_index]
            connector_index += 1
            begin_id, end_id, approximation_level, confidence = _resolve_connector(
                connector_info,
                uno_connector=snapshot,
                drawing_to_shape_id=drawing_to_shape_id,
                shape_name_to_id=shape_name_to_id,
                shape_boxes=shape_boxes,
            )
            emitted.append(
                Arrow(
                    id=None,
                    text=snapshot.text
                    or (connector_info.text if connector_info else ""),
                    l=_first_int(
                        snapshot.left,
                        connector_info.ref.left if connector_info else None,
                        default=0,
                    ),
                    t=_first_int(
                        snapshot.top,
                        connector_info.ref.top if connector_info else None,
                        default=0,
                    ),
                    w=_first_optional_int(
                        snapshot.width,
                        connector_info.ref.width if connector_info else None,
                    ),
                    h=_first_optional_int(
                        snapshot.height,
                        connector_info.ref.height if connector_info else None,
                    ),
                    rotation=_first_optional_float(
                        snapshot.rotation,
                        connector_info.rotation if connector_info else None,
                    ),
                    begin_arrow_style=connector_info.begin_arrow_style
                    if connector_info is not None
                    else None,
                    end_arrow_style=connector_info.end_arrow_style
                    if connector_info is not None
                    else None,
                    begin_id=begin_id,
                    end_id=end_id,
                    direction=_resolve_direction(
                        connector_info=connector_info,
                        uno_connector=snapshot,
                        begin_id=begin_id,
                        end_id=end_id,
                        shape_boxes=shape_boxes,
                    ),
                    provenance="libreoffice_uno",
                    approximation_level=approximation_level,
                    confidence=confidence,
                )
            )
            continue

        shape_snapshot, shape_info, shape_id = assigned_shapes[shape_index]
        shape_index += 1
        emitted.append(
            Shape(
                id=shape_id,
                text=shape_snapshot.text or (shape_info.text if shape_info else ""),
                l=_first_int(
                    shape_snapshot.left,
                    shape_info.ref.left if shape_info else None,
                    default=0,
                ),
                t=_first_int(
                    shape_snapshot.top,
                    shape_info.ref.top if shape_info else None,
                    default=0,
                ),
                w=_first_optional_int(
                    shape_snapshot.width,
                    shape_info.ref.width if shape_info else None,
                ),
                h=_first_optional_int(
                    shape_snapshot.height,
                    shape_info.ref.height if shape_info else None,
                ),
                rotation=_first_optional_float(
                    shape_snapshot.rotation,
                    shape_info.rotation if shape_info else None,
                ),
                type=shape_info.shape_type
                if shape_info is not None and shape_info.shape_type
                else _shape_type_from_uno(shape_snapshot.shape_type),
                provenance="libreoffice_uno",
                approximation_level="partial",
                confidence=0.75,
            )
        )
    return emitted


def _log_unmatched_ooxml_candidates(
    *,
    sheet_name: str,
    snapshots: Sequence[LibreOfficeDrawPageShape],
    drawing_shapes: Sequence[OoxmlShapeInfo],
    drawing_connectors: Sequence[OoxmlConnectorInfo],
) -> None:
    """Emit a debug log when snapshot-backed sheets drop OOXML-only candidates."""

    snapshot_shapes = [snapshot for snapshot in snapshots if not snapshot.is_connector]
    snapshot_connectors = [snapshot for snapshot in snapshots if snapshot.is_connector]
    unmatched_shape_count = len(drawing_shapes) - sum(
        shape_info is not None
        for shape_info in _match_shape_infos(snapshot_shapes, drawing_shapes)
    )
    unmatched_connector_count = len(drawing_connectors) - sum(
        connector_info is not None
        for connector_info in _match_connector_infos(
            snapshot_connectors, drawing_connectors
        )
    )
    if unmatched_shape_count <= 0 and unmatched_connector_count <= 0:
        return
    logger.debug(
        "Skipping %d OOXML-only shapes and %d OOXML-only connectors on sheet %s "
        "because UNO draw-page snapshots define the canonical emitted order.",
        unmatched_shape_count,
        unmatched_connector_count,
        sheet_name,
    )


def _resolve_connector(
    connector_info: OoxmlConnectorInfo | None,
    *,
    uno_connector: LibreOfficeDrawPageShape | None,
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
    uno_connector: LibreOfficeDrawPageShape | None,
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
    uno_connector: LibreOfficeDrawPageShape | None,
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


def _match_shape_infos(
    snapshots: Sequence[LibreOfficeDrawPageShape],
    candidates: Sequence[OoxmlShapeInfo],
) -> list[OoxmlShapeInfo | None]:
    """Match UNO shape snapshots to OOXML shape metadata."""

    return [
        candidates[index] if index is not None else None
        for index in _match_by_name_then_order(
            [snapshot.name for snapshot in snapshots],
            [candidate.ref.name for candidate in candidates],
        )
    ]


def _match_connector_infos(
    snapshots: Sequence[LibreOfficeDrawPageShape],
    candidates: Sequence[OoxmlConnectorInfo],
) -> list[OoxmlConnectorInfo | None]:
    """Match UNO connector snapshots to OOXML connector metadata."""

    return [
        candidates[index] if index is not None else None
        for index in _match_by_name_then_order(
            [snapshot.name for snapshot in snapshots],
            [candidate.ref.name for candidate in candidates],
        )
    ]


def _match_by_name_then_order(
    snapshot_names: Sequence[str],
    candidate_names: Sequence[str],
) -> list[int | None]:
    """Match snapshot names to candidate names by name first, then by order."""

    matches: list[int | None] = [None] * len(snapshot_names)
    unused = list(range(len(candidate_names)))

    for index, snapshot_name in enumerate(snapshot_names):
        matched_index = next(
            (
                candidate_index
                for candidate_index in unused
                if candidate_names[candidate_index] == snapshot_name
            ),
            None,
        )
        if matched_index is None:
            continue
        matches[index] = matched_index
        unused.remove(matched_index)

    remaining_snapshot_indexes = [
        index for index, match in enumerate(matches) if match is None
    ]
    if remaining_snapshot_indexes:
        for snapshot_index, candidate_index in zip(
            remaining_snapshot_indexes, unused, strict=False
        ):
            matches[snapshot_index] = candidate_index

    return matches


def _shape_type_from_uno(shape_type: str | None) -> str | None:
    """Collapse a fully qualified UNO shape type to its leaf name."""

    if not shape_type:
        return None
    return shape_type.rsplit(".", 1)[-1]


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
    uno_connector: LibreOfficeDrawPageShape | None,
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


def _match_chart_geometries(
    charts: Sequence[object],
    candidates: Sequence[LibreOfficeChartGeometry],
) -> list[LibreOfficeChartGeometry | None]:
    """Match OOXML chart entries to LibreOffice chart geometry candidates."""

    matches: list[LibreOfficeChartGeometry | None] = [None] * len(charts)
    unused = list(range(len(candidates)))

    for index, chart in enumerate(charts):
        chart_name = getattr(chart, "name", None)
        if not isinstance(chart_name, str):
            continue
        matched_index = next(
            (
                candidate_index
                for candidate_index in unused
                if candidates[candidate_index].name == chart_name
                or candidates[candidate_index].persist_name == chart_name
            ),
            None,
        )
        if matched_index is None:
            continue
        matches[index] = candidates[matched_index]
        unused.remove(matched_index)

    remaining_chart_indexes = [
        index for index, match in enumerate(matches) if match is None
    ]
    if remaining_chart_indexes and len(remaining_chart_indexes) == len(unused):
        for chart_index, candidate_index in zip(
            remaining_chart_indexes, unused, strict=False
        ):
            matches[chart_index] = candidates[candidate_index]

    return matches


def _first_int(*values: int | None, default: int) -> int:
    """Return the first non-``None`` integer or the provided default."""

    for value in values:
        if value is not None:
            return value
    return default


def _first_optional_int(*values: int | None) -> int | None:
    """Return the first non-``None`` integer, if any."""

    for value in values:
        if value is not None:
            return value
    return None


def _first_optional_float(*values: float | None) -> float | None:
    """Return the first non-``None`` float, if any."""

    for value in values:
        if value is not None:
            return value
    return None


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
