from __future__ import annotations

from collections.abc import Callable
import math
from pathlib import Path

from ...models import Arrow, Chart, Shape, SmartArt
from ..libreoffice import LibreOfficeSession
from ..ooxml_drawing import OoxmlConnectorInfo, OoxmlShapeInfo, read_sheet_drawings
from ..shapes import angle_to_compass, compute_line_angle_deg
from .base import ChartData, RichBackend, ShapeData


class LibreOfficeRichBackend(RichBackend):
    """Best-effort rich extraction backend gated by LibreOffice runtime availability."""

    def __init__(
        self,
        file_path: Path,
        *,
        session_factory: Callable[[], LibreOfficeSession] = LibreOfficeSession.from_env,
    ) -> None:
        self.file_path = file_path
        self._session_factory = session_factory
        self._runtime_checked = False

    def extract_shapes(self, *, mode: str) -> dict[str, list[Shape | Arrow | SmartArt]]:
        if mode != "libreoffice":
            raise ValueError("LibreOfficeRichBackend only supports libreoffice mode.")
        self._ensure_runtime()
        drawings = read_sheet_drawings(self.file_path)
        shape_data: ShapeData = {}
        for sheet_name, drawing in drawings.items():
            emitted: list[Shape | Arrow | SmartArt] = []
            drawing_to_shape_id: dict[int, int] = {}
            shape_boxes: dict[int, _ShapeBox] = {}
            next_shape_id = 0
            for shape_info in drawing.shapes:
                next_shape_id += 1
                shape_id = next_shape_id
                drawing_to_shape_id[shape_info.ref.drawing_id] = shape_id
                box = _to_shape_box(shape_info, shape_id=shape_id)
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
            for connector_info in drawing.connectors:
                begin_id, end_id, approximation_level, confidence = _resolve_connector(
                    connector_info,
                    drawing_to_shape_id=drawing_to_shape_id,
                    shape_boxes=shape_boxes,
                )
                direction = _resolve_direction(connector_info)
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
                        direction=direction,
                        provenance="libreoffice_uno",
                        approximation_level=approximation_level,
                        confidence=confidence,
                    )
                )
            shape_data[sheet_name] = emitted
        return shape_data

    def extract_charts(self, *, mode: str) -> dict[str, list[Chart]]:
        if mode != "libreoffice":
            raise ValueError("LibreOfficeRichBackend only supports libreoffice mode.")
        self._ensure_runtime()
        drawings = read_sheet_drawings(self.file_path)
        chart_data: ChartData = {}
        for sheet_name, drawing in drawings.items():
            charts: list[Chart] = []
            for chart_info in drawing.charts:
                charts.append(
                    Chart(
                        name=chart_info.name,
                        chart_type=chart_info.chart_type,
                        title=chart_info.title,
                        y_axis_title=chart_info.y_axis_title,
                        y_axis_range=chart_info.y_axis_range,
                        w=chart_info.anchor_width,
                        h=chart_info.anchor_height,
                        series=chart_info.series,
                        l=chart_info.anchor_left or 0,
                        t=chart_info.anchor_top or 0,
                        provenance="libreoffice_uno",
                        approximation_level="partial",
                        confidence=0.5,
                    )
                )
            chart_data[sheet_name] = charts
        return chart_data

    def _ensure_runtime(self) -> None:
        if self._runtime_checked:
            return
        with self._session_factory() as session:
            workbook = session.load_workbook(self.file_path)
            session.close_workbook(workbook)
        self._runtime_checked = True


def _resolve_connector(
    connector_info: OoxmlConnectorInfo,
    *,
    drawing_to_shape_id: dict[int, int],
    shape_boxes: dict[int, _ShapeBox],
) -> tuple[int | None, int | None, str, float]:
    start_id = connector_info.connection.start_drawing_id
    end_id = connector_info.connection.end_drawing_id
    begin_id = drawing_to_shape_id.get(start_id) if start_id is not None else None
    end_id_resolved = drawing_to_shape_id.get(end_id) if end_id is not None else None
    if begin_id is not None or end_id_resolved is not None:
        return (begin_id, end_id_resolved, "direct", 1.0)

    start_point, end_point = _connector_endpoints(connector_info)
    heuristic_begin = _nearest_shape_id(start_point, shape_boxes)
    heuristic_end = _nearest_shape_id(end_point, shape_boxes)
    return (heuristic_begin, heuristic_end, "heuristic", 0.6)


def _resolve_direction(
    connector_info: OoxmlConnectorInfo,
) -> str | None:
    dx = connector_info.direction_dx
    dy = connector_info.direction_dy
    if dx is None or dy is None:
        return None
    angle = compute_line_angle_deg(float(dx), float(dy))
    return angle_to_compass(angle)


def _connector_endpoints(
    connector_info: OoxmlConnectorInfo,
) -> tuple[tuple[float, float] | None, tuple[float, float] | None]:
    left = connector_info.ref.left
    top = connector_info.ref.top
    dx = connector_info.direction_dx
    dy = connector_info.direction_dy
    if left is None or top is None or dx is None or dy is None:
        return (None, None)
    start = (float(left), float(top))
    end = (float(left + dx), float(top + dy))
    return (start, end)


def _nearest_shape_id(
    point: tuple[float, float] | None, shape_boxes: dict[int, _ShapeBox]
) -> int | None:
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
    dx = max(box.left - x, 0.0, x - box.right)
    dy = max(box.top - y, 0.0, y - box.bottom)
    return math.hypot(dx, dy)


def _to_shape_box(shape_info: OoxmlShapeInfo, *, shape_id: int) -> _ShapeBox | None:
    left = shape_info.ref.left
    top = shape_info.ref.top
    width = shape_info.ref.width
    height = shape_info.ref.height
    if left is None or top is None or width is None or height is None:
        return None
    return _ShapeBox(
        shape_id=shape_id,
        left=float(left),
        top=float(top),
        right=float(left + width),
        bottom=float(top + height),
    )


class _ShapeBox:
    def __init__(
        self, *, shape_id: int, left: float, top: float, right: float, bottom: float
    ) -> None:
        self.shape_id = shape_id
        self.left = left
        self.top = top
        self.right = right
        self.bottom = bottom
