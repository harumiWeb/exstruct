"""OOXML drawing parsers for shapes, connectors, and charts."""

from __future__ import annotations

from dataclasses import dataclass, field
import logging
from pathlib import Path, PurePosixPath
from typing import Literal
from zipfile import ZipFile

from defusedxml import ElementTree

from ..models import ChartSeries

_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "spreadsheetml": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}
_EMU_PER_POINT = 12700.0
_DEFAULT_COLUMN_WIDTH_POINTS = 48.0
_DEFAULT_ROW_HEIGHT_POINTS = 15.0
_CHART_TAGS = {
    "areaChart",
    "barChart",
    "bubbleChart",
    "doughnutChart",
    "lineChart",
    "ofPieChart",
    "pieChart",
    "radarChart",
    "scatterChart",
    "stockChart",
    "surfaceChart",
}
_SHAPE_TYPE_MAP = {
    "ellipse": "Oval",
    "flowChartDecision": "FlowchartDecision",
    "flowChartProcess": "FlowchartProcess",
    "rect": "Rectangle",
    "straightConnector1": "StraightConnector1",
}
_WORKSHEET_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)
_DRAWING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
_CHART_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class DrawingShapeRef:
    """Geometric and identity metadata for a drawing object anchor."""

    drawing_id: int
    name: str
    kind: Literal["shape", "connector", "chart"]
    left: int | None
    top: int | None
    width: int | None
    height: int | None


@dataclass(frozen=True)
class DrawingConnectorRef:
    """Connection metadata linking a connector to drawing ids."""

    drawing_id: int
    start_drawing_id: int | None
    end_drawing_id: int | None


@dataclass(frozen=True)
class OoxmlRelationship:
    """Relationship metadata extracted from an OOXML ``.rels`` part."""

    target: str
    relationship_type: str


@dataclass(frozen=True)
class OoxmlChartInfo:
    """Chart metadata extracted from OOXML chart and drawing parts."""

    name: str
    chart_type: str
    title: str | None
    y_axis_title: str
    y_axis_range: list[float]
    series: list[ChartSeries]
    anchor_left: int | None
    anchor_top: int | None
    anchor_width: int | None
    anchor_height: int | None


@dataclass(frozen=True)
class OoxmlShapeInfo:
    """Shape metadata extracted from OOXML drawing anchors."""

    ref: DrawingShapeRef
    text: str = ""
    shape_type: str | None = None
    rotation: float | None = None
    direction_dx: int | None = None
    direction_dy: int | None = None
    begin_arrow_style: int | None = None
    end_arrow_style: int | None = None


@dataclass(frozen=True)
class OoxmlConnectorInfo:
    """Connector metadata extracted from OOXML drawing anchors."""

    ref: DrawingShapeRef
    connection: DrawingConnectorRef
    text: str = ""
    rotation: float | None = None
    direction_dx: int | None = None
    direction_dy: int | None = None
    begin_arrow_style: int | None = None
    end_arrow_style: int | None = None


@dataclass(frozen=True)
class SheetDrawingData:
    """Grouped OOXML drawing artifacts for a worksheet."""

    shapes: list[OoxmlShapeInfo] = field(default_factory=list)
    connectors: list[OoxmlConnectorInfo] = field(default_factory=list)
    charts: list[OoxmlChartInfo] = field(default_factory=list)


@dataclass
class SheetDrawingMetrics:
    """Worksheet row/column sizing used to resolve drawing anchor coordinates."""

    default_column_width_points: float = _DEFAULT_COLUMN_WIDTH_POINTS
    default_row_height_points: float = _DEFAULT_ROW_HEIGHT_POINTS
    column_width_points: dict[int, float] = field(default_factory=dict)
    row_height_points: dict[int, float] = field(default_factory=dict)

    def column_offset_points(self, col_index: int) -> float:
        """Return the horizontal offset before a zero-based column index."""

        offset = 0.0
        for index in range(col_index):
            offset += self.column_width_points.get(
                index,
                self.default_column_width_points,
            )
        return offset

    def row_offset_points(self, row_index: int) -> float:
        """Return the vertical offset before a zero-based row index."""

        offset = 0.0
        for index in range(row_index):
            offset += self.row_height_points.get(
                index,
                self.default_row_height_points,
            )
        return offset


def read_sheet_drawings(file_path: Path) -> dict[str, SheetDrawingData]:
    """Read worksheet drawing metadata directly from OOXML parts."""
    result: dict[str, SheetDrawingData] = {}
    with ZipFile(file_path) as archive:
        for sheet_name, sheet_xml_path in _iter_sheet_xml_paths(archive):
            try:
                drawing_path = _resolve_sheet_drawing_path(archive, sheet_xml_path)
                if drawing_path is None:
                    continue
                result[sheet_name] = _parse_sheet_drawing(
                    archive,
                    drawing_path,
                    _read_sheet_metrics(archive, sheet_xml_path),
                )
            except (
                ElementTree.ParseError,
                FileNotFoundError,
                KeyError,
                OSError,
                ValueError,
            ) as exc:
                logger.warning(
                    "Skipping OOXML drawing metadata for sheet %s in %s. (%r)",
                    sheet_name,
                    file_path,
                    exc,
                )
    return result


def _iter_sheet_xml_paths(archive: ZipFile) -> list[tuple[str, str]]:
    """Return workbook sheet names paired with their OOXML worksheet paths."""

    workbook_xml = archive.read("xl/workbook.xml")
    workbook_root = ElementTree.fromstring(workbook_xml)
    rel_map = _read_relationships(archive, "xl/_rels/workbook.xml.rels")
    paths: list[tuple[str, str]] = []
    for sheet in workbook_root.findall("spreadsheetml:sheets/spreadsheetml:sheet", _NS):
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{_NS['r']}}}id")
        if not name or not rel_id or rel_id not in rel_map:
            continue
        relationship = rel_map[rel_id]
        if relationship.relationship_type != _WORKSHEET_REL_TYPE:
            continue
        paths.append((name, relationship.target))
    return paths


def _resolve_sheet_drawing_path(archive: ZipFile, sheet_xml_path: str) -> str | None:
    """Resolve the drawing part referenced by a worksheet, if any."""

    rels_path = _rels_path(sheet_xml_path)
    if rels_path not in archive.namelist():
        return None
    rel_map = _read_relationships(archive, rels_path)
    for relationship in rel_map.values():
        if relationship.relationship_type != _DRAWING_REL_TYPE:
            continue
        return relationship.target
    return None


def _parse_sheet_drawing(
    archive: ZipFile,
    drawing_path: str,
    sheet_metrics: SheetDrawingMetrics,
) -> SheetDrawingData:
    """Parse shapes, connectors, and charts from a drawing part."""

    root = ElementTree.fromstring(archive.read(drawing_path))
    rel_map = {}
    drawing_rels_path = _rels_path(drawing_path)
    if drawing_rels_path in archive.namelist():
        rel_map = _read_relationships(archive, drawing_rels_path)

    shapes: list[OoxmlShapeInfo] = []
    connectors: list[OoxmlConnectorInfo] = []
    charts: list[OoxmlChartInfo] = []
    for anchor in root:
        if _local_name(anchor.tag) not in {
            "absoluteAnchor",
            "oneCellAnchor",
            "twoCellAnchor",
        }:
            continue
        if (shape_node := anchor.find("xdr:sp", _NS)) is not None:
            shape_info = _parse_shape_node(anchor, shape_node, sheet_metrics)
            if shape_info is not None:
                shapes.append(shape_info)
            continue
        if (connector_node := anchor.find("xdr:cxnSp", _NS)) is not None:
            connector_info = _parse_connector_node(
                anchor, connector_node, sheet_metrics
            )
            if connector_info is not None:
                connectors.append(connector_info)
            continue
        if (graphic_frame := anchor.find("xdr:graphicFrame", _NS)) is not None:
            chart_info = _parse_chart_node(
                archive,
                anchor,
                graphic_frame,
                rel_map,
                sheet_metrics,
            )
            if chart_info is not None:
                charts.append(chart_info)
    return SheetDrawingData(shapes=shapes, connectors=connectors, charts=charts)


def _parse_shape_node(
    anchor: ElementTree.Element,
    node: ElementTree.Element,
    sheet_metrics: SheetDrawingMetrics,
) -> OoxmlShapeInfo | None:
    """Parse an OOXML shape node into an ``OoxmlShapeInfo`` record."""

    c_nv_pr = node.find("xdr:nvSpPr/xdr:cNvPr", _NS)
    if c_nv_pr is None:
        return None
    drawing_id = _int_attr(c_nv_pr, "id")
    name = c_nv_pr.attrib.get("name", f"Shape {drawing_id or 0}")
    left, top, width, height, rotation, flip_h, flip_v = _parse_sp_geometry(
        node.find("xdr:spPr", _NS)
    )
    left, top, width, height = _merge_anchor_geometry(
        anchor,
        left=left,
        top=top,
        width=width,
        height=height,
        sheet_metrics=sheet_metrics,
        prefer_transform_position_when_sized=True,
    )
    ref = DrawingShapeRef(
        drawing_id=drawing_id or 0,
        name=name,
        kind="shape",
        left=left,
        top=top,
        width=width,
        height=height,
    )
    dx = None if width is None else (-width if flip_h else width)
    dy = None if height is None else (-height if flip_v else height)
    return OoxmlShapeInfo(
        ref=ref,
        text=_extract_text(node.find("xdr:txBody", _NS)),
        shape_type=_format_shape_type(node),
        rotation=rotation,
        direction_dx=dx,
        direction_dy=dy,
    )


def _parse_connector_node(
    anchor: ElementTree.Element,
    node: ElementTree.Element,
    sheet_metrics: SheetDrawingMetrics,
) -> OoxmlConnectorInfo | None:
    """Parse an OOXML connector node into an ``OoxmlConnectorInfo`` record."""

    c_nv_pr = node.find("xdr:nvCxnSpPr/xdr:cNvPr", _NS)
    if c_nv_pr is None:
        return None
    drawing_id = _int_attr(c_nv_pr, "id")
    name = c_nv_pr.attrib.get("name", f"Connector {drawing_id or 0}")
    left, top, width, height, rotation, flip_h, flip_v = _parse_sp_geometry(
        node.find("xdr:spPr", _NS)
    )
    left, top, width, height = _merge_anchor_geometry(
        anchor,
        left=left,
        top=top,
        width=width,
        height=height,
        sheet_metrics=sheet_metrics,
        prefer_transform_position_when_sized=True,
    )
    ref = DrawingShapeRef(
        drawing_id=drawing_id or 0,
        name=name,
        kind="connector",
        left=left,
        top=top,
        width=width,
        height=height,
    )
    connector_props = node.find("xdr:nvCxnSpPr/xdr:cNvCxnSpPr", _NS)
    start_node = (
        connector_props.find("a:stCxn", _NS) if connector_props is not None else None
    )
    end_node = (
        connector_props.find("a:endCxn", _NS) if connector_props is not None else None
    )
    dx = None if width is None else (-width if flip_h else width)
    dy = None if height is None else (-height if flip_v else height)
    line = node.find("xdr:spPr/a:ln", _NS)
    return OoxmlConnectorInfo(
        ref=ref,
        connection=DrawingConnectorRef(
            drawing_id=drawing_id or 0,
            start_drawing_id=_int_attr(start_node, "id"),
            end_drawing_id=_int_attr(end_node, "id"),
        ),
        text="",
        rotation=rotation,
        direction_dx=dx,
        direction_dy=dy,
        begin_arrow_style=2
        if line is not None and line.find("a:headEnd", _NS) is not None
        else None,
        end_arrow_style=2
        if line is not None and line.find("a:tailEnd", _NS) is not None
        else None,
    )


def _parse_chart_node(
    archive: ZipFile,
    anchor: ElementTree.Element,
    node: ElementTree.Element,
    rel_map: dict[str, OoxmlRelationship],
    sheet_metrics: SheetDrawingMetrics,
) -> OoxmlChartInfo | None:
    """Parse an OOXML graphic-frame chart node into chart metadata."""

    c_nv_pr = node.find("xdr:nvGraphicFramePr/xdr:cNvPr", _NS)
    if c_nv_pr is None:
        return None
    name = c_nv_pr.attrib.get("name", "Chart")
    rel_id = node.find("a:graphic/a:graphicData/c:chart", _NS)
    if rel_id is None:
        return None
    relationship = rel_map.get(rel_id.attrib.get(f"{{{_NS['r']}}}id", ""))
    if relationship is None or relationship.relationship_type != _CHART_REL_TYPE:
        return None
    target = relationship.target
    if target not in archive.namelist():
        return None
    chart_root = ElementTree.fromstring(archive.read(target))
    left, top, width, height, _rotation, _flip_h, _flip_v = _parse_xfrm_geometry(
        node.find("xdr:xfrm", _NS)
    )
    left, top, width, height = _merge_anchor_geometry(
        anchor,
        left=left,
        top=top,
        width=width,
        height=height,
        sheet_metrics=sheet_metrics,
        prefer_transform_position_when_sized=True,
    )
    return OoxmlChartInfo(
        name=name,
        chart_type=_extract_chart_type(chart_root),
        title=_extract_chart_title(chart_root),
        y_axis_title=_extract_y_axis_title(chart_root),
        y_axis_range=_extract_y_axis_range(chart_root),
        series=_extract_chart_series(chart_root),
        anchor_left=left,
        anchor_top=top,
        anchor_width=width,
        anchor_height=height,
    )


def _extract_chart_type(chart_root: ElementTree.Element) -> str:
    """Extract the ExStruct chart type label from a chart part."""

    plot_area = chart_root.find("c:chart/c:plotArea", _NS)
    if plot_area is None:
        return "unknown"
    for child in plot_area:
        tag = _local_name(child.tag)
        if tag not in _CHART_TAGS:
            continue
        if tag == "barChart":
            bar_dir = child.find("c:barDir", _NS)
            if bar_dir is not None and bar_dir.attrib.get("val") == "bar":
                return "Bar"
            return "Column"
        return {
            "areaChart": "Area",
            "bubbleChart": "Bubble",
            "doughnutChart": "Doughnut",
            "lineChart": "Line",
            "ofPieChart": "OfPie",
            "pieChart": "Pie",
            "radarChart": "Radar",
            "scatterChart": "Scatter",
            "stockChart": "Stock",
            "surfaceChart": "Surface",
        }.get(tag, tag.removesuffix("Chart"))
    return "unknown"


def _extract_chart_title(chart_root: ElementTree.Element) -> str | None:
    """Extract a chart title from a chart part."""

    title = chart_root.find("c:chart/c:title", _NS)
    return _extract_chart_text(title)


def _extract_y_axis_title(chart_root: ElementTree.Element) -> str:
    """Extract the first value-axis title from a chart part."""

    for axis in chart_root.findall(".//c:valAx", _NS):
        title = _extract_chart_text(axis.find("c:title", _NS))
        if title:
            return title
    return ""


def _extract_y_axis_range(chart_root: ElementTree.Element) -> list[float]:
    """Extract the first explicit value-axis min/max range."""

    for axis in chart_root.findall(".//c:valAx", _NS):
        scaling = axis.find("c:scaling", _NS)
        if scaling is None:
            continue
        min_node = scaling.find("c:min", _NS)
        max_node = scaling.find("c:max", _NS)
        if min_node is None or max_node is None:
            continue
        min_value = _float_attr(min_node, "val")
        max_value = _float_attr(max_node, "val")
        if min_value is None or max_value is None:
            continue
        return [min_value, max_value]
    return []


def _extract_chart_series(chart_root: ElementTree.Element) -> list[ChartSeries]:
    """Extract chart series labels and ranges from a chart part."""

    plot_area = chart_root.find("c:chart/c:plotArea", _NS)
    if plot_area is None:
        return []
    series: list[ChartSeries] = []
    for chart_node in plot_area:
        if _local_name(chart_node.tag) not in _CHART_TAGS:
            continue
        for series_node in chart_node.findall("c:ser", _NS):
            name_range = series_node.findtext(
                "c:tx/c:strRef/c:f", default=None, namespaces=_NS
            )
            literal_name = series_node.findtext(
                "c:tx/c:strRef/c:strCache/c:pt/c:v",
                default=None,
                namespaces=_NS,
            )
            if literal_name is None:
                literal_name = series_node.findtext(
                    "c:tx/c:v", default=None, namespaces=_NS
                )
            x_range = _extract_series_range(
                series_node,
                "c:xVal/c:numRef/c:f",
                "c:xVal/c:strRef/c:f",
                "c:cat/c:numRef/c:f",
                "c:cat/c:strRef/c:f",
            )
            y_range = _extract_series_range(
                series_node,
                "c:yVal/c:numRef/c:f",
                "c:yVal/c:strRef/c:f",
                "c:val/c:numRef/c:f",
            )
            series.append(
                ChartSeries(
                    name=literal_name or name_range or "",
                    name_range=name_range,
                    x_range=x_range,
                    y_range=y_range,
                )
            )
    return series


def _extract_chart_text(node: ElementTree.Element | None) -> str | None:
    """Collect text content from chart title and rich text nodes."""

    if node is None:
        return None
    texts = [
        text_node.text.strip()
        for text_node in (node.findall(".//a:t", _NS) + node.findall(".//c:v", _NS))
        if text_node.text and text_node.text.strip()
    ]
    if not texts:
        return None
    return "".join(texts)


def _extract_series_range(
    node: ElementTree.Element,
    *paths: str,
) -> str | None:
    """Return the first matching series formula path."""

    for path in paths:
        value = node.findtext(path, default=None, namespaces=_NS)
        if isinstance(value, str):
            return value
    return None


def _format_shape_type(node: ElementTree.Element) -> str | None:
    """Map an OOXML preset geometry into an ExStruct shape type label."""

    sp_pr = node.find("xdr:spPr", _NS)
    if sp_pr is None:
        return None
    prst = sp_pr.find("a:prstGeom", _NS)
    if prst is None:
        return None
    raw = prst.attrib.get("prst")
    if not raw:
        return None
    label = _SHAPE_TYPE_MAP.get(raw, raw)
    c_nv_sp_pr = node.find("xdr:nvSpPr/xdr:cNvSpPr", _NS)
    is_text_box = c_nv_sp_pr is not None and c_nv_sp_pr.attrib.get("txBox") == "1"
    prefix = "TextBox" if is_text_box else "AutoShape"
    return f"{prefix}-{label}"


def _parse_sp_geometry(
    sp_pr: ElementTree.Element | None,
) -> tuple[int | None, int | None, int | None, int | None, float | None, bool, bool]:
    """Parse position, size, rotation, and flips from a shape properties node."""

    if sp_pr is None:
        return (None, None, None, None, None, False, False)
    return _parse_xfrm_geometry(sp_pr.find("a:xfrm", _NS))


def _parse_xfrm_geometry(
    xfrm: ElementTree.Element | None,
) -> tuple[int | None, int | None, int | None, int | None, float | None, bool, bool]:
    """Parse position, size, rotation, and flips from an ``xfrm`` node."""

    if xfrm is None:
        return (None, None, None, None, None, False, False)
    off = xfrm.find("a:off", _NS)
    ext = xfrm.find("a:ext", _NS)
    left = _emu_attr_to_points(off, "x")
    top = _emu_attr_to_points(off, "y")
    width = _emu_attr_to_points(ext, "cx")
    height = _emu_attr_to_points(ext, "cy")
    rotation = None
    raw_rotation = xfrm.attrib.get("rot")
    if raw_rotation is not None:
        try:
            rotation = float(raw_rotation) / 60000.0
        except ValueError:
            rotation = None
    flip_h = xfrm.attrib.get("flipH") == "1"
    flip_v = xfrm.attrib.get("flipV") == "1"
    return (left, top, width, height, rotation, flip_h, flip_v)


def _merge_anchor_geometry(
    anchor: ElementTree.Element,
    *,
    left: int | None,
    top: int | None,
    width: int | None,
    height: int | None,
    sheet_metrics: SheetDrawingMetrics | None = None,
    prefer_transform_position_when_sized: bool = False,
) -> tuple[int | None, int | None, int | None, int | None]:
    """Use parent anchors for placement and child transforms for size when present."""

    anchor_left, anchor_top, anchor_width, anchor_height = _parse_anchor_geometry(
        anchor,
        sheet_metrics,
    )
    if (
        prefer_transform_position_when_sized
        and width not in {None, 0}
        and height not in {None, 0}
    ):
        resolved_left = left if left is not None else anchor_left
        resolved_top = top if top is not None else anchor_top
    else:
        resolved_left = anchor_left if anchor_left is not None else left
        resolved_top = anchor_top if anchor_top is not None else top
    resolved_width = width if width not in {None, 0} else anchor_width
    resolved_height = height if height not in {None, 0} else anchor_height
    return (resolved_left, resolved_top, resolved_width, resolved_height)


def _parse_anchor_geometry(
    anchor: ElementTree.Element,
    sheet_metrics: SheetDrawingMetrics | None = None,
) -> tuple[int | None, int | None, int | None, int | None]:
    """Parse approximate placement from a parent drawing anchor."""

    tag = _local_name(anchor.tag)
    if tag == "absoluteAnchor":
        pos = anchor.find("xdr:pos", _NS)
        ext = anchor.find("xdr:ext", _NS)
        return (
            _emu_attr_to_points(pos, "x"),
            _emu_attr_to_points(pos, "y"),
            _emu_attr_to_points(ext, "cx"),
            _emu_attr_to_points(ext, "cy"),
        )
    if tag == "oneCellAnchor":
        marker = anchor.find("xdr:from", _NS)
        ext = anchor.find("xdr:ext", _NS)
        left, top = _marker_to_points(marker, sheet_metrics)
        return (
            left,
            top,
            _emu_attr_to_points(ext, "cx"),
            _emu_attr_to_points(ext, "cy"),
        )
    if tag == "twoCellAnchor":
        start = _marker_to_points(anchor.find("xdr:from", _NS), sheet_metrics)
        end = _marker_to_points(anchor.find("xdr:to", _NS), sheet_metrics)
        if start[0] is None or start[1] is None or end[0] is None or end[1] is None:
            return (None, None, None, None)
        return (
            start[0],
            start[1],
            max(end[0] - start[0], 0),
            max(end[1] - start[1], 0),
        )
    return (None, None, None, None)


def _marker_to_points(
    marker: ElementTree.Element | None,
    sheet_metrics: SheetDrawingMetrics | None = None,
) -> tuple[int | None, int | None]:
    """Convert an OOXML anchor marker to approximate point coordinates."""

    if marker is None:
        return (None, None)
    if sheet_metrics is None:
        sheet_metrics = SheetDrawingMetrics()
    col = _find_int_text(marker, "xdr:col")
    col_off = _find_int_text(marker, "xdr:colOff")
    row = _find_int_text(marker, "xdr:row")
    row_off = _find_int_text(marker, "xdr:rowOff")
    if col is None or row is None:
        return (None, None)
    left = int(
        round(sheet_metrics.column_offset_points(col) + (col_off or 0) / _EMU_PER_POINT)
    )
    top = int(
        round(sheet_metrics.row_offset_points(row) + (row_off or 0) / _EMU_PER_POINT)
    )
    return (left, top)


def _read_sheet_metrics(
    archive: ZipFile,
    sheet_xml_path: str,
) -> SheetDrawingMetrics:
    """Read worksheet row/column sizing from a sheet XML part."""

    sheet_root = ElementTree.fromstring(archive.read(sheet_xml_path))
    return _parse_sheet_metrics(sheet_root)


def _parse_sheet_metrics(sheet_root: ElementTree.Element) -> SheetDrawingMetrics:
    """Parse worksheet row heights and column widths into drawing metrics."""

    metrics = SheetDrawingMetrics()
    sheet_format = sheet_root.find("spreadsheetml:sheetFormatPr", _NS)
    if sheet_format is not None:
        default_row_height = _float_attr(sheet_format, "defaultRowHeight")
        if default_row_height is not None and default_row_height > 0:
            metrics.default_row_height_points = default_row_height
        default_column_width = _float_attr(sheet_format, "defaultColWidth")
        if default_column_width is not None and default_column_width > 0:
            metrics.default_column_width_points = _column_width_to_points(
                default_column_width
            )

    cols = sheet_root.find("spreadsheetml:cols", _NS)
    if cols is not None:
        for col in cols.findall("spreadsheetml:col", _NS):
            min_index = _int_attr(col, "min")
            max_index = _int_attr(col, "max")
            width = _float_attr(col, "width")
            if (
                min_index is None
                or max_index is None
                or width is None
                or min_index <= 0
                or max_index < min_index
                or width <= 0
            ):
                continue
            width_points = _column_width_to_points(width)
            for index in range(min_index - 1, max_index):
                metrics.column_width_points[index] = width_points

    for row in sheet_root.findall("spreadsheetml:sheetData/spreadsheetml:row", _NS):
        row_index = _int_attr(row, "r")
        height = _float_attr(row, "ht")
        if row_index is None or row_index <= 0 or height is None or height <= 0:
            continue
        metrics.row_height_points[row_index - 1] = height

    return metrics


def _column_width_to_points(width: float) -> float:
    """Approximate an OOXML column width attribute in point units."""

    if width <= 0:
        return 0.0
    return (width * 7.0 + 5.0) * 72.0 / 96.0


def _read_relationships(
    archive: ZipFile, rels_path: str
) -> dict[str, OoxmlRelationship]:
    """Read a relationships part into a relationship-id keyed metadata map."""

    root = ElementTree.fromstring(archive.read(rels_path))
    base_path = _base_dir(_source_path_from_rels(rels_path))
    rel_map: dict[str, OoxmlRelationship] = {}
    for rel in root.findall("rel:Relationship", _NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        relationship_type = rel.attrib.get("Type")
        if not rel_id or not target or not relationship_type:
            continue
        rel_map[rel_id] = OoxmlRelationship(
            target=_normalize_zip_path(base_path, target),
            relationship_type=relationship_type,
        )
    return rel_map


def _source_path_from_rels(rels_path: str) -> str:
    """Recover the source part path that owns a relationships part."""

    rels = PurePosixPath(rels_path)
    if rels.parent.name != "_rels":
        return rels_path
    stem = rels.name.removesuffix(".rels")
    return str(rels.parent.parent / stem)


def _rels_path(source_path: str) -> str:
    """Return the relationships part path for a source part."""

    path = PurePosixPath(source_path)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def _base_dir(path: str) -> str:
    """Return the POSIX parent directory for a zip path."""

    return str(PurePosixPath(path).parent)


def _normalize_zip_path(base_dir: str, target: str) -> str:
    """Normalize a relative OOXML zip target against a base directory."""

    base = PurePosixPath(base_dir)
    normalized = base.joinpath(PurePosixPath(target)).as_posix()
    parts: list[str] = []
    for part in normalized.split("/"):
        if part in {"", "."}:
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts)


def _extract_text(node: ElementTree.Element | None) -> str:
    """Extract concatenated text from a drawing text body."""

    if node is None:
        return ""
    texts = [text.text for text in node.findall(".//a:t", _NS) if text.text]
    return "".join(texts).strip()


def _emu_attr_to_points(node: ElementTree.Element | None, attr: str) -> int | None:
    """Convert an EMU-valued XML attribute to rounded point units."""

    if node is None:
        return None
    raw = node.attrib.get(attr)
    if raw is None:
        return None
    try:
        return int(round(int(raw) / _EMU_PER_POINT))
    except ValueError:
        return None


def _int_attr(node: ElementTree.Element | None, attr: str) -> int | None:
    """Parse an integer XML attribute when present."""

    if node is None:
        return None
    raw = node.attrib.get(attr)
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _find_int_text(node: ElementTree.Element | None, path: str) -> int | None:
    """Parse an integer child text value when present."""

    if node is None:
        return None
    raw = node.findtext(path, default=None, namespaces=_NS)
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _float_attr(node: ElementTree.Element | None, attr: str) -> float | None:
    """Parse a floating-point XML attribute when present."""

    if node is None:
        return None
    raw = node.attrib.get(attr)
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _local_name(tag: str) -> str:
    """Return the local XML name without its namespace."""

    return tag.rsplit("}", 1)[-1]
