from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Literal
from xml.etree import ElementTree
from zipfile import ZipFile

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


@dataclass(frozen=True)
class DrawingShapeRef:
    drawing_id: int
    name: str
    kind: Literal["shape", "connector", "chart"]
    left: int | None
    top: int | None
    width: int | None
    height: int | None


@dataclass(frozen=True)
class DrawingConnectorRef:
    drawing_id: int
    start_drawing_id: int | None
    end_drawing_id: int | None


@dataclass(frozen=True)
class OoxmlChartInfo:
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
    shapes: list[OoxmlShapeInfo] = field(default_factory=list)
    connectors: list[OoxmlConnectorInfo] = field(default_factory=list)
    charts: list[OoxmlChartInfo] = field(default_factory=list)


def read_sheet_drawings(file_path: Path) -> dict[str, SheetDrawingData]:
    """Read worksheet drawing metadata directly from OOXML parts."""
    result: dict[str, SheetDrawingData] = {}
    with ZipFile(file_path) as archive:
        for sheet_name, sheet_xml_path in _iter_sheet_xml_paths(archive):
            drawing_path = _resolve_sheet_drawing_path(archive, sheet_xml_path)
            if drawing_path is None:
                continue
            result[sheet_name] = _parse_sheet_drawing(archive, drawing_path)
    return result


def _iter_sheet_xml_paths(archive: ZipFile) -> list[tuple[str, str]]:
    workbook_xml = archive.read("xl/workbook.xml")
    workbook_root = ElementTree.fromstring(workbook_xml)
    rel_map = _read_relationships(archive, "xl/_rels/workbook.xml.rels")
    paths: list[tuple[str, str]] = []
    for sheet in workbook_root.findall("spreadsheetml:sheets/spreadsheetml:sheet", _NS):
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{_NS['r']}}}id")
        if not name or not rel_id or rel_id not in rel_map:
            continue
        paths.append((name, rel_map[rel_id]))
    return paths


def _resolve_sheet_drawing_path(archive: ZipFile, sheet_xml_path: str) -> str | None:
    rels_path = _rels_path(sheet_xml_path)
    if rels_path not in archive.namelist():
        return None
    rel_map = _read_relationships(archive, rels_path)
    for rel_id, target in rel_map.items():
        _ = rel_id
        if target.startswith("xl/drawings/"):
            return target
    return None


def _parse_sheet_drawing(archive: ZipFile, drawing_path: str) -> SheetDrawingData:
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
            shape_info = _parse_shape_node(shape_node)
            if shape_info is not None:
                shapes.append(shape_info)
            continue
        if (connector_node := anchor.find("xdr:cxnSp", _NS)) is not None:
            connector_info = _parse_connector_node(connector_node)
            if connector_info is not None:
                connectors.append(connector_info)
            continue
        if (graphic_frame := anchor.find("xdr:graphicFrame", _NS)) is not None:
            chart_info = _parse_chart_node(archive, graphic_frame, rel_map)
            if chart_info is not None:
                charts.append(chart_info)
    return SheetDrawingData(shapes=shapes, connectors=connectors, charts=charts)


def _parse_shape_node(node: ElementTree.Element) -> OoxmlShapeInfo | None:
    c_nv_pr = node.find("xdr:nvSpPr/xdr:cNvPr", _NS)
    if c_nv_pr is None:
        return None
    drawing_id = _int_attr(c_nv_pr, "id")
    name = c_nv_pr.attrib.get("name", f"Shape {drawing_id or 0}")
    left, top, width, height, rotation, flip_h, flip_v = _parse_sp_geometry(
        node.find("xdr:spPr", _NS)
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


def _parse_connector_node(node: ElementTree.Element) -> OoxmlConnectorInfo | None:
    c_nv_pr = node.find("xdr:nvCxnSpPr/xdr:cNvPr", _NS)
    if c_nv_pr is None:
        return None
    drawing_id = _int_attr(c_nv_pr, "id")
    name = c_nv_pr.attrib.get("name", f"Connector {drawing_id or 0}")
    left, top, width, height, rotation, flip_h, flip_v = _parse_sp_geometry(
        node.find("xdr:spPr", _NS)
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
        if line is not None and line.find("a:tailEnd", _NS) is not None
        else None,
        end_arrow_style=2
        if line is not None and line.find("a:headEnd", _NS) is not None
        else None,
    )


def _parse_chart_node(
    archive: ZipFile, node: ElementTree.Element, rel_map: dict[str, str]
) -> OoxmlChartInfo | None:
    c_nv_pr = node.find("xdr:nvGraphicFramePr/xdr:cNvPr", _NS)
    if c_nv_pr is None:
        return None
    name = c_nv_pr.attrib.get("name", "Chart")
    rel_id = node.find("a:graphic/a:graphicData/c:chart", _NS)
    if rel_id is None:
        return None
    target = rel_map.get(rel_id.attrib.get(f"{{{_NS['r']}}}id", ""))
    if target is None or target not in archive.namelist():
        return None
    chart_root = ElementTree.fromstring(archive.read(target))
    left, top, width, height, _rotation, _flip_h, _flip_v = _parse_xfrm_geometry(
        node.find("xdr:xfrm", _NS)
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
    title = chart_root.find("c:chart/c:title", _NS)
    return _extract_chart_text(title)


def _extract_y_axis_title(chart_root: ElementTree.Element) -> str:
    for axis in chart_root.findall(".//c:valAx", _NS):
        title = _extract_chart_text(axis.find("c:title", _NS))
        if title:
            return title
    return ""


def _extract_y_axis_range(chart_root: ElementTree.Element) -> list[float]:
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
    plot_area = chart_root.find("c:chart/c:plotArea", _NS)
    if plot_area is None:
        return []
    chart_node = None
    for child in plot_area:
        if _local_name(child.tag) in _CHART_TAGS:
            chart_node = child
            break
    if chart_node is None:
        return []

    series: list[ChartSeries] = []
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
        x_range = series_node.findtext(
            "c:cat/c:numRef/c:f", default=None, namespaces=_NS
        )
        if x_range is None:
            x_range = series_node.findtext(
                "c:cat/c:strRef/c:f", default=None, namespaces=_NS
            )
        y_range = series_node.findtext(
            "c:val/c:numRef/c:f", default=None, namespaces=_NS
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


def _format_shape_type(node: ElementTree.Element) -> str | None:
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
    if sp_pr is None:
        return (None, None, None, None, None, False, False)
    return _parse_xfrm_geometry(sp_pr.find("a:xfrm", _NS))


def _parse_xfrm_geometry(
    xfrm: ElementTree.Element | None,
) -> tuple[int | None, int | None, int | None, int | None, float | None, bool, bool]:
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


def _read_relationships(archive: ZipFile, rels_path: str) -> dict[str, str]:
    root = ElementTree.fromstring(archive.read(rels_path))
    base_path = _base_dir(_source_path_from_rels(rels_path))
    rel_map: dict[str, str] = {}
    for rel in root.findall("rel:Relationship", _NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rel_id or not target:
            continue
        rel_map[rel_id] = _normalize_zip_path(base_path, target)
    return rel_map


def _source_path_from_rels(rels_path: str) -> str:
    rels = PurePosixPath(rels_path)
    if rels.parent.name != "_rels":
        return rels_path
    stem = rels.name.removesuffix(".rels")
    return str(rels.parent.parent / stem)


def _rels_path(source_path: str) -> str:
    path = PurePosixPath(source_path)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def _base_dir(path: str) -> str:
    return str(PurePosixPath(path).parent)


def _normalize_zip_path(base_dir: str, target: str) -> str:
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
    if node is None:
        return ""
    texts = [text.text for text in node.findall(".//a:t", _NS) if text.text]
    return "".join(texts).strip()


def _emu_attr_to_points(node: ElementTree.Element | None, attr: str) -> int | None:
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
    if node is None:
        return None
    raw = node.attrib.get(attr)
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _float_attr(node: ElementTree.Element | None, attr: str) -> float | None:
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
    return tag.rsplit("}", 1)[-1]
