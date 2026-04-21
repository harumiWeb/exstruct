"""Pure-Python OOXML rich extraction backend for non-COM workbook modes."""

from __future__ import annotations

import logging
from pathlib import Path
from zipfile import BadZipFile

from defusedxml import ElementTree

from ...models import Chart
from ..ooxml_drawing import SheetDrawingData, read_sheet_drawings
from .base import ChartData, RichBackend, ShapeData
from .libreoffice_backend import _build_shapes_from_ooxml

logger = logging.getLogger(__name__)

_OOXML_SUFFIXES = {".xlsx", ".xlsm"}


class OoxmlRichBackend(RichBackend):
    """Best-effort rich extraction backed only by OOXML workbook parts."""

    def __init__(self, file_path: Path) -> None:
        """Store the workbook path for lazy OOXML parsing."""

        self.file_path = file_path
        self._drawings: dict[str, SheetDrawingData] | None = None

    def extract_shapes(self, *, mode: str) -> ShapeData:
        """Extract shapes and connectors from OOXML drawing parts."""

        if mode != "light":
            raise ValueError("OoxmlRichBackend only supports light mode.")
        shape_data: ShapeData = {}
        for sheet_name, drawing in self._read_drawings().items():
            shape_data[sheet_name] = _build_shapes_from_ooxml(
                drawing.shapes,
                drawing.connectors,
                provenance="python_ooxml",
            )
        return shape_data

    def extract_charts(self, *, mode: str) -> ChartData:
        """Extract charts from OOXML drawing and chart parts."""

        if mode != "light":
            raise ValueError("OoxmlRichBackend only supports light mode.")
        chart_data: ChartData = {}
        for sheet_name, drawing in self._read_drawings().items():
            chart_data[sheet_name] = [
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
                    provenance="python_ooxml",
                    approximation_level="partial",
                    confidence=0.6,
                )
                for chart_info in drawing.charts
            ]
        return chart_data

    def _read_drawings(self) -> dict[str, SheetDrawingData]:
        """Read drawing parts once and degrade to an empty result on parse issues."""

        if self._drawings is not None:
            return self._drawings
        if self.file_path.suffix.lower() not in _OOXML_SUFFIXES:
            self._drawings = {}
            return self._drawings
        try:
            self._drawings = read_sheet_drawings(self.file_path)
        except (
            BadZipFile,
            ElementTree.ParseError,
            FileNotFoundError,
            KeyError,
            OSError,
            ValueError,
        ) as exc:
            logger.warning(
                "Failed to read OOXML drawing metadata from %s. (%r)",
                self.file_path,
                exc,
            )
            self._drawings = {}
        return self._drawings
