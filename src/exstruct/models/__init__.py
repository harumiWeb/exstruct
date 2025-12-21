from __future__ import annotations

from collections.abc import Generator
import json
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field


class Shape(BaseModel):
    """Shape metadata (position, size, text, and styling)."""

    id: int | None = Field(
        default=None, description="Sequential shape id within the sheet (if applicable)."
    )
    text: str = Field(description="Visible text content of the shape.")
    l: int = Field(description="Left offset (Excel units).")  # noqa: E741
    t: int = Field(description="Top offset (Excel units).")
    w: int | None = Field(default=None, description="Shape width (None if unknown).")
    h: int | None = Field(default=None, description="Shape height (None if unknown).")
    type: str | None = Field(default=None, description="Excel shape type name.")
    rotation: float | None = Field(
        default=None, description="Rotation angle in degrees."
    )
    begin_arrow_style: int | None = Field(
        default=None, description="Arrow style enum for the start of a connector."
    )
    end_arrow_style: int | None = Field(
        default=None, description="Arrow style enum for the end of a connector."
    )
    begin_id: int | None = Field(
        default=None,
        description=(
            "Shape id at the start of a connector (ConnectorFormat.BeginConnectedShape)."
        ),
    )
    end_id: int | None = Field(
        default=None,
        description=(
            "Shape id at the end of a connector (ConnectorFormat.EndConnectedShape)."
        ),
    )
    direction: Literal["E", "SE", "S", "SW", "W", "NW", "N", "NE"] | None = Field(
        default=None, description="Connector direction (compass heading)."
    )


class CellRow(BaseModel):
    """A single row of cells with optional hyperlinks."""

    r: int = Field(description="Row index (1-based).")
    c: dict[str, int | float | str] = Field(
        description="Column index (string) to cell value map."
    )
    links: dict[str, str] | None = Field(
        default=None, description="Optional hyperlinks per column index."
    )


class ChartSeries(BaseModel):
    """Series metadata for a chart."""

    name: str = Field(description="Series display name.")
    name_range: str | None = Field(
        default=None, description="Range reference for the series name."
    )
    x_range: str | None = Field(
        default=None, description="Range reference for X axis values."
    )
    y_range: str | None = Field(
        default=None, description="Range reference for Y axis values."
    )


class Chart(BaseModel):
    """Chart metadata including series and layout."""

    name: str = Field(description="Chart name.")
    chart_type: str = Field(description="Chart type (e.g., Column, Line).")
    title: str | None = Field(default=None, description="Chart title.")
    y_axis_title: str = Field(description="Y-axis title.")
    y_axis_range: list[float] = Field(
        default_factory=list, description="Y-axis range [min, max] when available."
    )
    w: int | None = Field(default=None, description="Chart width (None if unknown).")
    h: int | None = Field(default=None, description="Chart height (None if unknown).")
    series: list[ChartSeries] = Field(description="Series included in the chart.")
    l: int = Field(description="Left offset (Excel units).")  # noqa: E741
    t: int = Field(description="Top offset (Excel units).")
    error: str | None = Field(
        default=None, description="Extraction error detail if any."
    )


class PrintArea(BaseModel):
    """Cell coordinate bounds for a print area."""

    r1: int = Field(description="Start row (1-based).")
    c1: int = Field(description="Start column (1-based).")
    r2: int = Field(description="End row (1-based, inclusive).")
    c2: int = Field(description="End column (1-based, inclusive).")


class SheetData(BaseModel):
    """Structured data for a single sheet."""

    rows: list[CellRow] = Field(
        default_factory=list, description="Extracted rows with cell values and links."
    )
    shapes: list[Shape] = Field(
        default_factory=list, description="Shapes detected on the sheet."
    )
    charts: list[Chart] = Field(
        default_factory=list, description="Charts detected on the sheet."
    )
    table_candidates: list[str] = Field(
        default_factory=list, description="Cell ranges likely representing tables."
    )
    print_areas: list[PrintArea] = Field(
        default_factory=list, description="User-defined print areas."
    )
    auto_print_areas: list[PrintArea] = Field(
        default_factory=list, description="COM-computed auto page-break areas."
    )
    colors_map: dict[str, list[tuple[int, int]]] = Field(
        default_factory=dict,
        description="Mapping of hex color codes to lists of (row, column) tuples where the background color matches.",
    )

    def _as_payload(self) -> dict[str, object]:
        from ..io import dict_without_empty_values

        return dict_without_empty_values(self.model_dump(exclude_none=True))  # type: ignore

    def to_json(self, *, pretty: bool = False, indent: int | None = None) -> str:
        """
        Serialize the sheet into JSON text.
        """
        indent_val = 2 if pretty and indent is None else indent
        return json.dumps(self._as_payload(), ensure_ascii=False, indent=indent_val)

    def to_yaml(self) -> str:
        """
        Serialize the sheet into YAML text (requires pyyaml).
        """
        from ..io import _require_yaml

        yaml = _require_yaml()
        return str(
            yaml.safe_dump(
                self._as_payload(),
                allow_unicode=True,
                sort_keys=False,
                indent=2,
            )
        )

    def to_toon(self) -> str:
        """
        Serialize the sheet into TOON text (requires python-toon).
        """
        from ..io import _require_toon

        toon = _require_toon()
        return str(toon.encode(self._as_payload()))

    def save(
        self, path: str | Path, *, pretty: bool = False, indent: int | None = None
    ) -> Path:
        """
        Save the sheet to a file, inferring format from the extension.

        - .json → JSON
        - .yaml/.yml → YAML
        - .toon → TOON
        """
        dest = Path(path)
        fmt = (dest.suffix.lstrip(".") or "json").lower()
        match fmt:
            case "json":
                dest.write_text(
                    self.to_json(pretty=pretty, indent=indent), encoding="utf-8"
                )
            case "yaml" | "yml":
                dest.write_text(self.to_yaml(), encoding="utf-8")
            case "toon":
                dest.write_text(self.to_toon(), encoding="utf-8")
            case _:
                raise ValueError(f"Unsupported export format: {fmt}")
        return dest


class WorkbookData(BaseModel):
    """Workbook-level container with per-sheet data."""

    book_name: str = Field(description="Workbook file name (no path).")
    sheets: dict[str, SheetData] = Field(
        description="Mapping of sheet name to SheetData."
    )

    def to_json(self, *, pretty: bool = False, indent: int | None = None) -> str:
        """
        Serialize the workbook into JSON text.
        """
        from ..io import serialize_workbook

        return serialize_workbook(self, fmt="json", pretty=pretty, indent=indent)

    def to_yaml(self) -> str:
        """
        Serialize the workbook into YAML text (requires pyyaml).
        """
        from ..io import serialize_workbook

        return serialize_workbook(self, fmt="yaml")

    def to_toon(self) -> str:
        """
        Serialize the workbook into TOON text (requires python-toon).
        """
        from ..io import serialize_workbook

        return serialize_workbook(self, fmt="toon")

    def save(
        self, path: str | Path, *, pretty: bool = False, indent: int | None = None
    ) -> Path:
        """
        Save the workbook to a file, inferring format from the extension.

        - .json → JSON
        - .yaml/.yml → YAML
        - .toon → TOON
        """
        from ..io import save_as_json, save_as_toon, save_as_yaml

        dest = Path(path)
        fmt = (dest.suffix.lstrip(".") or "json").lower()
        match fmt:
            case "json":
                save_as_json(self, dest, pretty=pretty, indent=indent)
            case "yaml" | "yml":
                save_as_yaml(self, dest)
            case "toon":
                save_as_toon(self, dest)
            case _:
                raise ValueError(f"Unsupported export format: {fmt}")
        return dest

    def __getitem__(self, sheet_name: str) -> SheetData:
        """Return the SheetData for the given sheet name."""
        return self.sheets[sheet_name]

    def __iter__(self) -> Generator[tuple[str, SheetData], None, None]:
        """Iterate over (sheet_name, SheetData) pairs in order."""
        yield from self.sheets.items()


class PrintAreaView(BaseModel):
    """Slice of a sheet restricted to a print area (manual or auto)."""

    book_name: str = Field(description="Workbook name owning the area.")
    sheet_name: str = Field(description="Sheet name owning the area.")
    area: PrintArea = Field(description="Print area bounds.")
    shapes: list[Shape] = Field(
        default_factory=list, description="Shapes overlapping the area."
    )
    charts: list[Chart] = Field(
        default_factory=list, description="Charts overlapping the area."
    )
    rows: list[CellRow] = Field(
        default_factory=list, description="Rows within the area bounds."
    )
    table_candidates: list[str] = Field(
        default_factory=list, description="Table candidates intersecting the area."
    )

    def _as_payload(self) -> dict[str, object]:
        from ..io import dict_without_empty_values

        return dict_without_empty_values(self.model_dump(exclude_none=True))  # type: ignore

    def to_json(self, *, pretty: bool = False, indent: int | None = None) -> str:
        """
        Serialize the print-area view into JSON text.
        """
        indent_val = 2 if pretty and indent is None else indent
        return json.dumps(self._as_payload(), ensure_ascii=False, indent=indent_val)

    def to_yaml(self) -> str:
        """
        Serialize the print-area view into YAML text (requires pyyaml).
        """
        from ..io import _require_yaml

        yaml = _require_yaml()
        return str(
            yaml.safe_dump(
                self._as_payload(),
                allow_unicode=True,
                sort_keys=False,
                indent=2,
            )
        )

    def to_toon(self) -> str:
        """
        Serialize the print-area view into TOON text (requires python-toon).
        """
        from ..io import _require_toon

        toon = _require_toon()
        return str(toon.encode(self._as_payload()))

    def save(
        self, path: str | Path, *, pretty: bool = False, indent: int | None = None
    ) -> Path:
        """
        Save the print-area view to a file, inferring format from the extension.

        - .json JSON
        - .yaml/.yml YAML
        - .toon TOON
        """
        dest = Path(path)
        fmt = (dest.suffix.lstrip(".") or "json").lower()
        match fmt:
            case "json":
                dest.write_text(
                    self.to_json(pretty=pretty, indent=indent), encoding="utf-8"
                )
            case "yaml" | "yml":
                dest.write_text(self.to_yaml(), encoding="utf-8")
            case "toon":
                dest.write_text(self.to_toon(), encoding="utf-8")
            case _:
                raise ValueError(f"Unsupported export format: {fmt}")
        return dest
