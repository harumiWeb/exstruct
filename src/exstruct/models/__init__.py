from __future__ import annotations

from collections.abc import Generator
import json
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field


class Shape(BaseModel):
    text: str
    l: int  # noqa: E741
    t: int
    w: int | None
    h: int | None
    type: str | None = None
    rotation: float | None = None
    begin_arrow_style: int | None = None
    end_arrow_style: int | None = None
    direction: Literal["E", "SE", "S", "SW", "W", "NW", "N", "NE"] | None = None


class CellRow(BaseModel):
    r: int
    c: dict[str, int | float | str]
    links: dict[str, str] | None = None


class ChartSeries(BaseModel):
    name: str
    name_range: str | None = None
    x_range: str | None = None
    y_range: str | None = None


class Chart(BaseModel):
    name: str
    chart_type: str
    title: str | None
    y_axis_title: str
    y_axis_range: list[float] = Field(default_factory=list)
    w: int | None = None
    h: int | None = None
    series: list[ChartSeries]
    l: int  # noqa: E741
    t: int
    error: str | None = None


class PrintArea(BaseModel):
    r1: int
    c1: int
    r2: int
    c2: int


class SheetData(BaseModel):
    rows: list[CellRow] = Field(default_factory=list)
    shapes: list[Shape] = Field(default_factory=list)
    charts: list[Chart] = Field(default_factory=list)
    table_candidates: list[str] = Field(default_factory=list)
    print_areas: list[PrintArea] = Field(default_factory=list)

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
    book_name: str
    sheets: dict[str, SheetData]

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
    book_name: str
    sheet_name: str
    area: PrintArea
    shapes: list[Shape] = Field(default_factory=list)
    charts: list[Chart] = Field(default_factory=list)
    rows: list[CellRow] = Field(default_factory=list)
    table_candidates: list[str] = Field(default_factory=list)

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
