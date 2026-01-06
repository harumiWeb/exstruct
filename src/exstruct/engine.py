from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, TextIO, TypedDict, cast

from pydantic import BaseModel, ConfigDict, Field

from .core import cells as _cells
from .core.cells import set_table_detection_params
from .core.integrate import extract_workbook
from .io import (
    save_auto_page_break_views,
    save_print_area_views,
    save_sheets,
    serialize_workbook,
)
from .models import SheetData, WorkbookData
from .render import export_pdf, export_sheet_images

ExtractionMode = Literal["light", "standard", "verbose"]


class TableParams(TypedDict, total=False):
    table_score_threshold: float
    density_min: float
    coverage_min: float
    min_nonempty_cells: int


class ColorsOptions(BaseModel):
    """Color extraction options.

    Examples:
        >>> ColorsOptions(
        ...     include_default_background=False,
        ...     ignore_colors=["#FFFFFF", "AD3815", "theme:1:0.2", "indexed:64", "auto"],
        ... )
    """

    include_default_background: bool = Field(
        default=False, description="Include default (white) backgrounds."
    )
    ignore_colors: list[str] = Field(
        default_factory=list, description="List of color keys to ignore."
    )

    def ignore_colors_set(self) -> set[str]:
        """Return ignore_colors as a set of normalized strings.

        Returns:
            Set of color keys to ignore.
        """
        return set(self.ignore_colors)


@dataclass(frozen=True)
class StructOptions:
    """
    Extraction-time options for ExStructEngine.

    Attributes:
        mode: Extraction mode. One of "light", "standard", "verbose".
              - light: cells + table candidates only (no COM, shapes/charts empty)
              - standard: texted shapes + arrows + charts (if COM available)
              - verbose: all shapes (width/height), charts, table candidates
        table_params: Optional dict passed to `set_table_detection_params(**table_params)`
                      before extraction. Use this to tweak table detection heuristics
                      per engine instance without touching global state.
        include_colors_map: Whether to extract background color maps.
        include_merged_cells: Whether to extract merged cell ranges.
        include_merged_values_in_rows: Whether to keep merged values in rows.
        colors: Color extraction options.
    """

    mode: ExtractionMode = "standard"
    table_params: TableParams | None = (
        None  # forwarded to set_table_detection_params if provided
    )
    include_cell_links: bool | None = None  # None -> auto: verbose=True, others=False
    include_colors_map: bool | None = None  # None -> auto: verbose=True, others=False
    include_merged_cells: bool | None = None  # None -> auto: light=False, others=True
    include_merged_values_in_rows: bool = True
    colors: ColorsOptions = field(default_factory=ColorsOptions)


class FormatOptions(BaseModel):
    """Formatting options for serialization."""

    model_config = ConfigDict(arbitrary_types_allowed=True)
    fmt: Literal["json", "yaml", "yml", "toon"] = Field(
        default="json", description="Serialization format."
    )
    pretty: bool = Field(default=False, description="Pretty-print JSON output.")
    indent: int | None = Field(
        default=None,
        description="Indent width for JSON (defaults to 2 when pretty is True).",
    )


class FilterOptions(BaseModel):
    """Include/exclude filters for output."""

    model_config = ConfigDict(arbitrary_types_allowed=True)
    include_rows: bool = Field(default=True, description="Include cell rows.")
    include_shapes: bool = Field(default=True, description="Include shapes.")
    include_shape_size: bool | None = Field(
        default=None,
        description="Include shape size; None -> auto (verbose=True, others=False).",
    )
    include_charts: bool = Field(default=True, description="Include charts.")
    include_chart_size: bool | None = Field(
        default=None,
        description="Include chart size; None -> auto (verbose=True, others=False).",
    )
    include_tables: bool = Field(
        default=True, description="Include table candidate ranges."
    )
    include_print_areas: bool | None = Field(
        default=None,
        description="Include print areas; None -> auto (light=False, others=True).",
    )
    include_auto_print_areas: bool = Field(
        default=False, description="Include COM-computed auto page-break areas."
    )
    include_merged_cells: bool = Field(
        default=True, description="Include merged cell ranges."
    )


class DestinationOptions(BaseModel):
    """Destinations for optional side outputs."""

    model_config = ConfigDict(arbitrary_types_allowed=True)
    sheets_dir: str | Path | None = Field(
        default=None, description="Directory to write per-sheet files."
    )
    print_areas_dir: str | Path | None = Field(
        default=None, description="Directory to write per-print-area files."
    )
    auto_page_breaks_dir: str | Path | None = Field(
        default=None, description="Directory to write auto page-break files."
    )
    stream: TextIO | None = Field(
        default=None, description="Stream override for primary output (stdout/file)."
    )


class OutputOptions(BaseModel):
    """
    Output-time options for ExStructEngine.

    - format: serialization format/indent.
    - filters: include/exclude flags (rows/shapes/charts/tables/print_areas, size flags).
    - destinations: side outputs (per-sheet, per-print-area, stream override).
    """

    model_config = ConfigDict(extra="forbid")

    format: FormatOptions = Field(
        default_factory=FormatOptions, description="Formatting options."
    )
    filters: FilterOptions = Field(
        default_factory=FilterOptions, description="Include/exclude flags."
    )
    destinations: DestinationOptions = Field(
        default_factory=DestinationOptions, description="Side output destinations."
    )


class ExStructEngine:
    """
    Configurable engine for ExStruct extraction and export.

    Instances are immutable; override options per call if needed.

    Key behaviors:
        - StructOptions: extraction mode and optional table detection params.
        - OutputOptions: serialization format/pretty-print, include/exclude filters, per-sheet/per-print-area output dirs, etc.
        - Main methods:
            extract(path, mode=None) -> WorkbookData
                - Modes: light/standard/verbose
                - light: COM-free; cells + tables + print areas only (shapes/charts empty)
            serialize(workbook, ...) -> str
                - Applies include_* filters, then serializes
            export(workbook, ...)
                - Writes to file/stdout; optionally per-sheet and per-print-area files
            process(file_path, ...)
                - One-shot extract->export (CLI equivalent), with optional PDF/PNG
    """

    def __init__(
        self,
        options: StructOptions | None = None,
        output: OutputOptions | None = None,
    ) -> None:
        self.options = options or StructOptions()
        self.output = output or OutputOptions()

    @staticmethod
    def from_defaults() -> ExStructEngine:
        """Factory to create an engine with default options."""
        return ExStructEngine()

    def _apply_table_params(self) -> None:
        if self.options.table_params:
            set_table_detection_params(**self.options.table_params)

    @contextmanager
    def _table_params_scope(self) -> Iterator[None]:
        """
        Temporarily apply table_params and restore previous global config afterward.
        """
        if not self.options.table_params:
            yield
            return
        prev = cast(TableParams, dict(_cells._DETECTION_CONFIG))
        set_table_detection_params(**self.options.table_params)
        try:
            yield
        finally:
            set_table_detection_params(**prev)

    def _resolve_size_flags(self) -> tuple[bool, bool]:
        """
        Determine whether to include Shape/Chart size fields in output.
        Auto: verbose -> include, others -> exclude.
        """
        include_shape_size = (
            self.output.filters.include_shape_size
            if self.output.filters.include_shape_size is not None
            else self.options.mode == "verbose"
        )
        include_chart_size = (
            self.output.filters.include_chart_size
            if self.output.filters.include_chart_size is not None
            else self.options.mode == "verbose"
        )
        return include_shape_size, include_chart_size

    def _include_print_areas(self) -> bool:
        """
        Decide whether to include print areas in output.
        Auto: light -> False, others -> True.
        """
        if self.output.filters.include_print_areas is None:
            return self.options.mode != "light"
        return self.output.filters.include_print_areas

    def _include_auto_print_areas(self) -> bool:
        """
        Decide whether to include auto page-break areas in output.
        Defaults to False unless explicitly enabled.
        """
        return self.output.filters.include_auto_print_areas

    def _filter_sheet(
        self, sheet: SheetData, include_auto_override: bool | None = None
    ) -> SheetData:
        include_shape_size, include_chart_size = self._resolve_size_flags()
        include_print_areas = self._include_print_areas()
        include_auto_print_areas = (
            include_auto_override
            if include_auto_override is not None
            else self._include_auto_print_areas()
        )
        return SheetData(
            rows=sheet.rows if self.output.filters.include_rows else [],
            shapes=[
                s if include_shape_size else s.model_copy(update={"w": None, "h": None})
                for s in sheet.shapes
            ]
            if self.output.filters.include_shapes
            else [],
            charts=[
                c if include_chart_size else c.model_copy(update={"w": None, "h": None})
                for c in sheet.charts
            ]
            if self.output.filters.include_charts
            else [],
            table_candidates=sheet.table_candidates
            if self.output.filters.include_tables
            else [],
            colors_map=sheet.colors_map,
            print_areas=sheet.print_areas if include_print_areas else [],
            auto_print_areas=sheet.auto_print_areas if include_auto_print_areas else [],
            merged_cells=sheet.merged_cells
            if self.output.filters.include_merged_cells
            else None,
        )

    def _filter_workbook(
        self, wb: WorkbookData, *, include_auto_override: bool | None = None
    ) -> WorkbookData:
        filtered = {
            name: self._filter_sheet(sheet, include_auto_override=include_auto_override)
            for name, sheet in wb.sheets.items()
        }
        return WorkbookData(book_name=wb.book_name, sheets=filtered)

    @staticmethod
    def _ensure_path(path: str | Path) -> Path:
        """Normalize a string or Path input to a Path instance.

        Args:
            path: Path-like input value.

        Returns:
            Path constructed from the given value.
        """

        return path if isinstance(path, Path) else Path(path)

    @classmethod
    def _ensure_optional_path(cls, path: str | Path | None) -> Path | None:
        """Normalize an optional path-like value to Path when provided.

        Args:
            path: Optional path-like input value.

        Returns:
            Normalized Path when provided, otherwise None.
        """

        if path is None:
            return None
        return cls._ensure_path(path)

    def extract(
        self, file_path: str | Path, *, mode: ExtractionMode | None = None
    ) -> WorkbookData:
        """
        Extract a workbook and return normalized workbook data.

        Args:
            file_path: Path to the .xlsx/.xlsm/.xls file to extract.
            mode: Extraction mode; defaults to the engine's StructOptions.mode.
                - light: COM-free; cells, table candidates, and print areas only.
                - standard: Shapes with text/arrows plus charts; print areas included;
                  size fields retained but hidden from default output.
                - verbose: All shapes (with size) and charts (with size).
        """
        chosen_mode = mode or self.options.mode
        include_auto_page_breaks = (
            self.output.filters.include_auto_print_areas
            or self.output.destinations.auto_page_breaks_dir is not None
        )
        normalized_file_path = self._ensure_path(file_path)
        with self._table_params_scope():
            return extract_workbook(
                normalized_file_path,
                mode=chosen_mode,
                include_cell_links=self.options.include_cell_links,
                include_print_areas=None,
                include_auto_page_breaks=include_auto_page_breaks,
                include_colors_map=self.options.include_colors_map,
                include_default_background=self.options.colors.include_default_background,
                ignore_colors=self.options.colors.ignore_colors_set(),
                include_merged_cells=self.options.include_merged_cells,
                include_merged_values_in_rows=self.options.include_merged_values_in_rows,
            )

    def serialize(
        self,
        data: WorkbookData,
        *,
        fmt: Literal["json", "yaml", "yml", "toon"] | None = None,
        pretty: bool | None = None,
        indent: int | None = None,
    ) -> str:
        """
        Serialize a workbook after applying include/exclude filters.

        Args:
            data: Workbook to serialize after filtering.
            fmt: Serialization format; defaults to OutputOptions.format.fmt.
            pretty: Whether to pretty-print JSON output.
            indent: Indentation to use when pretty-printing JSON.
        """
        filtered = self._filter_workbook(data)
        use_fmt = fmt or self.output.format.fmt
        use_pretty = self.output.format.pretty if pretty is None else pretty
        use_indent = self.output.format.indent if indent is None else indent
        return serialize_workbook(
            filtered, fmt=use_fmt, pretty=use_pretty, indent=use_indent
        )

    def export(
        self,
        data: WorkbookData,
        output_path: str | Path | None = None,
        *,
        fmt: Literal["json", "yaml", "yml", "toon"] | None = None,
        pretty: bool | None = None,
        indent: int | None = None,
        sheets_dir: str | Path | None = None,
        print_areas_dir: str | Path | None = None,
        auto_page_breaks_dir: str | Path | None = None,
        stream: TextIO | None = None,
    ) -> None:
        """
        Write filtered workbook data to a file or stream.

        Includes optional per-sheet and per-print-area outputs when destinations are
        provided.

        Args:
            data: Workbook to serialize and write.
            output_path: Target file path (str or Path); writes to stdout when None.
            fmt: Serialization format; defaults to OutputOptions.format.fmt.
            pretty: Whether to pretty-print JSON output.
            indent: Indentation to use when pretty-printing JSON.
            sheets_dir: Directory for per-sheet outputs when provided (str or Path).
            print_areas_dir: Directory for per-print-area outputs when provided (str or Path).
            auto_page_breaks_dir: Directory for auto page-break outputs (str or Path; COM
                environments only).
            stream: Stream override when output_path is None.
        """
        text = self.serialize(data, fmt=fmt, pretty=pretty, indent=indent)
        target_stream = stream or self.output.destinations.stream
        chosen_fmt = fmt or self.output.format.fmt
        chosen_sheets_dir = (
            sheets_dir
            if sheets_dir is not None
            else self.output.destinations.sheets_dir
        )
        chosen_print_areas_dir = (
            print_areas_dir
            if print_areas_dir is not None
            else self.output.destinations.print_areas_dir
        )
        chosen_auto_page_breaks_dir = (
            auto_page_breaks_dir
            if auto_page_breaks_dir is not None
            else self.output.destinations.auto_page_breaks_dir
        )

        normalized_output_path = self._ensure_optional_path(output_path)
        normalized_sheets_dir = self._ensure_optional_path(chosen_sheets_dir)
        normalized_print_areas_dir = self._ensure_optional_path(chosen_print_areas_dir)
        normalized_auto_page_breaks_dir = self._ensure_optional_path(
            chosen_auto_page_breaks_dir
        )

        if normalized_output_path is not None:
            normalized_output_path.write_text(text, encoding="utf-8")
        elif (
            normalized_output_path is None
            and chosen_sheets_dir is None
            and chosen_print_areas_dir is None
            and chosen_auto_page_breaks_dir is None
        ):
            import sys

            stream_target = target_stream or sys.stdout
            stream_target.write(text)
            if not text.endswith("\n"):
                stream_target.write("\n")

        if normalized_sheets_dir is not None:
            filtered = self._filter_workbook(data)
            save_sheets(
                filtered,
                normalized_sheets_dir,
                fmt=chosen_fmt,
                pretty=self.output.format.pretty if pretty is None else pretty,
                indent=self.output.format.indent if indent is None else indent,
            )

        if normalized_print_areas_dir is not None:
            include_shape_size, include_chart_size = self._resolve_size_flags()
            if self._include_print_areas():
                filtered = self._filter_workbook(data)
                save_print_area_views(
                    filtered,
                    normalized_print_areas_dir,
                    fmt=chosen_fmt,
                    pretty=self.output.format.pretty if pretty is None else pretty,
                    indent=self.output.format.indent if indent is None else indent,
                    include_shapes=self.output.filters.include_shapes,
                    include_charts=self.output.filters.include_charts,
                    include_shape_size=include_shape_size,
                    include_chart_size=include_chart_size,
                )

        if normalized_auto_page_breaks_dir is not None:
            include_shape_size, include_chart_size = self._resolve_size_flags()
            filtered = self._filter_workbook(data, include_auto_override=True)
            save_auto_page_break_views(
                filtered,
                normalized_auto_page_breaks_dir,
                fmt=chosen_fmt,
                pretty=self.output.format.pretty if pretty is None else pretty,
                indent=self.output.format.indent if indent is None else indent,
                include_shapes=self.output.filters.include_shapes,
                include_charts=self.output.filters.include_charts,
                include_shape_size=include_shape_size,
                include_chart_size=include_chart_size,
            )

        return None

    def process(
        self,
        file_path: str | Path,
        output_path: str | Path | None = None,
        *,
        out_fmt: str | None = None,
        image: bool = False,
        pdf: bool = False,
        dpi: int = 72,
        mode: ExtractionMode | None = None,
        pretty: bool | None = None,
        indent: int | None = None,
        sheets_dir: str | Path | None = None,
        print_areas_dir: str | Path | None = None,
        auto_page_breaks_dir: str | Path | None = None,
        stream: TextIO | None = None,
    ) -> None:
        """
        One-shot extract->export wrapper (CLI equivalent) with optional PDF/PNG output.

        Args:
            file_path: Input Excel workbook path (str or Path).
            output_path: Target file path (str or Path); writes to stdout when None.
            out_fmt: Serialization format for structured output.
            image: Whether to export PNGs alongside structured output.
            pdf: Whether to export a PDF snapshot alongside structured output.
            dpi: DPI to use when rendering images.
            mode: Extraction mode; defaults to the engine's StructOptions.mode.
            pretty: Whether to pretty-print JSON output.
            indent: Indentation to use when pretty-printing JSON.
            sheets_dir: Directory for per-sheet structured outputs (str or Path).
            print_areas_dir: Directory for per-print-area structured outputs (str or Path).
            auto_page_breaks_dir: Directory for auto page-break outputs (str or Path).
            stream: Stream override when writing to stdout.
        """
        normalized_file_path = self._ensure_path(file_path)
        normalized_output_path = self._ensure_optional_path(output_path)
        normalized_sheets_dir = self._ensure_optional_path(sheets_dir)
        normalized_print_areas_dir = self._ensure_optional_path(print_areas_dir)
        normalized_auto_page_breaks_dir = self._ensure_optional_path(
            auto_page_breaks_dir
        )

        wb = self.extract(normalized_file_path, mode=mode)
        chosen_fmt = out_fmt or self.output.format.fmt
        self.export(
            wb,
            output_path=normalized_output_path,
            fmt=chosen_fmt,  # type: ignore[arg-type]
            pretty=pretty,
            indent=indent,
            sheets_dir=normalized_sheets_dir,
            print_areas_dir=normalized_print_areas_dir,
            auto_page_breaks_dir=normalized_auto_page_breaks_dir,
            stream=stream,
        )

        if pdf or image:
            base_target = normalized_output_path or normalized_file_path.with_suffix(
                ".yaml"
                if chosen_fmt in ("yaml", "yml")
                else ".toon"
                if chosen_fmt == "toon"
                else ".json"
            )
            pdf_path = base_target.with_suffix(".pdf")
            export_pdf(normalized_file_path, pdf_path)
            if image:
                images_dir = pdf_path.parent / f"{pdf_path.stem}_images"
                export_sheet_images(normalized_file_path, images_dir, dpi=dpi)
