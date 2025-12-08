from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Optional, TextIO
from contextlib import contextmanager

from .core.integrate import extract_workbook
from .core import cells as _cells
from .core.cells import set_table_detection_params
from .io import save_as_json, save_as_toon, save_as_yaml, save_sheets, serialize_workbook
from .models import SheetData, WorkbookData
from .render import export_pdf, export_sheet_images

ExtractionMode = Literal["light", "standard", "verbose"]


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
    """

    mode: ExtractionMode = "standard"
    table_params: Optional[dict] = None  # forwarded to set_table_detection_params if provided
    include_cell_links: Optional[bool] = None  # None -> auto: verbose=True, others=False


@dataclass(frozen=True)
class OutputOptions:
    """
    Output-time options for ExStructEngine.

    Attributes:
        fmt: Default export format. One of "json", "yaml", "yml", "toon".
        pretty: Whether to pretty-print JSON; default False (compact).
        indent: Explicit indent size. If None and pretty=True, indent=2 for JSON.
        include_rows: Include SheetData.rows in output (set False to drop).
        include_shapes: Include SheetData.shapes in output.
        include_charts: Include SheetData.charts in output.
        include_tables: Include SheetData.table_candidates in output.
        include_print_areas: Include SheetData.print_areas in output.
        sheets_dir: Optional directory to write per-sheet files (in the chosen fmt).
        stream: Optional default stream for stdout output when output_path is None.
    """

    fmt: Literal["json", "yaml", "yml", "toon"] = "json"
    pretty: bool = False
    indent: int | None = None
    include_rows: bool = True
    include_shapes: bool = True
    include_charts: bool = True
    include_tables: bool = True
    include_print_areas: bool = True
    sheets_dir: Path | None = None
    stream: TextIO | None = None


class ExStructEngine:
    """
    Configurable engine for ExStruct extraction and export.

    Instances are immutable; override options per call if needed.

    Key behaviors:
        - Uses StructOptions for extraction defaults (mode, table_params).
        - Uses OutputOptions for serialization defaults (fmt, pretty/indent, include* filters).
        - Methods:
            extract(path, mode=None) -> WorkbookData
            serialize(workbook, fmt=None, pretty=None, indent=None) -> str
            export(workbook, output_path=None, fmt=None, pretty=None, indent=None,
                   sheets_dir=None, stream=None) -> None
            process(file_path, output_path=None, out_fmt=None, image=False, pdf=False,
                    dpi=72, mode=None, pretty=None, indent=None, sheets_dir=None,
                    stream=None) -> None
    """

    def __init__(
        self,
        options: StructOptions | None = None,
        output: OutputOptions | None = None,
    ) -> None:
        self.options = options or StructOptions()
        self.output = output or OutputOptions()

    @staticmethod
    def from_defaults() -> "ExStructEngine":
        """Factory to create an engine with default options."""
        return ExStructEngine()

    def _apply_table_params(self) -> None:
        if self.options.table_params:
            set_table_detection_params(**self.options.table_params)

    @contextmanager
    def _table_params_scope(self):
        """
        Temporarily apply table_params and restore previous global config afterward.
        """
        if not self.options.table_params:
            yield
            return
        prev = dict(_cells._DETECTION_CONFIG)  # type: ignore[attr-defined]
        set_table_detection_params(**self.options.table_params)
        try:
            yield
        finally:
            set_table_detection_params(**prev)

    def _filter_sheet(self, sheet: SheetData) -> SheetData:
        return SheetData(
            rows=sheet.rows if self.output.include_rows else [],
            shapes=sheet.shapes if self.output.include_shapes else [],
            charts=sheet.charts if self.output.include_charts else [],
            table_candidates=sheet.table_candidates if self.output.include_tables else [],
            print_areas=sheet.print_areas if self.output.include_print_areas else [],
        )

    def _filter_workbook(self, wb: WorkbookData) -> WorkbookData:
        filtered = {
            name: self._filter_sheet(sheet)
            for name, sheet in wb.sheets.items()
        }
        return WorkbookData(book_name=wb.book_name, sheets=filtered)

    def extract(self, file_path: str | Path, *, mode: ExtractionMode | None = None) -> WorkbookData:
        """Extract workbook semantic structure with the configured options."""
        chosen_mode = mode or self.options.mode
        if chosen_mode not in ("light", "standard", "verbose"):
            raise ValueError(f"Unsupported mode: {chosen_mode}")
        include_links = (
            self.options.include_cell_links
            if self.options.include_cell_links is not None
            else chosen_mode == "verbose"
        )
        include_print_areas = chosen_mode != "light"
        with self._table_params_scope():
            return extract_workbook(
                Path(file_path),
                mode=chosen_mode,
                include_cell_links=include_links,
                include_print_areas=include_print_areas,
            )

    def serialize(
        self,
        data: WorkbookData,
        *,
        fmt: Optional[Literal["json", "yaml", "yml", "toon"]] = None,
        pretty: Optional[bool] = None,
        indent: int | None = None,
    ) -> str:
        """
        Serialize WorkbookData using configured output defaults, applying include/exclude filters.
        """
        filtered = self._filter_workbook(data)
        use_fmt = (fmt or self.output.fmt)
        use_pretty = self.output.pretty if pretty is None else pretty
        use_indent = self.output.indent if indent is None else indent
        return serialize_workbook(filtered, fmt=use_fmt, pretty=use_pretty, indent=use_indent)

    def export(
        self,
        data: WorkbookData,
        output_path: Path | None = None,
        *,
        fmt: Optional[Literal["json", "yaml", "yml", "toon"]] = None,
        pretty: Optional[bool] = None,
        indent: int | None = None,
        sheets_dir: Path | None = None,
        stream: TextIO | None = None,
    ) -> None:
        """
        Write WorkbookData to disk or stdout (when output_path is None).
        Applies include/exclude filters before serialization.
        """
        text = self.serialize(data, fmt=fmt, pretty=pretty, indent=indent)
        target_stream = stream or self.output.stream
        chosen_fmt = (fmt or self.output.fmt)
        chosen_sheets_dir = sheets_dir if sheets_dir is not None else self.output.sheets_dir

        def _suffix_for(fmt_val: str) -> str:
            if fmt_val in ("yaml", "yml"):
                return ".yaml"
            if fmt_val == "toon":
                return ".toon"
            if fmt_val == "json":
                return ".json"
            raise ValueError(f"Unsupported export format: {fmt_val}")

        if output_path is not None:
            output_path.write_text(text, encoding="utf-8")
        else:
            import sys

            stream_target = target_stream or sys.stdout
            stream_target.write(text)
            if not text.endswith("\n"):
                stream_target.write("\n")

        if chosen_sheets_dir is not None:
            filtered = self._filter_workbook(data)
            save_sheets(
                filtered,
                chosen_sheets_dir,
                fmt=chosen_fmt,
                pretty=self.output.pretty if pretty is None else pretty,
                indent=self.output.indent if indent is None else indent,
            )

        return None

    def process(
        self,
        file_path: Path,
        output_path: Path | None = None,
        *,
        out_fmt: Optional[str] = None,
        image: bool = False,
        pdf: bool = False,
        dpi: int = 72,
        mode: ExtractionMode | None = None,
        pretty: bool | None = None,
        indent: int | None = None,
        sheets_dir: Path | None = None,
        stream: TextIO | None = None,
    ) -> None:
        """
        Convenience wrapper: extract, export (to file or stdout), and optionally render PDF/PNG.
        """
        wb = self.extract(file_path, mode=mode)
        chosen_fmt = out_fmt or self.output.fmt
        self.export(
            wb,
            output_path=output_path,
            fmt=chosen_fmt,  # type: ignore[arg-type]
            pretty=pretty,
            indent=indent,
            sheets_dir=sheets_dir,
            stream=stream,
        )

        if pdf or image:
            base_target = output_path or file_path.with_suffix(
                ".yaml" if chosen_fmt in ("yaml", "yml") else ".toon" if chosen_fmt == "toon" else ".json"
            )
            pdf_path = base_target.with_suffix(".pdf")
            export_pdf(file_path, pdf_path)
            if image:
                images_dir = pdf_path.parent / f"{pdf_path.stem}_images"
                export_sheet_images(file_path, images_dir, dpi=dpi)
