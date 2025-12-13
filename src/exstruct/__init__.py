from __future__ import annotations

import logging
from pathlib import Path
from typing import Literal, TextIO

from .core.cells import set_table_detection_params
from .core.integrate import extract_workbook
from .engine import (
    DestinationOptions,
    ExStructEngine,
    FilterOptions,
    OutputOptions,
    StructOptions,
)
from .errors import (
    ConfigError,
    ExstructError,
    MissingDependencyError,
    PrintAreaError,
    RenderError,
    SerializationError,
)
from .io import (
    save_as_json,
    save_as_toon,
    save_as_yaml,
    save_auto_page_break_views,
    save_print_area_views,
    save_sheets,
    serialize_workbook,
)
from .models import (
    CellRow,
    Chart,
    ChartSeries,
    PrintArea,
    PrintAreaView,
    Shape,
    SheetData,
    WorkbookData,
)
from .render import export_pdf, export_sheet_images

logger = logging.getLogger(__name__)

__all__ = [
    "extract",
    "export",
    "export_sheets",
    "export_sheets_as",
    "export_print_areas_as",
    "export_auto_page_breaks",
    "export_pdf",
    "export_sheet_images",
    "ExstructError",
    "ConfigError",
    "MissingDependencyError",
    "RenderError",
    "SerializationError",
    "PrintAreaError",
    "process_excel",
    "ExtractionMode",
    "CellRow",
    "Shape",
    "ChartSeries",
    "Chart",
    "SheetData",
    "WorkbookData",
    "PrintArea",
    "PrintAreaView",
    "set_table_detection_params",
    "extract_workbook",
    "ExStructEngine",
    "StructOptions",
    "OutputOptions",
    "FilterOptions",
    "DestinationOptions",
    "serialize_workbook",
    "export_auto_page_breaks",
]


ExtractionMode = Literal["light", "standard", "verbose"]


def extract(file_path: str | Path, mode: ExtractionMode = "standard") -> WorkbookData:
    """
    Extract an Excel workbook into WorkbookData.

    Args:
        file_path: Path to .xlsx/.xlsm/.xls.
        mode: "light" / "standard" / "verbose"
            - light: cells + table detection only (no COM, shapes/charts empty). Print areas via openpyxl.
            - standard: texted shapes + arrows + charts (COM if available), print areas included. Shape/chart size is kept but hidden by default in output.
            - verbose: all shapes (including textless) with size, charts with size.
    """
    include_links = True if mode == "verbose" else False
    engine = ExStructEngine(
        options=StructOptions(mode=mode, include_cell_links=include_links)
    )
    return engine.extract(file_path, mode=mode)


def export(
    data: WorkbookData,
    path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] | None = None,
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> None:
    """
    Save WorkbookData to a file (format inferred from extension).

    Args:
        data: WorkbookData from `extract` or similar
        path: destination path; extension is used to infer format
        fmt: explicitly set format if desired (json/yaml/yml/toon)
        pretty: pretty-print JSON
        indent: JSON indent width (defaults to 2 when pretty=True and indent is None)
    """
    dest = Path(path)
    format_hint = (fmt or dest.suffix.lstrip(".") or "json").lower()
    match format_hint:
        case "json":
            save_as_json(data, dest, pretty=pretty, indent=indent)
        case "yaml" | "yml":
            save_as_yaml(data, dest)
        case "toon":
            save_as_toon(data, dest)
        case _:
            raise ValueError(f"Unsupported export format: {format_hint}")


def export_sheets(data: WorkbookData, dir_path: str | Path) -> dict[str, Path]:
    """
    Export each sheet as an individual JSON file.

    - Payload: {book_name, sheet_name, sheet: SheetData}
    - Returns: {sheet_name: Path}
    """
    return save_sheets(data, Path(dir_path), fmt="json")


def export_sheets_as(
    data: WorkbookData,
    dir_path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> dict[str, Path]:
    """Export each sheet in the given format (json/yaml/toon); returns sheet name to path map."""
    return save_sheets(data, Path(dir_path), fmt=fmt, pretty=pretty, indent=indent)


def export_print_areas_as(
    data: WorkbookData,
    dir_path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
    normalize: bool = False,
) -> dict[str, Path]:
    """
    Export each print area as a PrintAreaView.

    Args:
        data: WorkbookData that contains print areas
        dir_path: output directory
        fmt: json/yaml/yml/toon
        pretty/indent: JSON formatting options
        normalize: rebase row/col indices to the print-area origin when True
    Returns:
        dict mapping area key to path (e.g., "Sheet1#1": /.../Sheet1_area1_...json)
    """
    return save_print_area_views(
        data,
        Path(dir_path),
        fmt=fmt,
        pretty=pretty,
        indent=indent,
        normalize=normalize,
    )


def export_auto_page_breaks(
    data: WorkbookData,
    dir_path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
    normalize: bool = False,
) -> dict[str, Path]:
    """
    Export auto page-break areas (COM-computed) as PrintAreaView files.

    Args:
        data: WorkbookData containing auto_print_areas (COM extraction with auto breaks enabled)
        dir_path: output directory
        fmt: json/yaml/yml/toon
        pretty/indent: JSON formatting options
        normalize: rebase row/col indices to the area origin when True
    """
    if not any(sheet.auto_print_areas for sheet in data.sheets.values()):
        message = "No auto page-break areas found. Enable COM-based auto page breaks before exporting."
        logger.warning(message)
        raise PrintAreaError(message)
    return save_auto_page_break_views(
        data,
        Path(dir_path),
        fmt=fmt,
        pretty=pretty,
        indent=indent,
        normalize=normalize,
    )


def process_excel(
    file_path: Path,
    output_path: Path | None = None,
    out_fmt: str = "json",
    image: bool = False,
    pdf: bool = False,
    dpi: int = 72,
    mode: ExtractionMode = "standard",
    pretty: bool = False,
    indent: int | None = None,
    sheets_dir: Path | None = None,
    print_areas_dir: Path | None = None,
    auto_page_breaks_dir: Path | None = None,
    stream: TextIO | None = None,
) -> None:
    """
    Convenience wrapper: extract → serialize (file or stdout) → optional PDF/PNG.

    Args:
        file_path: input Excel
        output_path: None for stdout; otherwise, write to file
        out_fmt: json/yaml/yml/toon
        image/pdf: True to also output PNG/PDF (requires Excel + pypdfium2)
        dpi: DPI for image output
        mode: light/standard/verbose (same meaning as `extract`)
        pretty/indent: JSON formatting
        sheets_dir: directory to write per-sheet files
        print_areas_dir: directory to write per-print-area files
        auto_page_breaks_dir: directory to write per-auto-page-break files (COM only)
        stream: IO override when output_path is None
    """
    engine = ExStructEngine(
        options=StructOptions(mode=mode),
        output=OutputOptions(
            fmt=out_fmt,
            pretty=pretty,
            indent=indent,
            sheets_dir=sheets_dir,
            print_areas_dir=print_areas_dir,
            auto_page_breaks_dir=auto_page_breaks_dir,
            include_print_areas=None if mode == "light" else True,
            include_shape_size=True if mode == "verbose" else False,
            include_chart_size=True if mode == "verbose" else False,
            stream=stream,
        ),
    )
    engine.process(
        file_path=file_path,
        output_path=output_path,
        out_fmt=out_fmt,
        image=image,
        pdf=pdf,
        dpi=dpi,
        mode=mode,
        pretty=pretty,
        indent=indent,
        sheets_dir=sheets_dir,
        print_areas_dir=print_areas_dir,
        auto_page_breaks_dir=auto_page_breaks_dir,
        stream=stream,
    )
