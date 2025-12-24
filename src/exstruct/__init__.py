from __future__ import annotations

import logging
from pathlib import Path
from typing import Literal, TextIO

from .core.cells import set_table_detection_params
from .core.integrate import extract_workbook
from .engine import (
    ColorsOptions,
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
    "ColorsOptions",
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
            - verbose: all shapes (including textless) with size, charts with size, and colors_map.

    Returns:
        WorkbookData containing sheets, rows, shapes, charts, and print areas.

    Raises:
        ValueError: If an invalid mode is provided.

    Examples:
        Extract with hyperlinks (verbose) and inspect table candidates:

        >>> from exstruct import extract
        >>> wb = extract("input.xlsx", mode="verbose")
        >>> wb.sheets["Sheet1"].table_candidates
        ['A1:B5']
    """
    include_links = True if mode == "verbose" else False
    include_colors_map = True if mode == "verbose" else None
    engine = ExStructEngine(
        options=StructOptions(
            mode=mode,
            include_cell_links=include_links,
            include_colors_map=include_colors_map,
        )
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

    Raises:
        ValueError: If the format is unsupported.

    Examples:
        Write pretty JSON and YAML (requires pyyaml):

        >>> from exstruct import export, extract
        >>> wb = extract("input.xlsx")
        >>> export(wb, "out.json", pretty=True)
        >>> export(wb, "out.yaml", fmt="yaml")  # doctest: +SKIP
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

    Args:
        data: WorkbookData to split by sheet.
        dir_path: Output directory.

    Returns:
        Mapping from sheet name to written JSON path.

    Examples:
        >>> from exstruct import export_sheets, extract
        >>> wb = extract("input.xlsx")
        >>> paths = export_sheets(wb, "out_sheets")
        >>> "Sheet1" in paths
        True
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
    """
    Export each sheet in the given format (json/yaml/toon); returns sheet name to path map.

    Args:
        data: WorkbookData to split by sheet.
        dir_path: Output directory.
        fmt: Output format; inferred defaults to json.
        pretty: Pretty-print JSON.
        indent: JSON indent width (defaults to 2 when pretty=True and indent is None).

    Returns:
        Mapping from sheet name to written file path.

    Raises:
        ValueError: If an unsupported format is passed.

    Examples:
        Export per sheet as YAML (requires pyyaml):

        >>> from exstruct import export_sheets_as, extract
        >>> wb = extract("input.xlsx")
        >>> _ = export_sheets_as(wb, "out_yaml", fmt="yaml")  # doctest: +SKIP
    """
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
        pretty: Pretty-print JSON output.
        indent: JSON indent width (defaults to 2 when pretty is True and indent is None).
        normalize: rebase row/col indices to the print-area origin when True

    Returns:
        dict mapping area key to path (e.g., "Sheet1#1": /.../Sheet1_area1_...json)

    Examples:
        Export print areas when present:

        >>> from exstruct import export_print_areas_as, extract
        >>> wb = extract("input.xlsx", mode="standard")
        >>> paths = export_print_areas_as(wb, "areas")
        >>> isinstance(paths, dict)
        True
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
        pretty: Pretty-print JSON output.
        indent: JSON indent width (defaults to 2 when pretty is True and indent is None).
        normalize: rebase row/col indices to the area origin when True

    Returns:
        dict mapping area key to path (e.g., "Sheet1#1": /.../Sheet1_auto_page1_...json)

    Raises:
        PrintAreaError: If no auto page-break areas are present.

    Examples:
        >>> from exstruct import export_auto_page_breaks, extract
        >>> wb = extract("input.xlsx", mode="standard")
        >>> try:
        ...     export_auto_page_breaks(wb, "auto_areas")
        ... except PrintAreaError:
        ...     pass
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
    file_path: str | Path,
    output_path: str | Path | None = None,
    out_fmt: str = "json",
    image: bool = False,
    pdf: bool = False,
    dpi: int = 72,
    mode: ExtractionMode = "standard",
    pretty: bool = False,
    indent: int | None = None,
    sheets_dir: str | Path | None = None,
    print_areas_dir: str | Path | None = None,
    auto_page_breaks_dir: str | Path | None = None,
    stream: TextIO | None = None,
) -> None:
    """
    Convenience wrapper: extract -> serialize (file or stdout) -> optional PDF/PNG.

    Args:
        file_path: Input Excel workbook (path string or Path).
        output_path: None for stdout; otherwise, write to file (string or Path).
        out_fmt: json/yaml/yml/toon.
        image: True to also output PNGs (requires Excel + COM + pypdfium2).
        pdf: True to also output PDF (requires Excel + COM + pypdfium2).
        dpi: DPI for image output.
        mode: light/standard/verbose (same meaning as `extract`).
        pretty: Pretty-print JSON.
        indent: JSON indent width.
        sheets_dir: Directory to write per-sheet files (string or Path).
        print_areas_dir: Directory to write per-print-area files (string or Path).
        auto_page_breaks_dir: Directory to write per-auto-page-break files (COM only).
        stream: IO override when output_path is None.

    Raises:
        ValueError: If an unsupported format or mode is given.
        PrintAreaError: When exporting auto page breaks without available data.
        RenderError: When rendering fails (Excel/COM/pypdfium2 issues).

    Examples:
        Extract and write JSON to stdout, plus per-sheet files:

        >>> from pathlib import Path
        >>> from exstruct import process_excel
        >>> process_excel(Path("input.xlsx"), output_path=None, sheets_dir=Path("sheets"))

        Render PDF only (COM + Excel required):

        >>> process_excel(Path("input.xlsx"), output_path=Path("out.json"), pdf=True)  # doctest: +SKIP
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
