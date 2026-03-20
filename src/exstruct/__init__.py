"""Public ExStruct package exports and convenience APIs."""

from __future__ import annotations

from importlib import import_module
import logging
from pathlib import Path
from typing import TYPE_CHECKING, Literal, TextIO

if TYPE_CHECKING:
    from .core.cells import set_table_detection_params
    from .core.integrate import extract_workbook
    from .engine import (
        ColorsOptions,
        DestinationOptions,
        ExStructEngine,
        FilterOptions,
        FormatOptions,
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
    from .io import serialize_workbook
    from .models import (
        CellRow,
        Chart,
        ChartSeries,
        PrintArea,
        PrintAreaView,
        Shape,
        SheetData,
        WorkbookData,
        col_index_to_alpha,
        convert_row_keys_to_alpha,
        convert_sheet_keys_to_alpha,
        convert_workbook_keys_to_alpha,
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
    "FormatOptions",
    "DestinationOptions",
    "ColorsOptions",
    "serialize_workbook",
    "export_auto_page_breaks",
    "col_index_to_alpha",
    "convert_row_keys_to_alpha",
    "convert_sheet_keys_to_alpha",
    "convert_workbook_keys_to_alpha",
]


ExtractionMode = Literal["light", "libreoffice", "standard", "verbose"]

_LAZY_EXPORTS: dict[str, tuple[str, str]] = {
    "ColorsOptions": (".engine", "ColorsOptions"),
    "ConfigError": (".errors", "ConfigError"),
    "DestinationOptions": (".engine", "DestinationOptions"),
    "ExStructEngine": (".engine", "ExStructEngine"),
    "ExstructError": (".errors", "ExstructError"),
    "FilterOptions": (".engine", "FilterOptions"),
    "FormatOptions": (".engine", "FormatOptions"),
    "MissingDependencyError": (".errors", "MissingDependencyError"),
    "OutputOptions": (".engine", "OutputOptions"),
    "PrintArea": (".models", "PrintArea"),
    "PrintAreaError": (".errors", "PrintAreaError"),
    "PrintAreaView": (".models", "PrintAreaView"),
    "RenderError": (".errors", "RenderError"),
    "SerializationError": (".errors", "SerializationError"),
    "StructOptions": (".engine", "StructOptions"),
    "WorkbookData": (".models", "WorkbookData"),
    "CellRow": (".models", "CellRow"),
    "Chart": (".models", "Chart"),
    "ChartSeries": (".models", "ChartSeries"),
    "Shape": (".models", "Shape"),
    "SheetData": (".models", "SheetData"),
    "col_index_to_alpha": (".models", "col_index_to_alpha"),
    "convert_row_keys_to_alpha": (".models", "convert_row_keys_to_alpha"),
    "convert_sheet_keys_to_alpha": (".models", "convert_sheet_keys_to_alpha"),
    "convert_workbook_keys_to_alpha": (".models", "convert_workbook_keys_to_alpha"),
    "export_pdf": (".render", "export_pdf"),
    "export_sheet_images": (".render", "export_sheet_images"),
    "extract_workbook": (".core.integrate", "extract_workbook"),
    "serialize_workbook": (".io", "serialize_workbook"),
    "set_table_detection_params": (".core.cells", "set_table_detection_params"),
}


def _resolve_lazy_export(name: str) -> object:
    module_name, attr_name = _LAZY_EXPORTS[name]
    value = getattr(import_module(module_name, __name__), attr_name)
    globals()[name] = value
    return value


def __getattr__(name: str) -> object:
    if name not in _LAZY_EXPORTS:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    return _resolve_lazy_export(name)


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(__all__))


def extract(
    file_path: str | Path, mode: ExtractionMode = "standard", *, alpha_col: bool = False
) -> WorkbookData:
    """
    Extracts an Excel workbook into a WorkbookData structure.

    Parameters:
        file_path (str | Path): Path to the workbook file (.xlsx, .xlsm, .xls).
        mode (ExtractionMode): Extraction detail level. "light" includes cells and table detection only (no COM, shapes/charts empty; print areas via openpyxl). "libreoffice" is a best-effort non-COM mode that adds merged cells, shapes, connectors, and charts when the LibreOffice backend is available. "standard" includes texted shapes, arrows, charts (COM if available) and print areas. "verbose" also includes shape/chart sizes, cell link map, colors map, and formulas map.
        alpha_col: When True, convert CellRow column keys to Excel-style ABC names (A, B, ..., Z, AA, ...) instead of 0-based numeric strings.

    Returns:
        WorkbookData: Parsed workbook representation containing sheets, rows, shapes, charts, and print areas.
    """
    from .engine import ExStructEngine, StructOptions

    include_links = True if mode == "verbose" else False
    include_colors_map = True if mode == "verbose" else None
    include_formulas_map = True if mode == "verbose" else None
    engine = ExStructEngine(
        options=StructOptions(
            mode=mode,
            include_cell_links=include_links,
            include_colors_map=include_colors_map,
            include_formulas_map=include_formulas_map,
            alpha_col=alpha_col,
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
    include_backend_metadata: bool = False,
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
    from .io import save_as_json, save_as_toon, save_as_yaml

    dest = Path(path)
    format_hint = (fmt or dest.suffix.lstrip(".") or "json").lower()
    match format_hint:
        case "json":
            save_as_json(
                data,
                dest,
                pretty=pretty,
                indent=indent,
                include_backend_metadata=include_backend_metadata,
            )
        case "yaml" | "yml":
            save_as_yaml(data, dest, include_backend_metadata=include_backend_metadata)
        case "toon":
            save_as_toon(data, dest, include_backend_metadata=include_backend_metadata)
        case _:
            raise ValueError(f"Unsupported export format: {format_hint}")


def export_sheets(
    data: WorkbookData,
    dir_path: str | Path,
    *,
    include_backend_metadata: bool = False,
) -> dict[str, Path]:
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
    from .io import save_sheets

    return save_sheets(
        data,
        Path(dir_path),
        fmt="json",
        include_backend_metadata=include_backend_metadata,
    )


def export_sheets_as(
    data: WorkbookData,
    dir_path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
    include_backend_metadata: bool = False,
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
    from .io import save_sheets

    return save_sheets(
        data,
        Path(dir_path),
        fmt=fmt,
        pretty=pretty,
        indent=indent,
        include_backend_metadata=include_backend_metadata,
    )


def export_print_areas_as(
    data: WorkbookData,
    dir_path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
    normalize: bool = False,
    include_backend_metadata: bool = False,
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
    from .io import save_print_area_views

    return save_print_area_views(
        data,
        Path(dir_path),
        fmt=fmt,
        pretty=pretty,
        indent=indent,
        normalize=normalize,
        include_backend_metadata=include_backend_metadata,
    )


def export_auto_page_breaks(
    data: WorkbookData,
    dir_path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
    normalize: bool = False,
    include_backend_metadata: bool = False,
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
    from .errors import PrintAreaError
    from .io import save_auto_page_break_views

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
        include_backend_metadata=include_backend_metadata,
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
    *,
    alpha_col: bool = False,
    include_backend_metadata: bool = False,
) -> None:
    """
    Convenience wrapper: extract -> serialize (file or stdout) -> optional PDF/PNG.

    Args:
        file_path: Input Excel workbook (path string or Path).
        output_path: None for stdout; otherwise, write to file (string or Path).
        out_fmt: json/yaml/yml/toon.
        image: True to also output PNGs (requires Excel + COM + pypdfium2 and is
            not supported in `mode="libreoffice"`).
        pdf: True to also output PDF (requires Excel + COM + pypdfium2 and is not
            supported in `mode="libreoffice"`).
        dpi: DPI for image output.
        mode: light/libreoffice/standard/verbose (same meaning as `extract`).
        pretty: Pretty-print JSON.
        indent: JSON indent width.
        sheets_dir: Directory to write per-sheet files (string or Path).
        print_areas_dir: Directory to write per-print-area files (string or Path).
        auto_page_breaks_dir: Directory to write per-auto-page-break files (COM only
            and not supported in `mode="libreoffice"`).
        stream: IO override when output_path is None.
        alpha_col: When True, convert CellRow column keys to Excel-style
            ABC names (A, B, ...) instead of 0-based numeric strings.
        include_backend_metadata: When True, include shape/chart backend metadata
            fields (`provenance`, `approximation_level`, `confidence`) in output.

    Raises:
        ConfigError: If `mode="libreoffice"` is combined with PDF/PNG rendering or
            auto page-break export.
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
    from .engine import (
        DestinationOptions,
        ExStructEngine,
        FilterOptions,
        FormatOptions,
        OutputOptions,
        StructOptions,
    )

    engine = ExStructEngine(
        options=StructOptions(mode=mode, alpha_col=alpha_col),
        output=OutputOptions(
            format=FormatOptions(fmt=out_fmt, pretty=pretty, indent=indent),
            filters=FilterOptions(
                include_print_areas=None if mode == "light" else True,
                include_shape_size=True if mode == "verbose" else False,
                include_chart_size=True if mode == "verbose" else False,
                include_backend_metadata=include_backend_metadata,
            ),
            destinations=DestinationOptions(
                sheets_dir=sheets_dir,
                print_areas_dir=print_areas_dir,
                auto_page_breaks_dir=auto_page_breaks_dir,
                stream=stream,
            ),
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
