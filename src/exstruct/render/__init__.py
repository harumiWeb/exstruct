from __future__ import annotations

import logging
import multiprocessing as mp
import os
from pathlib import Path
import shutil
import tempfile
from types import ModuleType
from typing import Protocol, cast

import xlwings as xw

from ..errors import MissingDependencyError, RenderError

logger = logging.getLogger(__name__)


def _require_excel_app() -> xw.App:
    """Ensure Excel COM is available and return an App; otherwise raise."""
    try:
        app = xw.App(add_book=False, visible=False)
        return app
    except Exception as e:
        raise RenderError(
            "Excel (COM) is not available. Rendering (PDF/image) requires a desktop Excel installation."
        ) from e


def export_pdf(excel_path: str | Path, output_pdf: str | Path) -> list[str]:
    """Export an Excel workbook to PDF via Excel COM and return sheet names in order."""
    normalized_excel_path = Path(excel_path)
    normalized_output_pdf = Path(output_pdf)
    normalized_output_pdf.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as td:
        temp_dir = Path(td)
        temp_xlsx = temp_dir / "book.xlsx"
        temp_pdf = temp_dir / "book.pdf"

        app: xw.App | None = None
        wb: xw.Book | None = None
        try:
            app = _require_excel_app()
            app.display_alerts = False
            wb = app.books.open(str(normalized_excel_path))
            sheet_names = [s.name for s in wb.sheets]
            wb.api.SaveAs(str(temp_xlsx))
            wb.api.ExportAsFixedFormat(0, str(temp_pdf))
            shutil.copy(temp_pdf, normalized_output_pdf)
        except RenderError:
            raise
        except Exception as exc:
            raise RenderError(
                "Failed to export PDF for "
                f"'{normalized_excel_path}' to '{normalized_output_pdf}'."
            ) from exc
        finally:
            if wb is not None:
                wb.close()
            if app is not None:
                app.quit()
        if not normalized_output_pdf.exists():
            raise RenderError(f"Failed to export PDF to '{normalized_output_pdf}'.")
    return sheet_names


def _require_pdfium() -> ModuleType:
    """Ensure pypdfium2 is installed; otherwise raise with guidance."""
    try:
        import pypdfium2 as pdfium
    except ImportError as e:
        raise MissingDependencyError(
            "Image rendering requires pypdfium2. Install it via `pip install pypdfium2 pillow` or add the 'render' extra."
        ) from e
    return cast(ModuleType, pdfium)


def export_sheet_images(
    excel_path: str | Path, output_dir: str | Path, dpi: int = 144
) -> list[Path]:
    """
    Export each worksheet in the given Excel workbook to PNG files and return the image paths in workbook order.

    Returns:
        paths (list[Path]): Paths to the generated PNG files, ordered by the corresponding worksheets.

    Raises:
        RenderError: If export or rendering fails.
    """
    normalized_excel_path = Path(excel_path)
    normalized_output_dir = Path(output_dir)
    normalized_output_dir.mkdir(parents=True, exist_ok=True)
    use_subprocess = _use_render_subprocess()
    pdfium = _ensure_pdfium(use_subprocess)

    try:
        with tempfile.TemporaryDirectory() as td:
            temp_dir = Path(td)
            return _export_sheet_images_with_app(
                normalized_excel_path,
                normalized_output_dir,
                temp_dir,
                dpi,
                use_subprocess,
                pdfium,
            )
    except RenderError:
        raise
    except Exception as exc:
        raise RenderError(
            f"Failed to export sheet images to '{normalized_output_dir}'."
        ) from exc


def _sanitize_sheet_filename(name: str) -> str:
    r"""
    Create a filesystem-safe filename derived from an Excel sheet name.

    Replaces characters that are not allowed in filenames (\/:*?"<>|) with underscores, trims surrounding whitespace, and returns "sheet" if the result is empty.

    Parameters:
        name (str): Original sheet name.

    Returns:
        safe_name (str): Filename-safe string derived from `name`.
    """
    return "".join("_" if c in '\\/:*?"<>|' else c for c in name).strip() or "sheet"


class _PageSetupProtocol(Protocol):
    """Protocol for Excel PageSetup objects exposing PrintArea."""

    PrintArea: object


class _SheetApiProtocol(Protocol):
    """Protocol for Excel sheet COM APIs used by render helpers."""

    PageSetup: _PageSetupProtocol

    def ExportAsFixedFormat(  # noqa: N802
        self, file_format: int, output_path: str, *args: object, **kwargs: object
    ) -> None:
        """Export the sheet or workbook to a fixed-format file (for example, PDF or XPS).

        Parameters:
            file_format (int): Excel XlFixedFormatType enum value indicating the output format (e.g., the constant for PDF).
            output_path (str): Filesystem path where the fixed-format file will be written.
            *args (object): Additional positional arguments forwarded to the underlying Excel COM ExportAsFixedFormat call.
            **kwargs (object): Additional keyword arguments forwarded to the underlying Excel COM ExportAsFixedFormat call.
        """
        ...


def _iter_sheet_apis(wb: xw.Book) -> list[tuple[int, str, _SheetApiProtocol]]:
    """
    Enumerate workbook sheets and return each sheet's zero-based index, display name, and COM API handle in workbook order.

    If direct COM access to Worksheets is unavailable, falls back to iterating wb.sheets to build the same list.

    Returns:
        List[tuple[int, str, _SheetApiProtocol]]: Tuples of (zero-based sheet index, sheet name, sheet COM API handle) in workbook order.
    """
    try:
        ws_collection = getattr(getattr(wb, "api", None), "Worksheets", None)
        if ws_collection is None:
            raise AttributeError("Worksheets not available")
        count = int(ws_collection.Count)
        sheets: list[tuple[int, str, _SheetApiProtocol]] = []
        for i in range(1, count + 1):
            ws_api = cast(_SheetApiProtocol, ws_collection.Item(i))
            name = str(getattr(ws_api, "Name", f"Sheet{i}"))
            sheets.append((i - 1, name, ws_api))
        return sheets
    except Exception:
        return [
            (
                index,
                sheet.name,
                cast(_SheetApiProtocol, sheet.api),
            )
            for index, sheet in enumerate(wb.sheets)
        ]


def _build_sheet_export_plan(
    wb: xw.Book,
) -> list[tuple[str, _SheetApiProtocol, str | None]]:
    """
    Build an ordered export plan mapping each worksheet to its print areas.

    Each returned tuple is (sheet_name, sheet_api, print_area). The list preserves workbook sheet order; for sheets with no defined print areas `print_area` is `None`, and for sheets with multiple print areas there is one tuple per area.
    """
    plan: list[tuple[str, _SheetApiProtocol, str | None]] = []
    for _, sheet_name, sheet_api in _iter_sheet_apis(wb):
        areas = _extract_print_areas(sheet_api)
        if not areas:
            plan.append((sheet_name, sheet_api, None))
            continue
        for area in areas:
            plan.append((sheet_name, sheet_api, area))
    return plan


def _extract_print_areas(sheet_api: _SheetApiProtocol) -> list[str]:
    """
    Extract the sheet's print-area ranges as a list of strings.

    Retrieves the PageSetup.PrintArea value from the provided sheet API, splits it by commas while respecting single-quoted sections, and returns each range as a separate string. If the sheet has no print area or the property is inaccessible, an empty list is returned.

    Parameters:
        sheet_api (_SheetApiProtocol): Excel sheet API object exposing a `PageSetup.PrintArea` attribute.

    Returns:
        list[str]: List of print-area range strings in the order they appear, or an empty list if none are defined or on access failure.
    """
    try:
        page_setup = getattr(sheet_api, "PageSetup", None)
        if page_setup is None:
            return []
        raw = str(getattr(page_setup, "PrintArea", "") or "")
    except Exception:
        return []
    if not raw:
        return []
    return _split_csv_respecting_quotes(raw)


def _split_csv_respecting_quotes(raw: str) -> list[str]:
    """
    Split a comma-separated string into parts while treating single-quoted sections as atomic.

    This function splits raw on commas that are not inside single quotes. Text enclosed in single quotes is preserved (including internal commas). Two consecutive single quotes inside a quoted section are treated as an escaped single-quote pair. Leading and trailing whitespace is trimmed from each part and empty parts are removed.

    Parameters:
        raw (str): The input CSV-like string that may contain single-quoted segments.

    Returns:
        list[str]: A list of non-empty tokens obtained from splitting `raw` by unquoted commas,
                   with surrounding whitespace removed and quoted segments preserved.
    """
    parts: list[str] = []
    buf: list[str] = []
    in_quote = False
    i = 0
    while i < len(raw):
        ch = raw[i]
        if ch == "'":
            if in_quote and i + 1 < len(raw) and raw[i + 1] == "'":
                buf.append("''")
                i += 2
                continue
            in_quote = not in_quote
            buf.append(ch)
            i += 1
            continue
        if ch == "," and not in_quote:
            parts.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    if buf:
        parts.append("".join(buf).strip())
    return [p for p in parts if p]


def _rename_pages_for_print_area(
    paths: list[Path],
    output_dir: Path,
    base_index: int,
    safe_name: str,
) -> list[Path]:
    """
    Rename the given image files so each gets a unique numeric prefix based on a base index and a safe sheet name.

    Parameters:
        paths (list[Path]): Existing image files for a single sheet or print area (may include per-page suffixes).
        output_dir (Path): Directory where renamed files will reside.
        base_index (int): Zero-based starting index used to compute the numeric prefix for each output file.
        safe_name (str): Filesystem-safe base name to use after the numeric prefix.

    Returns:
        list[Path]: Paths to the renamed files in the same order as input, each named "{index:02d}_{safe_name}.png".
    """
    renamed: list[Path] = []
    for path in paths:
        page_index = _page_index_from_suffix(path.stem)
        new_index = base_index + page_index
        new_path = output_dir / f"{new_index + 1:02d}_{safe_name}.png"
        if path != new_path:
            path.replace(new_path)
        renamed.append(new_path)
    return renamed


def _page_index_from_suffix(stem: str) -> int:
    """
    Extracts a zero-based page index from a filename stem ending with a "_pNN" numeric suffix.

    If the stem ends with "_p" followed by digits, returns that number minus one. If the suffix is missing, non-numeric, or less than 1, returns 0.

    Parameters:
        stem (str): Filename stem to parse.

    Returns:
        int: Zero-based page index derived from the "_pNN" suffix, or 0 when no valid suffix is present.
    """
    if "_p" not in stem:
        return 0
    base, suffix = stem.rsplit("_p", 1)
    _ = base
    if suffix.isdigit():
        page_number = int(suffix)
        if page_number <= 0:
            return 0
        return page_number - 1
    return 0


def _export_sheet_pdf(
    sheet_api: _SheetApiProtocol,
    pdf_path: Path,
    *,
    ignore_print_areas: bool,
    print_area: str | None = None,
) -> None:
    """
    Export the given worksheet to a PDF file, optionally applying a temporary print area.

    If `print_area` is provided, it is applied to the sheet's PageSetup.PrintArea before exporting and restored afterwards. The function attempts to call ExportAsFixedFormat with an IgnorePrintAreas keyword; if that call fails due to an unexpected COM signature, it retries with a minimal argument set.

    Args:
        sheet_api: COM-like worksheet API exposing `PageSetup` and `ExportAsFixedFormat`.
        pdf_path (Path): Filesystem path to write the PDF to.
        ignore_print_areas (bool): If True, request that Excel ignore sheet print areas during export.
        print_area (str | None): Optional print area string to apply for this export; if None, the sheet's current print area is left unchanged.
    """
    original_print_area: object | None = None
    page_setup = None
    if print_area is not None:
        try:
            page_setup = getattr(sheet_api, "PageSetup", None)
            if page_setup is not None:
                original_print_area = getattr(page_setup, "PrintArea", None)
                page_setup.PrintArea = print_area
        except Exception as exc:
            logger.debug("Failed to set PrintArea. (%r)", exc)
            page_setup = None
    try:
        sheet_api.ExportAsFixedFormat(
            0, str(pdf_path), IgnorePrintAreas=ignore_print_areas
        )
    except TypeError:
        if ignore_print_areas:
            try:
                page_setup = page_setup or getattr(sheet_api, "PageSetup", None)
                if page_setup is not None:
                    page_setup.PrintArea = ""
            except Exception as exc:
                logger.debug("Failed to clear PrintArea for ignore. (%r)", exc)
        sheet_api.ExportAsFixedFormat(0, str(pdf_path))
    finally:
        if page_setup is not None and print_area is not None:
            try:
                page_setup.PrintArea = original_print_area
            except Exception as exc:
                logger.debug("Failed to restore PrintArea. (%r)", exc)


def _ensure_pdfium(use_subprocess: bool) -> ModuleType | None:
    """
    Ensure the pypdfium2 dependency is available and return the pdfium module for in-process rendering.

    Parameters:
        use_subprocess (bool): When True, confirm pypdfium2 is installed for subprocess rendering but do not keep the module in-process; when False, import and return the pdfium module for direct use.

    Returns:
        ModuleType | None: The imported `pdfium` module when `use_subprocess` is False, or `None` when `use_subprocess` is True.

    Raises:
        MissingDependencyError: If pypdfium2 (and required extras) is not installed.
    """
    if use_subprocess:
        _require_pdfium()
        return None
    return _require_pdfium()


def _export_sheet_images_with_app(
    excel_path: Path,
    output_dir: Path,
    temp_dir: Path,
    dpi: int,
    use_subprocess: bool,
    pdfium: ModuleType | None,
) -> list[Path]:
    """
    Export each worksheet of an Excel workbook to PNG images by exporting sheets to per-sheet PDFs and rendering those PDFs.

    Parameters:
        excel_path (Path): Path to the source Excel workbook.
        output_dir (Path): Directory where generated PNGs will be written.
        temp_dir (Path): Temporary directory for per-sheet intermediate PDF files.
        dpi (int): Dots per inch used when rasterizing PDF pages.
        use_subprocess (bool): If True, render PDF pages in a subprocess; otherwise render in-process.
        pdfium (ModuleType | None): In-process pypdfium2 module when rendering in-process, or None when subprocess rendering is used.

    Returns:
        list[Path]: Paths to generated PNG images in the order corresponding to the workbook's sheets and print-area splits.
    """
    written: list[Path] = []
    app: xw.App | None = None
    wb: xw.Book | None = None
    try:
        app = _require_excel_app()
        wb = app.books.open(str(excel_path))
        output_index = 0
        for sheet_name, sheet_api, print_area in _build_sheet_export_plan(wb):
            sheet_pdf = temp_dir / f"sheet_{output_index + 1:02d}.pdf"
            safe_name = _sanitize_sheet_filename(sheet_name)
            _export_sheet_pdf(
                sheet_api,
                sheet_pdf,
                ignore_print_areas=False,
                print_area=print_area,
            )
            sheet_paths = _render_sheet_images(
                pdfium,
                sheet_pdf,
                output_dir,
                output_index,
                safe_name,
                dpi,
                use_subprocess,
            )
            if not sheet_paths:
                _export_sheet_pdf(
                    sheet_api,
                    sheet_pdf,
                    ignore_print_areas=True,
                    print_area=print_area,
                )
                sheet_paths = _render_sheet_images(
                    pdfium,
                    sheet_pdf,
                    output_dir,
                    output_index,
                    safe_name,
                    dpi,
                    use_subprocess,
                )
            sheet_paths = _normalize_multipage_paths(
                sheet_paths,
                output_dir,
                output_index,
                safe_name,
            )
            written.extend(sheet_paths)
            output_index += max(1, len(sheet_paths))
        return written
    finally:
        if wb is not None:
            wb.close()
        if app is not None:
            app.quit()


def _render_sheet_images(
    pdfium: ModuleType | None,
    sheet_pdf: Path,
    output_dir: Path,
    output_index: int,
    safe_name: str,
    dpi: int,
    use_subprocess: bool,
) -> list[Path]:
    """
    Render a sheet PDF to one or more PNG files using either a subprocess or in-process renderer.

    Returns:
        paths (list[Path]): Paths to the generated PNG files in output order.

    Raises:
        RenderError: If in-process rendering is requested but the `pypdfium2` module (`pdfium`) is not provided.
    """
    if use_subprocess:
        return _render_pdf_pages_subprocess(
            sheet_pdf,
            output_dir,
            output_index,
            safe_name,
            dpi,
        )
    if pdfium is None:
        raise RenderError("pypdfium2 is required for in-process rendering.")
    return _render_pdf_pages_in_process(
        pdfium,
        sheet_pdf,
        output_dir,
        output_index,
        safe_name,
        dpi,
    )


def _normalize_multipage_paths(
    paths: list[Path],
    output_dir: Path,
    base_index: int,
    safe_name: str,
) -> list[Path]:
    """
    Assign distinct, ordered filenames for multi-page sheet outputs.

    If `paths` contains a single file, the list is returned unchanged. If `paths` contains multiple files, each file is given a unique, numbered filename in `output_dir` using `base_index` and `safe_name` so pages are ordered and do not collide.

    Parameters:
        paths (list[Path]): Existing file paths for a sheet's rendered pages.
        output_dir (Path): Directory containing or intended to contain the output files.
        base_index (int): Zero-based starting index used to compute numeric prefixes for filenames.
        safe_name (str): Filesystem-safe base name included in the generated filenames.

    Returns:
        list[Path]: Paths to the resulting files in `output_dir`. When multiple input paths are provided, returned paths reflect the new, uniquely prefixed filenames.
    """
    if len(paths) <= 1:
        return paths
    return _rename_pages_for_print_area(paths, output_dir, base_index, safe_name)


def _use_render_subprocess() -> bool:
    """
    Decide whether PDF-to-PNG rendering should be performed in a subprocess.

    Reads the environment variable EXSTRUCT_RENDER_SUBPROCESS (case-insensitive). Subprocess rendering is disabled when the variable is set to "0" or "false"; if the variable is unset or set to any other value, subprocess rendering is enabled.

    Returns:
        `true` if subprocess rendering is enabled, `false` otherwise.
    """
    return os.getenv("EXSTRUCT_RENDER_SUBPROCESS", "1").lower() not in {"0", "false"}


def _render_pdf_pages_in_process(
    pdfium: ModuleType,
    pdf_path: Path,
    output_dir: Path,
    sheet_index: int,
    safe_name: str,
    dpi: int,
) -> list[Path]:
    """Render PDF pages to PNGs in the current process."""
    scale = dpi / 72.0
    written: list[Path] = []
    with pdfium.PdfDocument(str(pdf_path)) as pdf:
        for page_index in range(len(pdf)):
            page = pdf[page_index]
            bitmap = page.render(scale=scale)
            pil_image = bitmap.to_pil()
            page_suffix = f"_p{page_index + 1:02d}" if page_index > 0 else ""
            img_path = (
                output_dir / f"{sheet_index + 1:02d}_{safe_name}{page_suffix}.png"
            )
            pil_image.save(img_path, format="PNG", dpi=(dpi, dpi))
            written.append(img_path)
    return written


def _render_pdf_pages_subprocess(
    pdf_path: Path,
    output_dir: Path,
    sheet_index: int,
    safe_name: str,
    dpi: int,
) -> list[Path]:
    """Render PDF pages to PNGs in a subprocess for memory isolation."""
    ctx = mp.get_context("spawn")
    queue: mp.Queue[dict[str, list[str] | str]] = ctx.Queue()
    process = ctx.Process(
        target=_render_pdf_pages_worker,
        args=(pdf_path, output_dir, sheet_index, safe_name, dpi, queue),
    )
    process.start()
    process.join()
    result = _get_subprocess_result(queue)
    if process.exitcode != 0 or "error" in result:
        message = result.get("error", "subprocess failed")
        raise RenderError(f"Failed to render PDF pages: {message}")
    paths = result.get("paths", [])
    return [Path(path) for path in paths]


def _get_subprocess_result(
    queue: mp.Queue[dict[str, list[str] | str]],
) -> dict[str, list[str] | str]:
    """Fetch the worker result from the queue with a timeout."""
    try:
        return queue.get(timeout=5)
    except Exception as exc:
        return {"error": f"subprocess did not return results ({exc})"}


def _render_pdf_pages_worker(
    pdf_path: Path,
    output_dir: Path,
    sheet_index: int,
    safe_name: str,
    dpi: int,
    queue: mp.Queue[dict[str, list[str] | str]],
) -> None:
    """Worker process to render PDF pages into PNG files."""
    try:
        import pypdfium2 as pdfium

        scale = dpi / 72.0
        output_dir.mkdir(parents=True, exist_ok=True)
        written: list[str] = []
        with pdfium.PdfDocument(str(pdf_path)) as pdf:
            for page_index in range(len(pdf)):
                page = pdf[page_index]
                bitmap = page.render(scale=scale)
                pil_image = bitmap.to_pil()
                page_suffix = f"_p{page_index + 1:02d}" if page_index > 0 else ""
                img_path = (
                    output_dir / f"{sheet_index + 1:02d}_{safe_name}{page_suffix}.png"
                )
                pil_image.save(img_path, format="PNG", dpi=(dpi, dpi))
                written.append(str(img_path))
        queue.put({"paths": written})
    except Exception as exc:
        queue.put({"error": str(exc)})


__all__ = ["export_pdf", "export_sheet_images"]
