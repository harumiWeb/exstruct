from __future__ import annotations

import logging
import multiprocessing as mp
import os
from pathlib import Path
import shutil
import tempfile
from types import ModuleType
from typing import Any, cast

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
    """Export each sheet as PNG (via PDF then pypdfium2 rasterization) and return paths in sheet order."""
    normalized_excel_path = Path(excel_path)
    normalized_output_dir = Path(output_dir)
    normalized_output_dir.mkdir(parents=True, exist_ok=True)
    use_subprocess = _use_render_subprocess()
    if not use_subprocess:
        pdfium = cast(Any, _require_pdfium())
    else:
        _require_pdfium()

    try:
        with tempfile.TemporaryDirectory() as td:
            written: list[Path] = []
            app: xw.App | None = None
            wb: xw.Book | None = None
            try:
                app = _require_excel_app()
                wb = app.books.open(str(normalized_excel_path))
                for sheet_index, sheet in enumerate(wb.sheets):
                    sheet_name = sheet.name
                    sheet_pdf = Path(td) / f"sheet_{sheet_index + 1:02d}.pdf"
                    sheet.api.ExportAsFixedFormat(0, str(sheet_pdf))
                    safe_name = _sanitize_sheet_filename(sheet_name)
                    if use_subprocess:
                        written.extend(
                            _render_pdf_pages_subprocess(
                                sheet_pdf,
                                normalized_output_dir,
                                sheet_index,
                                safe_name,
                                dpi,
                            )
                        )
                    else:
                        written.extend(
                            _render_pdf_pages_in_process(
                                pdfium,
                                sheet_pdf,
                                normalized_output_dir,
                                sheet_index,
                                safe_name,
                                dpi,
                            )
                        )
                return written
            finally:
                if wb is not None:
                    wb.close()
                if app is not None:
                    app.quit()
    except RenderError:
        raise
    except Exception as exc:
        raise RenderError(
            f"Failed to export sheet images to '{normalized_output_dir}'."
        ) from exc


def _sanitize_sheet_filename(name: str) -> str:
    return "".join("_" if c in '\\/:*?"<>|' else c for c in name).strip() or "sheet"


def _use_render_subprocess() -> bool:
    """Return True when PDF->PNG rendering should run in a subprocess."""
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
