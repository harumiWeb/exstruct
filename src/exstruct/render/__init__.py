from __future__ import annotations

import logging
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
        shutil.copy(normalized_excel_path, temp_xlsx)

        app: xw.App | None = None
        wb: xw.Book | None = None
        try:
            app = _require_excel_app()
            wb = app.books.open(str(temp_xlsx))
            sheet_names = [s.name for s in wb.sheets]
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
    pdfium = cast(Any, _require_pdfium())
    normalized_excel_path = Path(excel_path)
    normalized_output_dir = Path(output_dir)
    normalized_output_dir.mkdir(parents=True, exist_ok=True)

    try:
        with tempfile.TemporaryDirectory() as td:
            tmp_pdf = Path(td) / "book.pdf"
            sheet_names = export_pdf(normalized_excel_path, tmp_pdf)

            scale = dpi / 72.0
            written: list[Path] = []
            with pdfium.PdfDocument(str(tmp_pdf)) as pdf:
                for i, sheet_name in enumerate(sheet_names):
                    page = pdf[i]
                    bitmap = page.render(scale=scale)
                    pil_image = bitmap.to_pil()
                    safe_name = _sanitize_sheet_filename(sheet_name)
                    img_path = normalized_output_dir / f"{i + 1:02d}_{safe_name}.png"
                    pil_image.save(img_path, format="PNG", dpi=(dpi, dpi))
                    written.append(img_path)
            return written
    except RenderError:
        raise
    except Exception as exc:
        raise RenderError(
            f"Failed to export sheet images to '{normalized_output_dir}'."
        ) from exc


def _sanitize_sheet_filename(name: str) -> str:
    return "".join("_" if c in '\\/:*?"<>|' else c for c in name).strip() or "sheet"


__all__ = ["export_pdf", "export_sheet_images"]
