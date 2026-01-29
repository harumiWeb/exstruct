from __future__ import annotations

from pathlib import Path

import fitz  # PyMuPDF

from .common import ensure_dir
from .pdf_text import xlsx_to_pdf


def xlsx_to_pngs_via_pdf(
    xlsx_path: Path, out_dir: Path, dpi: int = 200, max_pages: int = 6
) -> list[Path]:
    """
    xlsx -> pdf (LibreOffice) -> png (PyMuPDF render)
    画像は VLM 入力に使う。OCRはしない。
    """
    ensure_dir(out_dir)
    tmp_pdf = out_dir / f"{xlsx_path.stem}.pdf"
    xlsx_to_pdf(xlsx_path, tmp_pdf)

    paths: list[Path] = []
    with fitz.open(tmp_pdf) as doc:
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        for i in range(min(doc.page_count, max_pages)):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            p = out_dir / f"page_{i + 1:02d}.png"
            pix.save(p)
            paths.append(p)

    return paths
