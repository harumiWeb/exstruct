from __future__ import annotations

import subprocess
from pathlib import Path

import fitz  # PyMuPDF

from .common import ensure_dir, write_text


def xlsx_to_pdf(xlsx_path: Path, out_pdf: Path) -> None:
    ensure_dir(out_pdf.parent)
    # LibreOffice headless convert
    # soffice --headless --convert-to pdf --outdir <dir> <xlsx>
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_pdf.parent),
        str(xlsx_path),
    ]
    try:
        subprocess.run(cmd, check=True, timeout=300)
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(f"soffice timed out after 300s: {xlsx_path}") from exc
    produced = out_pdf.parent / (xlsx_path.stem + ".pdf")
    produced.replace(out_pdf)


def pdf_to_text(pdf_path: Path, out_txt: Path) -> None:
    parts: list[str] = []
    with fitz.open(pdf_path) as doc:
        for i in range(doc.page_count):
            page = doc.load_page(i)
            parts.append(f"\n# PAGE {i + 1}")
            parts.append(page.get_text("text"))
    text = "\n".join(parts).strip()

    lines: list[str] = []
    lines.append("[DOC_META]")
    lines.append(f"source={pdf_path.name}")
    lines.append("method=pdf_text")
    lines.append("")
    lines.append("[CONTENT]")
    lines.append(text)

    write_text(out_txt, "\n".join(lines).strip() + "\n")
