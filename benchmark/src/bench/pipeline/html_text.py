from __future__ import annotations

from pathlib import Path
import subprocess

from bs4 import BeautifulSoup

from .common import ensure_dir, write_text


def xlsx_to_html(xlsx_path: Path, out_html: Path) -> None:
    ensure_dir(out_html.parent)
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--convert-to",
        "html",
        "--outdir",
        str(out_html.parent),
        str(xlsx_path),
    ]
    subprocess.run(cmd, check=True)
    produced = out_html.parent / (xlsx_path.stem + ".html")
    if not produced.exists():
        produced = out_html.parent / (xlsx_path.stem + ".htm")
    produced.replace(out_html)


def html_to_text(html_path: Path, out_txt: Path) -> None:
    soup = BeautifulSoup(html_path.read_text(encoding="utf-8", errors="ignore"), "lxml")

    # Excel HTMLはテーブルが中心。全テーブルのセルテキストを列挙。
    tables = soup.find_all("table")
    lines: list[str] = []
    lines.append("[DOC_META]")
    lines.append(f"source={html_path.name}")
    lines.append("method=html_text")
    lines.append("")
    lines.append("[CONTENT]")

    for t_i, table in enumerate(tables, start=1):
        lines.append(f"\n# TABLE {t_i}")
        rows = table.find_all("tr")
        for r in rows:
            cells = r.find_all(["td", "th"])
            vals = []
            for c in cells:
                txt = " ".join(c.get_text(separator=" ", strip=True).split())
                vals.append(txt)
            if any(v for v in vals):
                lines.append(" | ".join(vals))

    write_text(out_txt, "\n".join(lines).strip() + "\n")
