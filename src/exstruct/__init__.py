from __future__ import annotations

from pathlib import Path
from typing import Optional

from .core.integrate import extract_workbook
from .io import save_as_json, save_as_toon, save_as_yaml, save_sheets_as_json
from .models import CellRow, Chart, ChartSeries, Shape, SheetData, WorkbookData

__all__ = [
    "extract",
    "export",
    "export_sheets",
    "process_excel",
    "CellRow",
    "Shape",
    "ChartSeries",
    "Chart",
    "SheetData",
    "WorkbookData",
]


def extract(file_path: str | Path) -> WorkbookData:
    """High-level API entrypoint that extracts workbook semantic structure."""
    return extract_workbook(Path(file_path))


def export(data: WorkbookData, path: str | Path, fmt: Optional[str] = None) -> None:
    """Export WorkbookData to supported file formats (currently JSON)."""
    dest = Path(path)
    format_hint = (fmt or dest.suffix.lstrip(".") or "json").lower()
    if format_hint == "json":
        save_as_json(data, dest)
    elif format_hint in ("yaml", "yml"):
        save_as_yaml(data, dest)
    elif format_hint == "toon":
        save_as_toon(data, dest)
    else:
        raise ValueError(f"Unsupported export format: {format_hint}")


def export_sheets(data: WorkbookData, dir_path: str | Path) -> dict[str, Path]:
    """
    Export each sheet as a JSON file.
    Payload includes book_name and the SheetData for that sheet.
    Returns a mapping of sheet name to written path.
    """
    return save_sheets_as_json(data, Path(dir_path))


def process_excel(
    file_path: Path,
    output_path: Path,
    out_fmt: str = "json",
    image: bool = False,
    pdf: bool = False,
    dpi: int = 72,
) -> None:
    """Compatibility wrapper used by CLI prototypes."""
    workbook_model = extract(file_path)
    if out_fmt == "json":
        save_as_json(workbook_model, output_path)
    print(f"{file_path.name} -> {output_path} 完了")
