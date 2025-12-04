from __future__ import annotations

from pathlib import Path
from typing import Literal, Optional

from .core.integrate import extract_workbook
from .io import save_as_json, save_as_toon, save_as_yaml, save_sheets
from .models import CellRow, Chart, ChartSeries, Shape, SheetData, WorkbookData
from .render import export_pdf, export_sheet_images

__all__ = [
    "extract",
    "export",
    "export_sheets",
    "export_pdf",
    "export_sheet_images",
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


def export(
    data: WorkbookData,
    path: str | Path,
    fmt: Optional[Literal["json", "yaml", "yml", "toon"]] = None,
) -> None:
    """Export WorkbookData to supported file formats (currently JSON)."""
    dest = Path(path)
    format_hint = (fmt or dest.suffix.lstrip(".") or "json").lower()
    match format_hint:
        case "json":
            save_as_json(data, dest)
        case "yaml" | "yml":
            save_as_yaml(data, dest)
        case "toon":
            save_as_toon(data, dest)
        case _:
            raise ValueError(f"Unsupported export format: {format_hint}")


def export_sheets(data: WorkbookData, dir_path: str | Path) -> dict[str, Path]:
    """
    Export each sheet as a JSON file.
    Payload includes book_name and the SheetData for that sheet.
    Returns a mapping of sheet name to written path.
    """
    return save_sheets(data, Path(dir_path), fmt="json")


def export_sheets_as(
    data: WorkbookData,
    dir_path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
) -> dict[str, Path]:
    """
    Export each sheet in the given format (json/yaml/toon).
    Payload includes book_name and the SheetData for that sheet.
    Returns a mapping of sheet name to written path.
    """
    return save_sheets(data, Path(dir_path), fmt=fmt)


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
    match out_fmt:
        case "json":
            save_as_json(workbook_model, output_path)
        case "yaml" | "yml":
            save_as_yaml(workbook_model, output_path)
        case "toon":
            save_as_toon(workbook_model, output_path)
        case _:
            raise ValueError(f"Unsupported export format: {out_fmt}")

    if pdf or image:
        pdf_path = output_path.with_suffix(".pdf")
        export_pdf(file_path, pdf_path)
        if image:
            images_dir = output_path.parent / f"{output_path.stem}_images"
            export_sheet_images(file_path, images_dir, dpi=dpi)

    print(f"{file_path.name} -> {output_path} done.")
