from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict

import toon
import yaml

from ..models import WorkbookData


def dict_without_empty_values(obj: Any):
    """Recursively drop empty values from nested structures."""
    if isinstance(obj, dict):
        return {
            k: dict_without_empty_values(v)
            for k, v in obj.items()
            if v not in [None, "", [], {}]
        }
    if isinstance(obj, list):
        return [
            dict_without_empty_values(v) for v in obj if v not in [None, "", [], {}]
        ]
    if hasattr(obj, "model_dump"):
        return dict_without_empty_values(obj.model_dump(exclude_none=True))
    return obj


def save_as_json(model: WorkbookData, path: Path) -> None:
    filtered_dict = dict_without_empty_values(model)
    path.write_text(
        json.dumps(filtered_dict, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def save_as_yaml(model: WorkbookData, path: Path) -> None:
    filtered_dict = dict_without_empty_values(model)
    text = yaml.safe_dump(
        filtered_dict,
        allow_unicode=True,
        sort_keys=False,
        indent=2,
    )
    path.write_text(text, encoding="utf-8")


def save_as_toon(model: WorkbookData, path: Path) -> None:
    filtered_dict = dict_without_empty_values(model)
    text = toon.encode(filtered_dict)
    path.write_text(text, encoding="utf-8")


def _sanitize_sheet_filename(name: str) -> str:
    """Make a sheet name safe for filesystem usage."""
    safe = re.sub(r"[\\/:*?\"<>|]", "_", name)
    return safe or "sheet"


def save_sheets_as_json(workbook: WorkbookData, output_dir: Path) -> Dict[str, Path]:
    """
    Save each sheet as an individual JSON file.
    Contents include book_name and the sheet's SheetData.
    Returns a map of sheet name -> written path.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    written: Dict[str, Path] = {}
    for sheet_name, sheet_data in workbook.sheets.items():
        payload = dict_without_empty_values(
            {
                "book_name": workbook.book_name,
                "sheet_name": sheet_name,
                "sheet": sheet_data,
            }
        )
        file_name = f"{_sanitize_sheet_filename(sheet_name)}.json"
        path = output_dir / file_name
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        written[sheet_name] = path
    return written


__all__ = [
    "dict_without_empty_values",
    "save_as_json",
    "save_as_yaml",
    "save_as_toon",
    "save_sheets_as_json",
]
