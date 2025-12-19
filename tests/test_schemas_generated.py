from __future__ import annotations

from collections.abc import Mapping
import json
from pathlib import Path
from typing import Any


def _load_schema(name: str) -> dict[str, Any]:
    schema_path = Path("schemas") / name
    assert schema_path.exists(), f"Missing schema file: {schema_path}"
    loaded = json.loads(schema_path.read_text(encoding="utf-8"))
    assert isinstance(loaded, dict), f"Schema is not an object: {name}"
    return loaded


def test_workbook_schema_has_required_metadata() -> None:
    schema = _load_schema("workbook.json")
    assert "$schema" in schema
    assert schema.get("title") == "WorkbookData"


def test_sheet_schema_definitions_present() -> None:
    schema = _load_schema("sheet.json")
    # Ensure nested models are included as definitions for downstream use.
    assert "$defs" in schema
    defs = schema["$defs"]
    assert isinstance(defs, Mapping)
    assert "CellRow" in defs
    assert "PrintArea" in defs
