from __future__ import annotations

import json
from pathlib import Path

from pydantic import BaseModel

from exstruct.models import (
    CellRow,
    Chart,
    ChartSeries,
    PrintArea,
    PrintAreaView,
    Shape,
    SheetData,
    WorkbookData,
)

_SCHEMA_VERSION = "https://json-schema.org/draft/2020-12/schema"


def _model_schema(model: type[BaseModel]) -> dict[str, object]:
    """Build a JSON Schema for a Pydantic model with deterministic ordering."""
    schema = model.model_json_schema()
    schema.setdefault("$schema", _SCHEMA_VERSION)
    return schema


def _write_schema(name: str, model: type[BaseModel], output_dir: Path) -> Path:
    schema = _model_schema(model)
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"{name}.json"
    text = json.dumps(schema, ensure_ascii=False, indent=2, sort_keys=True)
    path.write_text(text, encoding="utf-8")
    return path


def main() -> int:
    """Generate JSON Schemas for ExStruct public models."""
    project_root = Path(__file__).resolve().parent.parent
    output_dir = project_root / "schemas"
    targets: dict[str, type[BaseModel]] = {
        "workbook": WorkbookData,
        "sheet": SheetData,
        "cell_row": CellRow,
        "shape": Shape,
        "chart": Chart,
        "chart_series": ChartSeries,
        "print_area": PrintArea,
        "print_area_view": PrintAreaView,
    }

    for name, model in targets.items():
        _write_schema(name, model, output_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
