# JSON Schemas for ExStruct Models

This page lists the JSON Schemas generated from the public Pydantic models. The
schema files live at the repository root under `schemas/` (one file per model).
They are not bundled into the MkDocs site build; clone or download the
repository to access the raw files.

## Available Schemas

- `schemas/workbook.json` — `WorkbookData`
- `schemas/sheet.json` — `SheetData`
- `schemas/cell_row.json` — `CellRow`
- `schemas/shape.json` — `Shape`
- `schemas/chart.json` — `Chart`
- `schemas/chart_series.json` — `ChartSeries`
- `schemas/print_area.json` — `PrintArea`
- `schemas/print_area_view.json` — `PrintAreaView`

## Regeneration

After modifying models, regenerate the schemas to keep them in sync:

```bash
python scripts/gen_json_schema.py
```

The generator produces deterministic output (sorted keys, draft 2020-12
`$schema`).
