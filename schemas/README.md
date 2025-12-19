# ExStruct JSON Schemas

This directory contains JSON Schemas for the public Pydantic models of ExStruct,
one file per model (e.g., `workbook.json`, `sheet.json`, `print_area.json`).

## Regenerating

Run the generator after model changes to keep the schemas in sync:

```bash
python scripts/gen_json_schema.py
```

The script uses Pydantic's `model_json_schema()` and writes deterministic output
with sorted keys and draft 2020-12 `$schema` metadata.
