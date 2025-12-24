# API Reference

This page shows the primary APIs, minimal runnable examples, expected outputs, and the dependencies required for optional features. Hyperlinks are included when `include_cell_links=True` (or when using `mode="verbose"`). Auto page-break areas are COM-only and appear when auto page-break extraction/output is enabled (CLI exposes the option only when COM is available).

## TOC

<!-- TOC -->

- [API Reference](#api-reference)
  - [TOC](#toc)
  - [Quick Examples](#quick-examples)
  - [Dependencies](#dependencies)
  - [Auto-generated API mkdocstrings](#auto-generated-api-mkdocstrings)
    - [Core functions](#core-functions)
    - [Engine and options](#engine-and-options)
  - [Models](#models)
    - [Model helpers SheetData / WorkbookData](#model-helpers-sheetdata--workbookdata)
  - [Error Handling](#error-handling)
  - [Tuning Examples](#tuning-examples)

<!-- /TOC -->

## Quick Examples

```python
from exstruct import extract, export

wb = extract("sample.xlsx", mode="standard")
export(wb, "out.json")  # compact JSON by default
```

Expected JSON snippet (links appear when enabled):

```json
{
  "book_name": "sample.xlsx",
  "sheets": {
    "Sheet1": {
      "rows": [{ "r": 1, "c": { "0": "Name", "1": "Age" }, "links": null }],
      "shapes": [
        {
          "text": "note",
          "l": 10,
          "t": 20,
          "w": 80,
          "h": 24,
          "type": "TextBox"
        }
      ],
      "charts": [],
      "table_candidates": ["A1:B5"]
    }
  }
}
```

CLI-equivalent flow via Python:

```python
from pathlib import Path
from exstruct import process_excel

process_excel(
    file_path=Path("input.xlsx"),
    output_path=None,  # default: stdout (redirect if you want a file)
    sheets_dir=Path("out_sheets"),  # optional per-sheet outputs
    out_fmt="json",
    image=True,
    pdf=True,
    mode="standard",
    pretty=True,
)
# Same as: exstruct input.xlsx --format json --pdf --image --mode standard --pretty --sheets-dir out_sheets > out.json
```

## Dependencies

- Core extraction: pandas, openpyxl (installed with the package).
- YAML export: `pyyaml` (lazy import; missing module raises `MissingDependencyError`).
- TOON export: `python-toon` (lazy import; missing module raises `MissingDependencyError`).
- Auto page-break extraction/export: **Excel + COM** required (feature is skipped when COM is unavailable; CLI hides the flag in non-COM environments).
- Rendering (PDF/PNG): **Excel + COM + `pypdfium2`** are mandatory. Missing Excel/COM or `pypdfium2` surfaces as `RenderError`/`MissingDependencyError`.

## Auto-generated API (mkdocstrings)

Python APIの最新情報は以下の自動生成セクションを参照してください（docstringベースで同期）。

### Core functions

::: exstruct.extract
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.export
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.export_sheets
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.export_sheets_as
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.export_print_areas_as
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.export_auto_page_breaks
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.export_pdf
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.export_sheet_images
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.process_excel
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

### Engine and options

::: exstruct.engine.ExStructEngine
    handler: python
    options:
      show_signature_annotations: true
      members_order: source
      show_root_heading: true

::: exstruct.engine.StructOptions
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.engine.OutputOptions
    handler: python
    options:
      show_signature_annotations: true
      members_order: source
      show_root_heading: true

::: exstruct.engine.FormatOptions
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.engine.FilterOptions
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.engine.DestinationOptions
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

::: exstruct.engine.ColorsOptions
    handler: python
    options:
      show_signature_annotations: true
      show_root_heading: true

## Models

See generated/models.md for the detailed model fields (run `python scripts/gen_model_docs.py` to refresh).

### Model helpers (SheetData / WorkbookData)

- `to_json(pretty=False, indent=None)` → JSON string (pretty when requested)
- `to_yaml()` → YAML string (requires `pyyaml`)
- `to_toon()` → TOON string (requires `python-toon`)
- `save(path, pretty=False, indent=None)` → infers format from suffix (`.json/.yaml/.yml/.toon`)
- `WorkbookData.__getitem__(name)` → get a SheetData by name
- `WorkbookData.__iter__()` → yields `(sheet_name, SheetData)` in order

Example:

```python
wb = extract("input.xlsx")
first = wb["Sheet1"]
for name, sheet in wb:
    print(name, len(sheet.rows))
wb.save("out.json", pretty=True)
first.save("sheet.yaml")  # requires pyyaml
```

## Error Handling

- Exception types:
  - `SerializationError`: Unsupported format requested (`serialize_workbook`, export APIs).
  - `MissingDependencyError`: Optional dependency (`pyyaml` / `python-toon` / `pypdfium2`) is missing; message includes install instructions.
  - `RenderError`: Excel/COM is unavailable or PDF/PNG rendering fails.
  - `PrintAreaError` (ValueError-compatible): `export_auto_page_breaks` invoked when no `auto_print_areas` are available.
  - `OutputError`: Writing output to disk/stream failed (original exception kept in `__cause__`).
  - `ValueError`: Invalid inputs such as an unsupported `mode`.
- Excel COM unavailable: extraction falls back to cells + `table_candidates`; `shapes`/`charts` are empty, warning is logged.
- No print areas: `export_print_areas_as` writes nothing and returns `{}`; this is not an error.
- Auto page-break export: `export_auto_page_breaks` raises `PrintAreaError` if no auto page-break areas are present (enable them via `DestinationOptions.auto_page_breaks_dir`).
- CLI mirrors these behaviors: exits non-zero on failures, prints messages in English.

## Tuning Examples

- Reduce false positives (layout frames):

```python
set_table_detection_params(table_score_threshold=0.4, coverage_min=0.25)
```

- Recover missed tiny tables:

```python
set_table_detection_params(density_min=0.03, min_nonempty_cells=2)
```
