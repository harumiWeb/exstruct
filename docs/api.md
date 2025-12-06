# API Reference

This page shows the primary APIs, minimal runnable examples, expected outputs, and the dependencies required for optional features. Hyperlinks are included when `include_cell_links=True` (or when using `mode="verbose"`).

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
      "rows": [{"r": 1, "c": {"0": "Name", "1": "Age"}, "links": null}],
      "shapes": [{"text": "note", "l": 10, "t": 20, "w": 80, "h": 24, "type": "TextBox"}],
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
- YAML export: `pyyaml` (imported lazily; missing module raises `RuntimeError`).
- TOON export: `python-toon` (lazy import; missing module raises `RuntimeError`).
- Rendering (PDF/PNG): **Excel + COM + `pypdfium2`** are mandatory. Without Excel/COM, rendering APIs raise `RuntimeError`.

## Functions

### extract(file_path, mode="standard")

Extracts workbook structure. Modes: `light`, `standard` (default), `verbose`. Raises `ValueError` on invalid mode. Hyperlinks are included when `mode="verbose"`; for other modes pass `include_cell_links=True` via `StructOptions` (see engine below).

```python
from exstruct import extract

wb = extract("input.xlsx", mode="verbose")  # includes cell hyperlinks in rows[*].links
print(wb.sheets["Sheet1"].table_candidates)
```

### export(data, path, fmt=None, *, pretty=False, indent=None)

Exports `WorkbookData` to JSON/YAML/TOON. `fmt` defaults from path suffix or `json`. Raises `ValueError` on unsupported fmt. JSON is compact unless `pretty=True` or `indent` is set.

```python
export(wb, "out.yaml", fmt="yaml")  # requires pyyaml
export(wb, "out.json", pretty=True)  # pretty-printed JSON
```

### export_sheets(data, dir_path)

Writes one JSON per sheet, including `book_name` and the `SheetData`. Returns `{sheet_name: Path}`.

```python
paths = export_sheets(wb, "out_dir")
print(paths["Sheet1"])  # out_dir/Sheet1.json
```

### export_sheets_as(data, dir_path, fmt="json", *, pretty=False, indent=None)

Same as `export_sheets` but supports `json`/`yaml`/`yml`/`toon`. Raises `ValueError` on invalid fmt.

```python
export_sheets_as(wb, "yaml_dir", fmt="yaml")  # requires pyyaml
```

### process_excel(file_path, output_path=None, out_fmt="json", image=False, pdf=False, dpi=72, mode="standard", pretty=False, indent=None, sheets_dir=None, stream=None)

Convenience wrapper used by the CLI. Writes to stdout when `output_path` is omitted, can optionally split per sheet (`sheets_dir`), and can render PDF/PNG (Excel required). Invalid `mode` or `out_fmt` raises `ValueError`.

### export_pdf(file_path, pdf_path)

Exports the workbook to PDF via Excel COM. Requires Excel and `pypdfium2` if images will be generated afterwards.

### export_sheet_images(file_path, images_dir, dpi=72)

Converts a PDF (rendered via Excel) into per-sheet PNGs using `pypdfium2`. Requires Excel + COM + `pypdfium2`.

### set_table_detection_params(...)

Tunes table detection heuristics at runtime. Higher thresholds reduce false positives; lower thresholds catch faint tables.

```python
from exstruct import set_table_detection_params

# Before: a layout frame was misdetected as a table.
set_table_detection_params(table_score_threshold=0.35, density_min=0.05)

# After lowering thresholds to catch very sparse tables:
set_table_detection_params(table_score_threshold=0.25, coverage_min=0.15)
```

### ExStructEngine(options=StructOptions(), output=OutputOptions())

Configurable engine for per-instance extraction/output settings.

- `StructOptions`: `mode`, optional `table_params` (forwarded to `set_table_detection_params`), `include_cell_links` (None -> auto: verbose=True, others=False).
- `OutputOptions`: defaults for `fmt`, `pretty`, `indent`, include/exclude flags for rows/shapes/charts/tables, `sheets_dir`, `stream`.
- Methods:
  - `extract(path, mode=None)` → WorkbookData
  - `serialize(workbook, fmt=None, pretty=None, indent=None)` → str (filters applied)
  - `export(workbook, output_path=None, fmt=None, pretty=None, indent=None, sheets_dir=None, stream=None)`
  - `process(file_path, output_path=None, out_fmt=None, image=False, pdf=False, dpi=72, mode=None, pretty=None, indent=None, sheets_dir=None, stream=None)`

Example:

```python
from exstruct import ExStructEngine, StructOptions, OutputOptions

engine = ExStructEngine(
    options=StructOptions(mode="standard", include_cell_links=True),  # enable hyperlinks in standard mode
    output=OutputOptions(include_shapes=False, pretty=True),
)
wb = engine.extract("input.xlsx")
engine.export(wb, "out.json")              # writes filtered JSON (no shapes)
engine.process("input.xlsx", pdf=False)    # end-to-end extract + export
```

## Models

| Model         | Key fields                                                                                 |
| ------------- | ------------------------------------------------------------------------------------------ |
| `WorkbookData`| `book_name: str`, `sheets: dict[str, SheetData]`                                           |
| `SheetData`   | `rows: list[CellRow]`, `shapes: list[Shape]`, `charts: list[Chart]`, `table_candidates: list[str]` |
| `CellRow`     | `r: int`, `c: dict[str, int | float | str]`, `links: dict[str, str] \| None`               |
| `Shape`       | `text: str`, `l/t/w/h: int|None`, `type`, `rotation`, arrow styles, `direction`|
| `Chart`       | `name`, `chart_type`, `title`, `series`, `y_axis_range`, `l/t`, `error: str|None`          |
| `ChartSeries` | `name`, `name_range`, `x_range`, `y_range`                                                 |

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

- Excel COM unavailable: extraction falls back to cells + `table_candidates`; `shapes`/`charts` are empty, warning is logged.
- Invalid `fmt` or `mode`: `ValueError`.
- Missing optional dependency (`pyyaml`, `python-toon`, `pypdfium2`): `RuntimeError` with install hint.
- Rendering without Excel/COM: `RuntimeError`.
- CLI mirrors these: exits non-zero on failures, prints messages in English.

## Tuning Examples

- Reduce false positives (layout frames):

```python
set_table_detection_params(table_score_threshold=0.4, coverage_min=0.25)
```

- Recover missed tiny tables:

```python
set_table_detection_params(density_min=0.03, min_nonempty_cells=2)
```
