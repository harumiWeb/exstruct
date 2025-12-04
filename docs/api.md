# API Reference

This page shows the primary APIs, minimal runnable examples, expected outputs, and the dependencies required for optional features.

## Quick Examples

```python
from exstruct import extract, export

wb = extract("sample.xlsx", mode="standard")
export(wb, "out.json")  # compact JSON by default
```

Expected JSON snippet:

```json
{
  "book_name": "sample.xlsx",
  "sheets": {
    "Sheet1": {
      "rows": [{"r": 1, "c": {"0": "Name", "1": "Age"}}],
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
    output_path=Path("out.json"),
    out_fmt="json",
    image=True,
    pdf=True,
    mode="standard",
    pretty=True,
)
# Same as: exstruct input.xlsx --format json --pdf --image --mode standard --pretty
```

## Dependencies

- Core extraction: pandas, openpyxl (installed with the package).
- YAML export: `pyyaml` (imported lazily; missing module raises `RuntimeError`).
- TOON export: `python-toon` (lazy import; missing module raises `RuntimeError`).
- Rendering (PDF/PNG): **Excel + COM + `pypdfium2`** are mandatory. Without Excel/COM, rendering APIs raise `RuntimeError`.

## Functions

### extract(file_path, mode="standard")

Extracts workbook structure. Modes: `light`, `standard` (default), `verbose`. Raises `ValueError` on invalid mode.

```python
from exstruct import extract

wb = extract("input.xlsx", mode="verbose")
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

### process_excel(file_path, output_path, out_fmt="json", image=False, pdf=False, dpi=72, mode="standard", pretty=False, indent=None)

Convenience wrapper used by the CLI. Writes the chosen format and optionally PDF/PNG (Excel required). Invalid `mode` or `out_fmt` raises `ValueError`.

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

## Models

| Model         | Key fields                                                                                 |
| ------------- | ------------------------------------------------------------------------------------------ |
| `WorkbookData`| `book_name: str`, `sheets: dict[str, SheetData]`                                           |
| `SheetData`   | `rows: list[CellRow]`, `shapes: list[Shape]`, `charts: list[Chart]`, `table_candidates: list[str]` |
| `CellRow`     | `r: int`, `c: dict[str, int | float | str]`                                                |
| `Shape`       | `text: str`, `l/t/w/h: int|None`, `type`, `rotation`, arrow styles, `direction`|
| `Chart`       | `name`, `chart_type`, `title`, `series`, `y_axis_range`, `l/t`, `error: str|None`          |
| `ChartSeries` | `name`, `name_range`, `x_range`, `y_range`                                                 |

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
