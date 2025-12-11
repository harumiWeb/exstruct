# API Reference

This page shows the primary APIs, minimal runnable examples, expected outputs, and the dependencies required for optional features. Hyperlinks are included when `include_cell_links=True` (or when using `mode="verbose"`). Auto page-break areas are COM-only and appear when auto page-break extraction/output is enabled.

## TOC

<!-- TOC -->

- [API Reference](#api-reference)
  - [TOC](#toc)
  - [Quick Examples](#quick-examples)
  - [Dependencies](#dependencies)
  - [Functions](#functions)
    - [extractfile_path, mode="standard"](#extractfile_path-modestandard)
    - [exportdata, path, fmt=None, *, pretty=False, indent=None](#exportdata-path-fmtnone--prettyfalse-indentnone)
    - [export_sheetsdata, dir_path](#export_sheetsdata-dir_path)
    - [export_sheets_asdata, dir_path, fmt="json", *, pretty=False, indent=None](#export_sheets_asdata-dir_path-fmtjson--prettyfalse-indentnone)
    - [export_print_areas_asdata, dir_path, fmt="json", *, pretty=False, indent=None, normalize=False](#export_print_areas_asdata-dir_path-fmtjson--prettyfalse-indentnone-normalizefalse)
    - [export_auto_page_breaksdata, dir_path, fmt="json", *, pretty=False, indent=None, normalize=False](#export_auto_page_breaksdata-dir_path-fmtjson--prettyfalse-indentnone-normalizefalse)
    - [process_excelfile_path, output_path=None, out_fmt="json", image=False, pdf=False, dpi=72, mode="standard", pretty=False, indent=None, sheets_dir=None, print_areas_dir=None, auto_page_breaks_dir=None, stream=None](#process_excelfile_path-output_pathnone-out_fmtjson-imagefalse-pdffalse-dpi72-modestandard-prettyfalse-indentnone-sheets_dirnone-print_areas_dirnone-auto_page_breaks_dirnone-streamnone)
    - [export_pdffile_path, pdf_path](#export_pdffile_path-pdf_path)
    - [export_sheet_imagesfile_path, images_dir, dpi=72](#export_sheet_imagesfile_path-images_dir-dpi72)
    - [set_table_detection_params...](#set_table_detection_params)
    - [ExStructEngineoptions=StructOptions, output=OutputOptions](#exstructengineoptionsstructoptions-outputoutputoptions)
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
- YAML export: `pyyaml` (imported lazily; missing module raises `RuntimeError`).
- TOON export: `python-toon` (lazy import; missing module raises `RuntimeError`).
- Auto page-break extraction/export: **Excel + COM** required (feature is skipped when COM is unavailable).
- Rendering (PDF/PNG): **Excel + COM + `pypdfium2`** are mandatory. Without Excel/COM, rendering APIs raise `RuntimeError`.

## Functions

### extract(file_path, mode="standard")

Extracts workbook structure. Modes: `light`, `standard` (default), `verbose`. Raises `ValueError` on invalid mode. Hyperlinks are included when `mode="verbose"`; for other modes pass `include_cell_links=True` via `StructOptions` (see engine below).

```python
from exstruct import extract

wb = extract("input.xlsx", mode="verbose")  # includes cell hyperlinks in rows[*].links
print(wb.sheets["Sheet1"].table_candidates)
```

### export(data, path, fmt=None, \*, pretty=False, indent=None)

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

### export_sheets_as(data, dir_path, fmt="json", \*, pretty=False, indent=None)

Same as `export_sheets` but supports `json`/`yaml`/`yml`/`toon`. Raises `ValueError` on invalid fmt.

```python
export_sheets_as(wb, "yaml_dir", fmt="yaml")  # requires pyyaml
```

### export_print_areas_as(data, dir_path, fmt="json", \*, pretty=False, indent=None, normalize=False)

Writes one file per print area as `PrintAreaView`. If no print areas exist, returns an empty dict and writes nothing.

```python
from exstruct import export_print_areas_as
paths = export_print_areas_as(wb, "areas", fmt="json", pretty=True)  # only when print areas exist
```

Args:
- data: WorkbookData containing print areas
- dir_path: output directory
- fmt: json/yaml/yml/toon
- pretty/indent: JSON formatting options
- normalize: rebase row/col indices to the area origin

Returns: dict of area key -> Path (e.g., `"Sheet1#1": areas/Sheet1_area1_...json`)

### export_auto_page_breaks(data, dir_path, fmt="json", \*, pretty=False, indent=None, normalize=False)

Writes one file per auto page-break area (`SheetData.auto_print_areas`). Requires COM-based extraction with auto page-breaks enabled. Raises `ValueError` if no auto page breaks exist.

```python
from exstruct import export_auto_page_breaks
paths = export_auto_page_breaks(wb, "auto_areas", fmt="json", pretty=True)  # COM + auto breaks enabled
```

Args:
- data: WorkbookData containing auto page-break areas
- dir_path: output directory
- fmt: json/yaml/yml/toon
- pretty/indent: JSON formatting options
- normalize: rebase row/col indices to the area origin

Returns: dict of area key -> Path (e.g., `"Sheet1#1": auto_areas/Sheet1_auto_page1_...json`)

### process_excel(file_path, output_path=None, out_fmt="json", image=False, pdf=False, dpi=72, mode="standard", pretty=False, indent=None, sheets_dir=None, print_areas_dir=None, auto_page_breaks_dir=None, stream=None)

Convenience wrapper used by the CLI. Writes to stdout when `output_path` is omitted, can optionally split per sheet (`sheets_dir`), split per print area (`print_areas_dir`), split per auto page-break area (`auto_page_breaks_dir`, COM only), and can render PDF/PNG (Excel required). Invalid `mode` or `out_fmt` raises `ValueError`.

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

Configurable engine with per-instance extraction/output settings and automatic include/exclude flags.

- `StructOptions`
  - `mode`: `"light" | "standard" | "verbose"`.
  - `table_params`: forwarded to `set_table_detection_params` inside a temporary scope.
  - `include_cell_links`: `None` -> auto (`verbose` only), otherwise respect the boolean.
- `OutputOptions` (nested: `format`, `filters`, `destinations`)
  - Format: `fmt` (`json`/`yaml`/`yml`/`toon`), `pretty`, `indent`.
  - Filters: `include_rows`, `include_shapes`, `include_shape_size` (None -> auto: verbose only), `include_charts`, `include_chart_size` (None -> auto: verbose only), `include_tables`, `include_print_areas` (None -> auto: light=False, others=True), `include_auto_print_areas` (default False).
  - Destinations: `sheets_dir`, `print_areas_dir`, `auto_page_breaks_dir`, `stream`.
- Methods:
  - `extract(path, mode=None)` -> WorkbookData (hyperlinks toggled via `StructOptions.include_cell_links`)
  - `serialize(workbook, fmt=None, pretty=None, indent=None)` -> str (applies filters/size flags/print-area include rules)
  - `export(workbook, output_path=None, fmt=None, pretty=None, indent=None, sheets_dir=None, print_areas_dir=None, auto_page_breaks_dir=None, stream=None)`
  - `process(file_path, output_path=None, out_fmt=None, image=False, pdf=False, dpi=72, mode=None, pretty=None, indent=None, sheets_dir=None, print_areas_dir=None, auto_page_breaks_dir=None, stream=None)` (PDF/PNG require Excel + pypdfium2)

Example:

```python
from pathlib import Path
from exstruct import (
    DestinationOptions,
    ExStructEngine,
    FilterOptions,
    FormatOptions,
    OutputOptions,
    StructOptions,
)

engine = ExStructEngine(
    options=StructOptions(mode="standard", include_cell_links=True),  # enable hyperlinks in standard mode
    output=OutputOptions(
        format=FormatOptions(pretty=True),
        filters=FilterOptions(
            include_shapes=False,
            include_print_areas=None,  # auto: light=False, others=True
            include_auto_print_areas=True,
        ),
        destinations=DestinationOptions(auto_page_breaks_dir=Path("auto_areas")),
    ),
)
wb = engine.extract("input.xlsx")
engine.export(wb, "out.json")              # writes filtered JSON and auto page-break files (COM only)
engine.process("input.xlsx", pdf=False)    # end-to-end extract + export
```

## Models

| Model           | Key fields                                                                                                                                                                                           |
| --------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `WorkbookData`  | `book_name: str`, `sheets: dict[str, SheetData]`                                                                                                                                                     |
| `SheetData`     | `rows: list[CellRow]`, `shapes: list[Shape]`, `charts: list[Chart]`, `table_candidates: list[str]`, `print_areas: list[PrintArea]`, `auto_print_areas: list[PrintArea]`                               |
| `CellRow`       | `r: int`, `c: dict[str, int                                                                        or float                                                or str]`, `links: dict[str, str] \| None` |
| `Shape`         | `text: str`, `l/t/w/h: int \| None`, `type`, `rotation`, arrow styles, `direction`                                                                                                                   |
| `Chart`         | `name`, `chart_type`, `title`, `series`, `y_axis_range`, `w/h: int \| None`, `l/t`, `error: str \| None`                                                                                             |
| `PrintArea`     | `r1/c1/r2/c2: int`                                                                                                                                                                                   |
| `PrintAreaView` | `book_name`, `sheet_name`, `area: PrintArea`, `rows`, `shapes`, `charts`, `table_candidates`                                                                                                         |
| `ChartSeries`   | `name`, `name_range`, `x_range`, `y_range`                                                                                                                                                           |

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
- No print areas: `export_print_areas_as` writes nothing and returns `{}`; this is not an error.
- Auto page-break export: `export_auto_page_breaks` raises `ValueError` if no auto page-break areas are present (enable them via `DestinationOptions.auto_page_breaks_dir`).

## Tuning Examples

- Reduce false positives (layout frames):

```python
set_table_detection_params(table_score_threshold=0.4, coverage_min=0.25)
```

- Recover missed tiny tables:

```python
set_table_detection_params(density_min=0.03, min_nonempty_cells=2)
```
