# CLI User Guide

This page explains how to run ExStruct from the command line, what each flag does, and common workflows. The CLI wraps `process_excel` under the hood.

## Basic usage

```bash
exstruct INPUT.xlsx > out.json              # compact JSON to stdout
exstruct INPUT.xlsx -o out.json --pretty    # pretty JSON to a file
exstruct INPUT.xlsx --format yaml           # YAML output (needs pyyaml)
exstruct INPUT.xlsx --format toon           # TOON output (needs python-toon)
```

- `INPUT.xlsx` supports `.xlsx/.xlsm/.xls`.
- Exit code `0` on success, `1` on failure.

## Options

| Flag | Description |
| ---- | ----------- |
| `-o, --output PATH` | Output path. Omit to write to stdout. |
| `-f, --format {json,yaml,yml,toon}` | Serialization format (default: `json`). |
| `-m, --mode {light,standard,verbose}` | Extraction detail level.<br>- light: cells + table candidates + print areas (no COM shapes/charts).<br>- standard: shapes with text/arrows + charts + print areas.<br>- verbose: all shapes/charts with size + hyperlinks. |
| `--pretty` | Pretty-print JSON (indent=2). |
| `--image` | Render per-sheet PNGs (requires Excel + COM + `pypdfium2`). |
| `--pdf` | Render PDF (requires Excel + COM + `pypdfium2`). |
| `--dpi INT` | DPI for rendered images (default: 144). |
| `--sheets-dir DIR` | Write one file per sheet (format follows `--format`). |
| `--print-areas-dir DIR` | Write one file per print area (format follows `--format`). |

## Common workflows

Split per sheet and keep a combined JSON:

```bash
exstruct sample.xlsx -o out.json --sheets-dir sheets/ --pretty
```

Export print areas (writes nothing if none exist):

```bash
exstruct sample.xlsx --print-areas-dir areas/
```

Verbose mode with hyperlinks, plus per-sheet YAML:

```bash
exstruct sample.xlsx --mode verbose --format yaml --sheets-dir sheets_yaml/  # needs pyyaml
```

Render PDF/PNG (Windows + Excel + `pypdfium2` required):

```bash
exstruct sample.xlsx --pdf --image --dpi 144 -o out.json
```

## Notes

- Optional dependencies are lazy-imported. Missing packages raise a `MissingDependencyError` with install hints.
- On non-COM environments, shapes/charts are empty and print areas come from openpyxl; commands still succeed.
- `--sheets-dir` and `--print-areas-dir` accept existing or new directories (created if missing).
