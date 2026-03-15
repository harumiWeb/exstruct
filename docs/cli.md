# CLI User Guide

This page explains how to run ExStruct from the command line, what each flag
does, and common workflows.

- Extraction keeps the legacy `exstruct INPUT.xlsx ...` form and wraps
  `process_excel`.
- Editing uses subcommands such as `exstruct patch` and wraps `exstruct.edit`.

## Basic usage

```bash
exstruct INPUT.xlsx > out.json              # compact JSON to stdout
exstruct INPUT.xlsx -o out.json --pretty    # pretty JSON to a file
exstruct INPUT.xlsx --format yaml           # YAML output (needs pyyaml)
exstruct INPUT.xlsx --format toon           # TOON output (needs python-toon)
```

- `INPUT.xlsx` supports `.xlsx/.xlsm/.xls`.
- Exit code `0` on success, `1` on failure.

## Editing commands

Phase 2 adds JSON-first editing commands while keeping the extraction entrypoint
unchanged.

```bash
exstruct patch --input book.xlsx --ops ops.json --backend openpyxl
exstruct patch --input book.xlsx --ops - --dry-run --pretty < ops.json
exstruct make --output new.xlsx --ops ops.json --backend openpyxl
exstruct ops list
exstruct ops describe create_chart --pretty
exstruct validate --input book.xlsx --pretty
```

- `patch` serializes `PatchResult` to stdout and exits `1` only when
  `PatchResult.error` is present.
- `make` serializes `PatchResult` for new workbook creation.
- `ops list` returns compact `{op, description}` summaries.
- `ops describe` returns the detailed schema for one patch op.
- `validate` returns input readability checks (`is_readable`, `warnings`,
  `errors`).

## Editing options

### `patch`

| Flag | Description |
| ---- | ----------- |
| `--input PATH` | Existing workbook to edit. |
| `--ops FILE\|-` | JSON array of patch ops from a file or stdin. |
| `--output PATH` | Optional output workbook path. If omitted, the existing default patch output naming applies. |
| `--sheet TEXT` | Top-level sheet fallback for patch ops. |
| `--on-conflict {overwrite,skip,rename}` | Output conflict policy. |
| `--backend {auto,com,openpyxl}` | Backend selection. |
| `--auto-formula` | Treat `=...` values in `set_value` ops as formulas. |
| `--dry-run` | Simulate changes without saving. |
| `--return-inverse-ops` | Return inverse ops when supported. |
| `--preflight-formula-check` | Run formula-health validation before saving when supported. |
| `--pretty` | Pretty-print JSON output. |

### `make`

`make` accepts the same flags as `patch`, except that `--output PATH` is
required and `--input` is not used. `--ops` is optional; omitting it creates an
empty workbook.

### `ops` and `validate`

- `exstruct ops list [--pretty]`
- `exstruct ops describe OP [--pretty]`
- `exstruct validate --input PATH [--pretty]`

## Options

| Flag | Description |
| ---- | ----------- |
| `-o, --output PATH` | Output path. Omit to write to stdout. |
| `-f, --format {json,yaml,yml,toon}` | Serialization format (default: `json`). |
| `-m, --mode {light,libreoffice,standard,verbose}` | Extraction detail level.<br>- light: cells + table candidates + print areas only.<br>- libreoffice: best-effort non-COM mode for `.xlsx/.xlsm`; adds merged cells, shapes, connectors, and charts when LibreOffice runtime is available.<br>- standard: shapes with text/arrows + charts + print areas via Excel COM.<br>- verbose: all shapes/charts with size + hyperlinks/maps via Excel COM. |
| `--alpha-col` | Output column keys as Excel-style names (`A`, `B`, ..., `AA`) instead of 0-based numeric keys (`"0"`, `"1"`, ...). Default: disabled (legacy numeric keys). |
| `--pretty` | Pretty-print JSON (indent=2). |
| `--image` | Render per-sheet PNGs (requires Excel + COM + `pypdfium2`; not supported in `--mode libreoffice`). |
| `--pdf` | Render PDF (requires Excel + COM + `pypdfium2`; not supported in `--mode libreoffice`). |
| `--dpi INT` | DPI for rendered images (default: 144). |
| `--include-backend-metadata` | Include shape/chart backend metadata (`provenance`, `approximation_level`, `confidence`) in structured output. |
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
- Editing commands are JSON-first and do not add interactive confirmation,
  backup creation, or path-restriction flags in Phase 2.
- On non-COM environments, prefer `--mode libreoffice` for best-effort rich extraction on `.xlsx/.xlsm`, or `--mode light` for minimal extraction.
- `--mode libreoffice` is best-effort, not a strict subset of COM output. It does not render PDFs/PNGs and does not compute auto page-break areas in v1.
- `--mode libreoffice` combined with `--pdf`, `--image`, or `--auto-page-breaks-dir` fails early with a configuration error instead of silently ignoring the option.
- `--sheets-dir` and `--print-areas-dir` accept existing or new directories (created if missing).
- `--alpha-col` switches row column keys from legacy numeric strings (`"0"`, `"1"`, ...) to Excel-style keys (`"A"`, `"B"`, ...). CLI default is disabled for backward compatibility.
