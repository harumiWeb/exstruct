# CLI User Guide

This page explains how to run ExStruct from the command line, what each flag
does, and the recommended workflows for extraction and workbook editing.

- Extraction keeps the legacy `exstruct INPUT.xlsx ...` form and wraps
  `process_excel`.
- Editing uses subcommands such as `exstruct patch`, wraps `exstruct.edit`, and
  serves as the canonical operational / agent interface for workbook editing.

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
unchanged. Prefer these commands for local shell automation or AI-agent edit
workflows. If you are writing direct Python workbook-editing code,
`openpyxl` / `xlwings` are usually simpler; use `exstruct.edit` only when you
need ExStruct's patch contract inside Python.

```bash
exstruct patch --input book.xlsx --ops ops.json --backend openpyxl
exstruct patch --input book.xlsx --ops - --dry-run --pretty < ops.json
exstruct make --output new.xlsx --ops ops.json --backend openpyxl
exstruct ops list
exstruct ops describe create_chart --pretty
exstruct validate --input book.xlsx --pretty
```

- `patch` serializes `PatchResult` to stdout once request parsing and execution
  begin. Invalid JSON, request validation failures, and local runtime errors
  are printed to stderr and exit `1` before any JSON payload is produced.
- `make` follows the same stdout/stderr contract for new workbook creation.
- `ops list` returns compact `{op, description}` summaries.
- `ops describe` returns the detailed schema for one patch op.
- `validate` returns input readability checks (`is_readable`, `warnings`,
  `errors`).

## Recommended editing workflow

1. Build or load your patch op JSON.
2. Run `exstruct patch --dry-run` and inspect `PatchResult`, warnings, and
   `patch_diff`.
3. If you want the dry run and the real apply to use the same engine, pin
   `--backend openpyxl`.
4. If you keep `--backend auto`, inspect `PatchResult.engine`; on
   Windows/Excel hosts the non-dry-run apply may switch from openpyxl to COM.
5. Apply the chosen backend without `--dry-run` only after the result is
   acceptable.
6. Re-run extraction or another validation step if the workbook is part of a
   larger pipeline.

If you need host-managed path restrictions, transport mapping, or artifact
mirroring, switch to MCP instead of extending the local CLI path.

## Editing backend guidance

- `--backend openpyxl` is the default choice for ordinary cell/style edits and
  all `--dry-run` / `--return-inverse-ops` / `--preflight-formula-check`
  workflows.
- `--backend com` is required for COM-only behavior such as `create_chart` and
  `.xls` editing.
- `--backend auto` keeps the existing backend-selection policy and reports the
  actual backend in `PatchResult.engine`. `--dry-run`,
  `--return-inverse-ops`, and `--preflight-formula-check` still force the
  openpyxl path.

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
| `-m, --mode {light,libreoffice,standard,verbose}` | Extraction detail level.<br>- light: pure-Python baseline; returns cells + table candidates + print areas, and on `.xlsx/.xlsm` adds best-effort OOXML shapes, connectors, and charts.<br>- libreoffice: optional non-COM enrichment mode for `.xlsx/.xlsm`; starts from the same OOXML baseline as `light` and uses LibreOffice runtime to refine merged cells, shapes, connectors, and charts when available.<br>- standard: shapes with text/arrows + charts + print areas via Excel COM.<br>- verbose: all shapes/charts with size + hyperlinks/maps via Excel COM. |
| `--alpha-col` | Output column keys as Excel-style names (`A`, `B`, ..., `AA`) instead of 0-based numeric keys (`"0"`, `"1"`, ...). Default: disabled (legacy numeric keys). |
| `--pretty` | Pretty-print JSON (indent=2). |
| `--image` | Render per-sheet PNGs (requires Excel + COM + `pypdfium2`; not supported in `--mode libreoffice`). |
| `--pdf` | Render PDF (requires Excel + COM + `pypdfium2`; not supported in `--mode libreoffice`). |
| `--dpi INT` | DPI for rendered images (default: 144). |
| `--include-backend-metadata` | Include shape/chart backend metadata (`provenance`, `approximation_level`, `confidence`) in structured output. |
| `--sheets-dir DIR` | Write one file per sheet (format follows `--format`). |
| `--print-areas-dir DIR` | Write one file per print area (format follows `--format`). |
| `--auto-page-breaks-dir DIR` | Write one file per auto page-break area. The flag is always shown in help, but execution requires `--mode standard` or `--mode verbose` with Excel COM. |

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
- Use the CLI for local operational flows; use MCP when you need host-owned
  safety policy. For direct Python workbook editing, `openpyxl` / `xlwings`
  are usually the better fit.
- On non-COM environments, prefer `--mode light` as the pure-Python baseline for `.xlsx/.xlsm`; use `--mode libreoffice` only when you want optional LibreOffice-based enrichment on top of that baseline.
- `--mode libreoffice` is best-effort, not a strict subset of COM output. It does not render PDFs/PNGs and does not compute auto page-break areas in v1.
- `--auto-page-breaks-dir` is always shown in help output and is validated at execution time.
- `--mode libreoffice` combined with `--pdf`, `--image`, or `--auto-page-breaks-dir` fails early with a configuration error instead of silently ignoring the option.
- `--mode light` also rejects `--auto-page-breaks-dir`; use `--mode standard` or `--mode verbose` with Excel COM for auto page-break export.
- `--sheets-dir` and `--print-areas-dir` accept existing or new directories (created if missing).
- `--alpha-col` switches row column keys from legacy numeric strings (`"0"`, `"1"`, ...) to Excel-style keys (`"A"`, `"B"`, ...). CLI default is disabled for backward compatibility.
