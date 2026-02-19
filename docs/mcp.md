# MCP Server

This guide explains how to run ExStruct as an MCP (Model Context Protocol) server
so AI agents can call it safely as a tool.

## What it provides

- Convert Excel into structured JSON (file output)
- Create a new workbook and apply initial ops in one call
- Edit Excel by applying patch operations (cell/sheet updates)
- Read large JSON outputs in chunks
- Read A1 ranges / specific cells / formulas directly from extracted JSON
- Pre-validate input files

## Installation

### Option 1: Using uvx (recommended)

No installation required! Run directly with uvx:

```bash
uvx --from 'exstruct[mcp]' exstruct-mcp --root C:\data
```

Benefits:
- No `pip install` needed
- Automatic dependency management  
- Environment isolation
- Easy version pinning: `uvx --from 'exstruct[mcp]==0.4.4' exstruct-mcp`

### Option 2: Traditional pip install

```bash
pip install exstruct[mcp]
```

### Option 3: Development version from Git

```bash
uvx --from 'exstruct[mcp] @ git+https://github.com/harumiWeb/exstruct.git@main' exstruct-mcp --root .
```

**Note:** When using Git URLs, the `[mcp]` extra must be explicitly included in the dependency specification.

## Start (stdio)

```bash
exstruct-mcp --root C:\\data --log-file C:\\logs\\exstruct-mcp.log --on-conflict rename
```

### Key options

- `--root`: Allowed root directory (required)
- `--deny-glob`: Deny glob patterns (repeatable)
- `--log-level`: `DEBUG` / `INFO` / `WARNING` / `ERROR`
- `--log-file`: Log file path (stderr is still used by default)
- `--on-conflict`: Output conflict policy (`overwrite` / `skip` / `rename`)
- `--warmup`: Preload heavy imports to reduce first-call latency

## Tools

- `exstruct_extract`
- `exstruct_make`
- `exstruct_patch`
- `exstruct_read_json_chunk`
- `exstruct_read_range`
- `exstruct_read_cells`
- `exstruct_read_formulas`
- `exstruct_validate_input`

### `exstruct_extract` defaults and mode guide

- `options.alpha_col` defaults to `true` in MCP (column keys become `A`, `B`, ...).
- Set `options.alpha_col=false` if you need legacy 0-based numeric string keys.
- `mode` is an extraction detail level (not sheet scope):

| Mode | When to use | Main output characteristics |
|---|---|---|
| `light` | Fast, structure-first extraction | cells + table candidates + print areas |
| `standard` | Default for most agent flows | balanced detail and size |
| `verbose` | Need the richest metadata | adds links/maps and richer metadata |

## Quick start for agents (recommended)

1. Validate file readability with `exstruct_validate_input`
2. Run `exstruct_extract` with `mode="standard"`
3. Read the result with `exstruct_read_json_chunk` using `sheet` and `max_bytes`

Example sequence:

```json
{ "tool": "exstruct_validate_input", "xlsx_path": "C:\\data\\book.xlsx" }
```

```json
{ "tool": "exstruct_extract", "xlsx_path": "C:\\data\\book.xlsx", "mode": "standard", "format": "json" }
```

```json
{ "tool": "exstruct_read_json_chunk", "out_path": "C:\\data\\book.json", "sheet": "Sheet1", "max_bytes": 50000 }
```

## Basic flow

1. Call `exstruct_extract` to generate the output JSON file
2. Use `exstruct_read_json_chunk` to read only the parts you need

## Direct read tools (A1-oriented)

Use these tools when you already know the target addresses and want faster,
less verbose reads than chunk traversal.

- `exstruct_read_range`
  - Read a rectangular A1 range (example: `A1:G10`)
  - Optional: `include_formulas`, `include_empty`, `max_cells`
- `exstruct_read_cells`
  - Read specific cells in one call (example: `["J98", "J124"]`)
  - Optional: `include_formulas`
- `exstruct_read_formulas`
  - Read formulas only (optionally restricted by A1 range)
  - Optional: `include_values`

Examples:

```json
{
  "tool": "exstruct_read_range",
  "out_path": "C:\\data\\book.json",
  "sheet": "Data",
  "range": "A1:G10"
}
```

```json
{
  "tool": "exstruct_read_cells",
  "out_path": "C:\\data\\book.json",
  "sheet": "Data",
  "addresses": ["J98", "J124"]
}
```

```json
{
  "tool": "exstruct_read_formulas",
  "out_path": "C:\\data\\book.json",
  "sheet": "Data",
  "range": "J2:J201",
  "include_values": true
}
```

## Chunking guide

### Key parameters

- `sheet`: target sheet name. Strongly recommended when workbook has multiple sheets.
- `max_bytes`: chunk size budget in bytes. Start at `50_000`; increase (for example `120_000`) if chunks are too small.
- `filter.rows`: `[start, end]` (1-based, inclusive).
- `filter.cols`: `[start, end]` (1-based, inclusive). Works for both numeric keys (`"0"`, `"1"`) and alpha keys (`"A"`, `"B"`).
- `cursor`: pagination cursor (`next_cursor` from the previous response).

### Retry guide by error/warning

| Message | Meaning | Next action |
|---|---|---|
| `Output is too large...` | Whole JSON cannot fit in one response | Retry with `sheet`, or narrow with `filter.rows`/`filter.cols` |
| `Sheet is required when multiple sheets exist...` | Workbook has multiple sheets and target is ambiguous | Pick one value from `workbook_meta.sheet_names` and set `sheet` |
| `Base payload exceeds max_bytes...` | Even metadata-only payload is larger than `max_bytes` | Increase `max_bytes` |
| `max_bytes too small...` | Row payload is too large for the current size | Increase `max_bytes`, or narrow row/col filters |

### Cursor example

1. Call without `cursor`
2. If response has `next_cursor`, call again with that cursor
3. Repeat until `next_cursor` is `null`

## Edit flow (make/patch)

### New workbook flow (`exstruct_make`)

1. Build patch operations (`ops`) for initial sheets/cells
2. Call `exstruct_make` with `out_path`
3. Re-run `exstruct_extract` to verify results if needed

### `exstruct_make` highlights

- Creates a new workbook and applies `ops` in one call
- `out_path` is required
- `ops` is optional (empty list is allowed)
- Supported output extensions: `.xlsx`, `.xlsm`, `.xls`
- Initial sheet is normalized to `Sheet1`
- Reuses patch pipeline, so `patch_diff`/`error` shape is compatible with `exstruct_patch`
- Supports the same extended flags as `exstruct_patch`:
  - `dry_run`
  - `return_inverse_ops`
  - `preflight_formula_check`
  - `auto_formula`
  - `backend`
- `.xls` constraints:
  - requires Windows Excel COM
  - rejects `backend="openpyxl"`
  - rejects `dry_run`/`return_inverse_ops`/`preflight_formula_check`

Example:

```json
{
  "tool": "exstruct_make",
  "out_path": "C:\\data\\new_book.xlsx",
  "ops": [
    { "op": "add_sheet", "sheet": "Data" },
    { "op": "set_value", "sheet": "Data", "cell": "A1", "value": "hello" }
  ]
}
```

## Edit flow (patch)

1. Inspect workbook structure with `exstruct_extract` (and `exstruct_read_json_chunk` if needed)
2. Build patch operations (`ops`) for target cells/sheets
3. Call `exstruct_patch` to apply edits
4. Re-run `exstruct_extract` to verify results if needed

### `exstruct_patch` highlights

- Atomic apply: all operations succeed, or no changes are saved
- `ops` accepts an object list as the canonical form.
  For compatibility, JSON object strings in `ops` are also accepted and normalized.
- Supports:
  - `set_value`
  - `set_formula`
  - `add_sheet`
  - `set_range_values`
  - `fill_formula`
  - `set_value_if`
  - `set_formula_if`
  - `draw_grid_border`
  - `set_bold`
  - `set_font_size`
  - `set_fill_color`
  - `set_dimensions`
  - `merge_cells`
  - `unmerge_cells`
  - `set_alignment`
  - `restore_design_snapshot` (internal inverse op)
- Useful flags:
  - `dry_run`: compute diff only (no file write)
  - `return_inverse_ops`: return undo operations
  - `preflight_formula_check`: detect formula issues before save
  - `auto_formula`: treat `=...` in `set_value` as formula
- Backend selection:
  - `backend="auto"` (default): prefers COM when available; otherwise openpyxl.
    Also uses openpyxl when `dry_run`/`return_inverse_ops`/`preflight_formula_check` is enabled.
  - `backend="com"`: forces COM. Requires Excel COM and rejects
    `dry_run`/`return_inverse_ops`/`preflight_formula_check`.
  - `backend="openpyxl"`: forces openpyxl (`.xls` is not supported).
- Output includes `engine` (`"com"` or `"openpyxl"`) to show which backend was actually used.
- Conflict handling follows server `--on-conflict` unless overridden per tool call
- `restore_design_snapshot` remains openpyxl-only.

## AI agent configuration examples

### Using uvx (recommended)

#### Claude Desktop / GitHub Copilot

```json
{
  "mcpServers": {
    "exstruct": {
      "command": "uvx",
      "args": [
        "--from",
        "exstruct[mcp]",
        "exstruct-mcp",
        "--root",
        "C:\\data",
        "--log-file",
        "C:\\logs\\exstruct-mcp.log",
        "--on-conflict",
        "rename"
      ]
    }
  }
}
```

#### Codex

```toml
[mcp_servers.exstruct]
command = "uvx"
args = [
  "--from",
  "exstruct[mcp]",
  "exstruct-mcp",
  "--root",
  "C:\\data",
  "--log-file",
  "C:\\logs\\exstruct-mcp.log",
  "--on-conflict",
  "rename"
]
```

### Using pip install

#### Codex

`~/.codex/config.toml`

```toml
[mcp_servers.exstruct]
command = "exstruct-mcp"
args = ["--root", "C:\\data", "--log-file", "C:\\logs\\exstruct-mcp.log", "--on-conflict", "rename"]
```

#### GitHub Copilot / Claude Desktop / Gemini CLI

Register an MCP server with a command + args in your MCP settings:

```json
{
  "mcpServers": {
    "exstruct": {
      "command": "exstruct-mcp",
      "args": ["--root", "C:\\data"]
    }
  }
}
```

## Operational notes

- Logs go to stderr (and optionally `--log-file`) to avoid contaminating stdio responses.
- On Windows with Excel, standard/verbose can use COM for richer extraction.
  On non-Windows, COM is unavailable and openpyxl-based fallbacks are used.
- For large outputs, use `read_json_chunk` to avoid hitting client limits.
