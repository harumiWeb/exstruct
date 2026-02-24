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
- `--artifact-bridge-dir`: Directory used by `mirror_artifact=true` to copy output files
- `--warmup`: Preload heavy imports to reduce first-call latency

## Tools

- `exstruct_extract`
- `exstruct_make`
- `exstruct_patch`
- `exstruct_list_ops`
- `exstruct_describe_op`
- `exstruct_read_json_chunk`
- `exstruct_read_range`
- `exstruct_read_cells`
- `exstruct_read_formulas`
- `exstruct_validate_input`
- `exstruct_get_runtime_info`

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

If path behavior is unclear, inspect runtime info first:

```json
{ "tool": "exstruct_get_runtime_info" }
```

When a path is outside `--root`, the error message also recommends
`exstruct_get_runtime_info` with a relative path example.

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
- Initial sheet behavior:
  - default is `Sheet1`
  - when `sheet` is specified and the same name is not created by `add_sheet`,
    the initial sheet is created with that `sheet` name
  - when `add_sheet` creates the same name, initial sheet remains `Sheet1`
- Reuses patch pipeline, so `patch_diff`/`error` shape is compatible with `exstruct_patch`
- Supports the same extended flags as `exstruct_patch`:
  - `dry_run`
  - `return_inverse_ops`
  - `preflight_formula_check`
  - `auto_formula`
  - `backend`
  - `sheet` (top-level default sheet for non-`add_sheet` ops)
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

### Internal implementation note

The patch implementation is layered to keep compatibility while enabling refactoring:

- `exstruct.mcp.patch_runner`: compatibility facade (existing import path)
- `exstruct.mcp.patch.service`: patch/make orchestration
- `exstruct.mcp.patch.engine.*`: backend execution boundaries (openpyxl/com)
- `exstruct.mcp.patch.runtime`: runtime utilities (path/backend selection)
- `exstruct.mcp.patch.ops.*`: backend-specific op application entrypoints

This keeps MCP tool I/O stable while allowing internal module separation.

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
  - `set_font_color`
  - `set_fill_color`
  - `set_dimensions`
  - `auto_fit_columns`
  - `merge_cells`
  - `unmerge_cells`
  - `set_alignment`
  - `set_style`
  - `apply_table_style`
  - `restore_design_snapshot` (internal inverse op)
- Useful flags:
  - `dry_run`: compute diff only (no file write)
  - `return_inverse_ops`: return undo operations
  - `preflight_formula_check`: detect formula issues before save
  - `auto_formula`: treat `=...` in `set_value` as formula
  - `sheet`: top-level default sheet used when `op.sheet` is omitted (non-`add_sheet` only)
  - `mirror_artifact`: copy output workbook to `--artifact-bridge-dir` on success
- Large ops guidance:
  - `ops` over `200` still runs, but returns a warning that recommends splitting into batches.
- Backend selection:
  - `backend="auto"` (default): prefers COM when available; otherwise openpyxl.
    Also uses openpyxl when `dry_run`/`return_inverse_ops`/`preflight_formula_check` is enabled.
    Requests including `apply_table_style` are also routed to openpyxl.
  - `backend="com"`: forces COM. Requires Excel COM and rejects
    `dry_run`/`return_inverse_ops`/`preflight_formula_check`.
    If `apply_table_style` is included, returns a warning and falls back to openpyxl.
  - `backend="openpyxl"`: forces openpyxl (`.xls` is not supported).
- Output includes `engine` (`"com"` or `"openpyxl"`) to show which backend was actually used.
- Output includes `mirrored_out_path` when mirroring is requested and succeeds.
- Conflict handling follows server `--on-conflict` unless overridden per tool call
- `restore_design_snapshot` remains openpyxl-only.
- Sheet resolution order:
  - `op.sheet` is used when present
  - otherwise top-level `sheet` is used for non-`add_sheet` ops
  - `add_sheet` still requires explicit `op.sheet` (or alias `name`)

### `set_style` quick guide

- Purpose: apply multiple style fields in one op.
- Target: exactly one of `cell` or `range`.
- Need at least one style field: `bold`, `font_size`, `color`, `fill_color`,
  `horizontal_align`, `vertical_align`, `wrap_text`.

Example:

```json
{
  "tool": "exstruct_patch",
  "xlsx_path": "C:\\data\\book.xlsx",
  "ops": [
    {
      "op": "set_style",
      "sheet": "Sheet1",
      "range": "A1:D1",
      "bold": true,
      "color": "#FFFFFF",
      "fill_color": "#1F3864",
      "horizontal_align": "center",
      "vertical_align": "center",
      "wrap_text": true
    }
  ]
}
```

### `apply_table_style` quick guide

- Purpose: create a table and apply an Excel table style in one op.
- Required: `sheet`, `range`, `style`.
- Optional: `table_name`.
- Fails when range intersects an existing table, or table name duplicates.

Example:

```json
{
  "tool": "exstruct_patch",
  "xlsx_path": "C:\\data\\book.xlsx",
  "ops": [
    {
      "op": "apply_table_style",
      "sheet": "Sheet1",
      "range": "A1:D11",
      "style": "TableStyleMedium9",
      "table_name": "SalesTable"
    }
  ]
}
```

### `auto_fit_columns` quick guide

- Purpose: auto-fit column widths and optionally clamp with min/max bounds.
- Required: `sheet`.
- Optional: `columns`, `min_width`, `max_width`.
- `columns` supports Excel letters and numeric indexes (for example `["A", 2]`).
- If `columns` is omitted, used columns are targeted.

Example:

```json
{
  "tool": "exstruct_patch",
  "xlsx_path": "C:\\data\\book.xlsx",
  "ops": [
    {
      "op": "auto_fit_columns",
      "sheet": "Sheet1",
      "columns": ["A", 2],
      "min_width": 8,
      "max_width": 40
    }
  ]
}
```

### Color fields (`color` / `fill_color`)

- `set_font_color` uses `color` (font color only)
- `set_fill_color` uses `fill_color` (background fill only)
- Accepted formats for both fields:
  - `RRGGBB`
  - `AARRGGBB`
  - `#RRGGBB`
  - `#AARRGGBB`
- Values are normalized internally to uppercase with leading `#`.

Examples:

```json
{
  "tool": "exstruct_patch",
  "xlsx_path": "C:\\data\\book.xlsx",
  "ops": [
    { "op": "set_font_color", "sheet": "Sheet1", "cell": "A1", "color": "1f4e79" },
    { "op": "set_fill_color", "sheet": "Sheet1", "range": "A1:C1", "fill_color": "CC336699" }
  ]
}
```

### Alias and shorthand inputs

- `add_sheet`: `name` is accepted as an alias of `sheet`
- `set_dimensions`:
  - `row` -> `rows`
  - `col` -> `columns`
  - `height` -> `row_height`
  - `width` -> `column_width`
  - `columns` accepts both letters (`"A"`, `"AA"`) and positive indexes (`1`, `2`, ...)
- `draw_grid_border`: `range` shorthand is accepted and normalized to
  `base_cell` + `row_count` + `col_count`
- `set_alignment`:
  - `horizontal` -> `horizontal_align`
  - `vertical` -> `vertical_align`
- `set_fill_color`:
  - `color` -> `fill_color`
- Relative `out_path` for `exstruct_make` is resolved from MCP `--root`.

### Mirror artifact handoff

- `exstruct_patch` / `exstruct_make` input:
  - `mirror_artifact` (default: `false`)
- Output:
  - `mirrored_out_path` (`null` when not mirrored)
- Behavior:
  - Mirroring runs only on success.
  - If `--artifact-bridge-dir` is not set, process still succeeds and warning is returned.
  - If copy fails, process still succeeds and warning is returned.

## Op schema discovery tools

- `exstruct_list_ops`
  - Returns available op names and short descriptions.
- `exstruct_describe_op`
  - Input: `op`
  - Output: `required`, `optional`, `constraints`, `example`, `aliases`

Examples:

```json
{ "tool": "exstruct_list_ops" }
```

```json
{ "tool": "exstruct_describe_op", "op": "set_fill_color" }
```

## Mistake catalog (error -> fix)

- Wrong (conflicting alias and canonical field):
  - `{"op":"set_fill_color","sheet":"Sheet1","cell":"A1","color":"#D9E1F2","fill_color":"#FFFFFF"}`
- Correct:
  - `{"op":"set_fill_color","sheet":"Sheet1","cell":"A1","fill_color":"#D9E1F2"}`

- Wrong (conflicting alias and canonical field):
  - `{"op":"set_alignment","sheet":"Sheet1","cell":"A1","horizontal":"center","horizontal_align":"left"}`
- Correct:
  - `{"op":"set_alignment","sheet":"Sheet1","cell":"A1","horizontal_align":"center","vertical_align":"center"}`

- Note:
  - `color` (`set_fill_color`) and `horizontal`/`vertical` (`set_alignment`) are accepted aliases.
  - Canonical fields (`fill_color`, `horizontal_align`, `vertical_align`) are recommended.

### In-place overwrite recipe

To overwrite the original workbook path explicitly:

1. set `out_name` to the same filename as the input workbook
2. set `on_conflict` to `overwrite`
3. keep `out_dir` empty (or set it to the same directory)

Example:

```json
{
  "tool": "exstruct_patch",
  "xlsx_path": "C:\\data\\book.xlsx",
  "out_name": "book.xlsx",
  "on_conflict": "overwrite",
  "ops": [
    { "op": "set_value", "sheet": "Sheet1", "cell": "A1", "value": "updated" }
  ]
}
```

### Runtime info tool

- `exstruct_get_runtime_info` returns:
  - `root`
  - `cwd`
  - `platform`
  - `path_examples` (`relative` and `absolute`)

Example response (shape):

```json
{
  "root": "C:\\data",
  "cwd": "C:\\Users\\agent\\workspace",
  "platform": "win32",
  "path_examples": {
    "relative": "outputs/book.xlsx",
    "absolute": "C:\\data\\outputs\\book.xlsx"
  }
}
```

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
