# MCP Server

This guide explains how to run ExStruct as an MCP (Model Context Protocol) server
so AI agents can call it safely as a tool.

## What it provides

- Convert Excel into structured JSON (file output)
- Edit Excel by applying patch operations (cell/sheet updates)
- Read large JSON outputs in chunks
- Pre-validate input files

## Installation

```bash
pip install exstruct[mcp]
```

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
- `exstruct_patch`
- `exstruct_read_json_chunk`
- `exstruct_validate_input`

### `exstruct_extract` defaults

- `options.alpha_col` defaults to `true` in MCP (column keys become `A`, `B`, ...).
- Set `options.alpha_col=false` if you need legacy 0-based numeric string keys.

## Basic flow

1. Call `exstruct_extract` to generate the output JSON file
2. Use `exstruct_read_json_chunk` to read only the parts you need

## Edit flow (patch)

1. Inspect workbook structure with `exstruct_extract` (and `exstruct_read_json_chunk` if needed)
2. Build patch operations (`ops`) for target cells/sheets
3. Call `exstruct_patch` to apply edits
4. Re-run `exstruct_extract` to verify results if needed

### `exstruct_patch` highlights

- Atomic apply: all operations succeed, or no changes are saved
- Supports:
  - `set_value`
  - `set_formula`
  - `add_sheet`
  - `set_range_values`
  - `fill_formula`
  - `set_value_if`
  - `set_formula_if`
- Useful flags:
  - `dry_run`: compute diff only (no file write)
  - `return_inverse_ops`: return undo operations
  - `preflight_formula_check`: detect formula issues before save
  - `auto_formula`: treat `=...` in `set_value` as formula
- Conflict handling follows server `--on-conflict` unless overridden per tool call

## AI agent configuration examples

### Codex

`~/.codex/config.toml`

```toml
[mcp_servers.exstruct]
command = "exstruct-mcp"
args = ["--root", "C:\\data", "--log-file", "C:\\logs\\exstruct-mcp.log", "--on-conflict", "rename"]
```

### GitHub Copilot / Claude Desktop / Gemini CLI

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
