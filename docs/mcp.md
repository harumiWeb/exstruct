# MCP Server

This guide explains how to run ExStruct as an MCP (Model Context Protocol) server
so AI agents can call it safely as a tool.

## What it provides

- Convert Excel into structured JSON (file output)
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
- `exstruct_read_json_chunk`
- `exstruct_validate_input`

## Basic flow

1. Call `exstruct_extract` to generate the output JSON file
2. Use `exstruct_read_json_chunk` to read only the parts you need

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
