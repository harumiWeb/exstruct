# ExStruct - Excel structure extraction engine overview

**ExStruct** is a Python library that extracts semantic structure from Excel workbooks.
It combines openpyxl, Excel COM (xlwings), and a LibreOffice backend to generate structured data that LLMs can work with easily.

## Features

- Unifies the extraction flow with a pipeline-oriented design
- Switches extraction granularity by mode (`light` / `libreoffice` / `standard` / `verbose`)
- Abstracts openpyxl / COM / LibreOffice backends
- Supports JSON / YAML / TOON output when dependencies are installed
- Supports `print_areas` and `auto_page_breaks` export
- Makes fallback reasons visible through unified logging

## Extraction targets

- Cells (values / links / coordinates)
- Tables (candidate ranges)
- Shapes / Arrows / SmartArt (position / text / arrows / layout)
- Charts (series / axes / type / title)
- Print Areas / Auto Page Breaks
- Colors Map (including conditional formatting)

## Usage examples at a glance

- Use `extract(path, mode="standard")` to obtain `WorkbookData`
- Use `process_excel` for file output or directory output
- Use the CLI as `exstruct file.xlsx --format json`

## Directory layout at a glance

```txt
docs/                 public documentation
dev-docs/             internal documentation
src/exstruct/
  core/               extraction pipeline and backends
  models/             Pydantic models
  io/                 JSON/YAML/TOON output
  render/             PDF/PNG output
  cli/                CLI
tests/                tests
```

AI agents should read `docs/` as the public contract and use `dev-docs/specs/` and `dev-docs/adr/` to fill in internal behavior.
