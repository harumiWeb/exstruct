# ExStruct Architecture Overview

ExStruct uses a pipeline-centric extraction architecture where
openpyxl, Excel COM (xlwings), and LibreOffice backends
are assigned roles based on the selected mode.

## Overall Structure

```txt
exstruct/
  core/
    pipeline.py
    integrate.py
    modeling.py
    workbook.py
    backends/
      base.py
      openpyxl_backend.py
      com_backend.py
      libreoffice_backend.py
    cells.py
    shapes.py
    charts.py
    ranges.py
    logging_utils.py
  models/
    __init__.py
    maps.py
  io/
    serialize.py
  render/
  edit/
    __init__.py
    a1.py
    api.py
    chart_types.py
    engine/
      __init__.py
      openpyxl_engine.py
      xlwings_engine.py
    errors.py
    internal.py
    output_path.py
    runtime.py
    service.py
    models.py
    normalize.py
    op_schema.py
    specs.py
    types.py
  cli/
    edit.py
    main.py
```

## Pipeline Design

- `resolve_extraction_inputs` normalizes include_* and mode
- `PipelinePlan` holds only the static step configuration for pre-com / com
- Execution state is separated into `PipelineState` / `PipelineResult`
- `run_extraction_pipeline` centrally manages COM availability checks and fallback

## Module Responsibilities

### core/

The central extraction layer (aggregates external dependencies)

- `pipeline.py` → extraction flow, COM determination, fallback, raw data generation
- `backends/*` → abstraction over openpyxl/COM/LibreOffice
- `cells.py` → cell extraction, table detection, colors_map
- `shapes.py` → shape extraction, direction estimation
- `charts.py` → chart analysis
- `ranges.py` → shared range analysis utilities
- `workbook.py` → openpyxl/xlwings context managers
- `modeling.py` → builds WorkbookData/SheetData from RawData
- `integrate.py` → thin entry point dedicated to pipeline calls

### models/

Public data structures via Pydantic
(external API returns BaseModel)

### io/

Output formats (JSON / YAML / TOON) and file writing

### render/

PDF/PNG output (for RAG use cases)

### cli/

CLI entry point

- `main.py` keeps the legacy extraction CLI and dispatches to editing
  subcommands only when the first token matches `patch` / `make` / `ops` /
  `validate`
- `edit.py` contains the Phase 2 editing parser, JSON serialization helpers,
  and wrappers around `exstruct.edit`
- `exstruct.__init__`, `exstruct.edit.__init__`, and lightweight CLI startup
  paths must remain side-effect-free: `--help` and `ops` routing should defer
  heavy extraction/edit implementation imports until command execution needs
  them

### edit/

First-class public workbook editing API

- `api.py` → public patch/make entry points for Python callers
- `service.py` → canonical patch/make orchestration used by both Python API and MCP
- `models.py` → canonical edit request/result models
- `runtime.py` → canonical backend selection, fallback, and policy-free path/runtime helpers
- `internal.py` → edit-owned low-level patch implementation and structured patch errors
- `output_path.py` → edit-owned output/conflict helpers reusable by host shims
- `engine/*` → canonical backend execution boundaries
- `a1.py` → A1 helpers owned by the edit core
- `normalize.py` / `specs.py` / `op_schema.py` → public patch-op normalization and schema metadata
- `edit/` does not import `mcp/`; MCP is allowed to depend on `edit`, not vice versa

### mcp/patch (Patch Implementation)

MCP editing remains the integration layer around the public edit API.

- `patch_runner.py` → compatibility facade for maintaining existing import paths and syncing host overrides
- `patch/internal.py` → compatibility facade re-exporting edit-owned internal implementation
- `patch/service.py` / `patch/runtime.py` / `patch/engine/*` → compatibility shims around `exstruct.edit`
- Legacy monkeypatch compatibility in these shims should prefer live module lookup over copied function aliases, and override precedence should be verified at the highest public compatibility entrypoint.
- `patch/ops/openpyxl_ops.py` / `patch/ops/xlwings_ops.py` → legacy op entry points kept for compatibility
- `patch/normalize.py` / `patch/specs.py` → op normalization and spec metadata
- `shared/a1.py` / `shared/output_path.py` → shared utilities for A1 notation and output paths

### mcp/

Host-layer responsibilities for MCP and agent tooling

- `io.py` → `PathPolicy` sandbox boundary
- `tools.py` / `server.py` → tool transport, artifact mirroring, runtime defaults, and thread offloading
- The MCP layer keeps safety policy and host behavior out of the public Python editing API

---

## Guide for AI Agents

- Reflect model changes in the core RawData conversion as well
- Contain external dependencies (openpyxl/xlwings) within the core boundary
- `PipelinePlan` is immutable; separate execution state into `PipelineState`
- For public contracts, refer to `docs/`; for internal model specifications, refer to `dev-docs/specs/data-model.md`
