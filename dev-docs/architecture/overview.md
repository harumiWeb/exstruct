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
  cli/
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

### mcp/patch (Patch Implementation)

Patch functionality is responsibility-separated under `src/exstruct/mcp/patch/`.

- `patch_runner.py` → compatibility facade for maintaining existing import paths
- `patch/internal.py` → internal compatibility layer for patch implementation (non-public)
- `patch/service.py` → orchestration of `run_patch` / `run_make`
- `patch/runtime.py` → runtime utilities for path/backend selection
- `patch/engine/openpyxl_engine.py` → openpyxl execution boundary
- `patch/engine/xlwings_engine.py` → xlwings (COM) execution boundary
- `patch/ops/openpyxl_ops.py` → op application entry point for openpyxl
- `patch/ops/xlwings_ops.py` → op application entry point for xlwings
- `patch/normalize.py` / `patch/specs.py` → op normalization and spec metadata
- `shared/a1.py` / `shared/output_path.py` → shared utilities for A1 notation and output paths

---

## Guide for AI Agents

- Reflect model changes in the core RawData conversion as well
- Contain external dependencies (openpyxl/xlwings) within the core boundary
- `PipelinePlan` is immutable; separate execution state into `PipelineState`
- For public contracts, refer to `docs/`; for internal model specifications, refer to `dev-docs/specs/data-model.md`
