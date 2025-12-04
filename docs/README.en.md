# ExStruct — Excel Structured Extraction Engine

ExStruct reads Excel workbooks and outputs structured data (tables, shapes, charts) as JSON by default, with optional YAML/TOON formats. It targets both COM/Excel environments (rich extraction) and non-COM environments (cells + table candidates), with tunable detection heuristics and multiple output modes to fit LLM/RAG pipelines.

## Features

- **Excel → Structured JSON**: cells, shapes, charts, and table candidates per sheet.
- **Output modes**: `light` (cells + table candidates only), `standard` (texted shapes + arrows, charts), `verbose` (all shapes with width/height).
- **Formats**: JSON (compact by default, `--pretty` available), YAML, TOON (optional dependencies).
- **Table detection tuning**: adjust heuristics at runtime via API.
- **CLI rendering** (Excel required): optional PDF and per-sheet PNGs.
- **Graceful fallback**: if Excel COM is unavailable, extraction falls back to cells + table candidates without crashing.

## Installation

```bash
pip install exstruct
```

Optional extras:

- YAML: `pip install pyyaml`
- TOON: `pip install python-toon`
- Rendering (PDF/PNG): Excel + `pip install pypdfium2`

## Quick Start (CLI)

```bash
exstruct input.xlsx                # compact JSON (default)
exstruct input.xlsx --pretty       # pretty-printed JSON
exstruct input.xlsx --format yaml  # YAML (needs pyyaml)
exstruct input.xlsx --format toon  # TOON (needs python-toon)
exstruct input.xlsx --mode light   # cells + table candidates only
exstruct input.xlsx --pdf --image  # PDF and PNGs (Excel required)
```

## Quick Start (Python)

```python
from pathlib import Path
from exstruct import extract, export, set_table_detection_params

# Tune table detection (optional)
set_table_detection_params(table_score_threshold=0.3, density_min=0.04)

# Extract with modes: "light", "standard", "verbose"
wb = extract("input.xlsx", mode="standard")
export(wb, Path("out.json"), pretty=False)  # compact JSON
```

## Table Detection Tuning

```python
from exstruct import set_table_detection_params

set_table_detection_params(
    table_score_threshold=0.35,  # increase to be stricter
    density_min=0.05,
    coverage_min=0.2,
    min_nonempty_cells=3,
)
```

Use higher thresholds to reduce false positives; lower them if true tables are missed.

## Output Modes

- **light**: cells + table candidates (no COM needed).
- **standard**: texted shapes + arrows, charts (COM if available), table candidates.
- **verbose**: all shapes (with width/height), charts, table candidates.

## Error Handling / Fallbacks

- Excel COM unavailable → falls back to cells + table candidates; shapes/charts empty.
- Shape extraction failure → logs warning, still returns cells + table candidates.
- CLI prints errors to stdout/stderr and returns non-zero on failures.

## Optional Rendering

Requires Excel and `pypdfium2`.

```bash
exstruct input.xlsx --pdf --image --dpi 144
```

Creates `<output>.pdf` and `<output>_images/` PNGs per sheet.

## Notes

- Default JSON is compact to reduce tokens; use `--pretty` or `pretty=True` when readability matters.
- Field `table_candidates` replaces `tables`; adjust downstream consumers accordingly.
