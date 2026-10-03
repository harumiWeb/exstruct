# ADR-0013: Default OOXML Backend for Light Extraction

## 状態

`proposed`

## 背景

Issue #149 added a reusable OOXML session for workbook core data and rich
drawings. Issue #150 selects that implementation for the normal `.xlsx` /
`.xlsm` `light` path. This changes backend priority and failure handling, so it
requires a policy decision beyond the capability work recorded in ADR-0011.

ADR-0011 records direct openpyxl cell reading and extraction-scoped workbook
ownership. Its workbook-reuse and compatibility guarantees remain useful, but
the normal `light` OOXML path needs to avoid importing openpyxl, pandas, SciPy,
and xlwings while preserving the existing output model and legacy overrides.
Unsupported OOXML constructs also need a defined recovery rule: partial direct
results cannot safely be combined with results from a second reader.

## 決定

- For `.xlsx` and `.xlsm` in `light` mode, select the direct OOXML pipeline by
  default. One `OoxmlExtractionSession` owns one ZIP archive shared by core
  workbook extraction and rich shapes/charts through final model construction.
- The normal supported path must not import openpyxl, pandas, SciPy, or xlwings.
  Keep scalar normalization and table-candidate detection on the direct path
  independent of those optional/runtime readers. The direct light table
  heuristic explicitly passes `cluster_backend="python"` per call; it does not
  change `EXSTRUCT_BORDER_CLUSTER_BACKEND`. Other modes and compatibility
  extraction retain ADR-0012's configured backend selection.
- If any uncaught failure or unsupported construct occurs in any direct OOXML
  stage, discard all partial OOXML results, close the session in `finally`, and
  restart the complete extraction through `_run_openpyxl_pipeline`.
- Select that complete compatibility pipeline when `colors_map` is explicitly
  requested or a legacy pipeline/helper/workbook override is active. Known
  unsupported options and overrides should be detected before opening the ZIP
  where possible.
- Log the compatibility restart as a warning with
  `FallbackReason.OOXML_COMPATIBILITY` (`ooxml_compatibility`) and carry the
  reason in the extraction state when compatibility completes normally. If the
  compatibility pipeline takes its own fallback, preserve that later reason
  in state and retain the initial OOXML warning in the log.
- Preserve existing per-sheet best-effort drawing behavior: a drawing failure
  isolated to one sheet may omit that sheet's affected rich artifacts while
  healthy sheets and core extraction continue.
- Preserve the current output semantics and models. Other extraction modes
  and `.xls` selection/behavior remain unchanged.
- ADR-0011 continues to govern openpyxl resource ownership and direct reading
  for the compatibility path and other applicable extraction paths. This ADR
  refines only normal `.xlsx` / `.xlsm` `light` backend selection, so it does
  not supersede ADR-0011 as a whole.

## 影響

- Normal supported `light` extraction avoids loading four heavyweight reader
  or acceleration modules and reuses one archive for core and rich extraction.
- Openpyxl remains necessary for compatibility fallback and existing paths.
  Unsupported workbooks can pay the cost of a complete second extraction after
  the direct attempt, although partial results are never mixed.
- Direct OOXML readers and table heuristics must preserve the established
  normalization, table-candidate, rich-artifact, and serialized-output
  contracts. Parity regressions and an explicit import-boundary check are
  required.
- The recorded Issue #150 Windows run produced matching serialized-output
  hashes for all seven inputs. The direct pipeline had lower cold in-process
  time, warm median and peak working set on each input; the report and raw
  baseline contain the measurements and their limits. These are local
  observations, not universal performance guarantees.

## 根拠

- Tests: `tests/core/test_light_ooxml_pipeline.py` covers previous-pipeline
  parity, tracked sample parity, one-ZIP ownership and closure, whole-result
  restart after a stage failure, colors opt-in compatibility, and unchanged
  selection for other modes. `tests/core/test_backend_import_isolation.py`
  covers the normal-path import boundary.
- Code: `src/exstruct/core/pipeline.py`,
  `src/exstruct/core/ooxml_session.py`,
  `src/exstruct/core/ooxml_scalars.py`,
  `src/exstruct/core/backends/ooxml_backend.py`, and
  `src/exstruct/errors.py`.
- Related specs: `dev-docs/specs/excel-extraction.md`,
  `dev-docs/specs/ooxml-core-extraction.md`,
  `dev-docs/specs/extraction-performance.md`, `docs/api.md`, `docs/cli.md`,
  `README.md`, and `README.ja.md`.
- Benchmark: `benchmark/issue150_light.py`; final results are summarized in
  `benchmark/issue150-results.md` and recorded at
  `benchmark/baselines/issue150-2026-10-03-windows.json`.
- Related decisions: ADR-0010 defines the light rich-output boundary;
  ADR-0011 defines extraction-scoped openpyxl ownership and compatibility;
  ADR-0012 defines the optional SciPy selection and the direct-light exception.

## Supersedes

- None

## Superseded by

- None
