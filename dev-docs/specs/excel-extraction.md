# Excel Extraction Specification — ExStruct

This document summarizes the current specification for Excel extraction processing.

## Overall Flow

1. `resolve_extraction_inputs` normalizes include_* and mode.
2. For `light` on `.xlsx` / `.xlsm`, the pipeline selects direct OOXML core
   extraction and the OOXML rich backend by default. They share one ZIP through
   final model construction.
3. If any uncaught failure or unsupported construct occurs in a direct OOXML
   stage, the pipeline closes the session, discards every partial OOXML result,
   and restarts the complete extraction through the openpyxl compatibility
   pipeline. It logs `FallbackReason.OOXML_COMPATIBILITY`
   (`ooxml_compatibility`) as a warning.
4. A drawing error isolated to one sheet retains the existing best-effort
   behavior: omit the affected drawing artifacts for that sheet and continue
   with healthy sheets.
5. An explicit `colors_map` request or an active legacy pipeline/helper/
   workbook override selects the complete openpyxl compatibility pipeline.
6. `libreoffice` retains its OOXML rich baseline and optional LibreOffice
   enrichment; `standard` / `verbose` retain their COM behavior. `.xls`
   selection and behavior are unchanged.
7. When COM succeeds, `colors_map` is overwritten with COM results.

## Workbook Resource Lifetime

- `standard` / `verbose` retain file-first extraction after the Issue #151
  experiment (ADR-0014). The bulk COM cell reader is a benchmark prototype;
  public extraction reads saved cell values through the file backend.
- A new Excel app is owned before opening the workbook, so opening failures
  still run app cleanup. Existing workbooks/apps are borrowed and never closed
  or terminated by extraction.

- A supported `light` OOXML extraction owns one `OoxmlExtractionSession` and
  one ZIP archive shared by core workbook parsing, drawing/chart extraction,
  table candidates and final model construction. The session closes in a
  `finally` path on success and failure before any compatibility restart.
- The openpyxl compatibility pipeline owns an `OpenpyxlExtractionSession`.
  Compatible
  OOXML stages share one regular `data_only=True`, `read_only=False` workbook,
  including hyperlink and per-sheet table detection, until final modeling ends.
- An openpyxl formula map opens a second regular `data_only=False` workbook
  lazily. Disabled formula extraction and COM-only `.xls` formula extraction
  must not open that variant.
- The normal supported `light` OOXML path opens one ZIP and no openpyxl
  workbook. If direct extraction fails, the ZIP is closed and the full
  compatibility pipeline starts from the beginning; results from the failed
  attempt are not merged.
- Standalone path-based helper calls own and close their own workbooks. An
  extraction does not cache workbooks across invocations or share them with
  another concurrent extraction.
- Success, fallback and exceptions release all successfully opened resources.
  Existing helper and pipeline overrides remain observable and select the
  compatibility path for `light` OOXML inputs.

## OOXML core extraction session

Issue #149 added the streaming `OoxmlExtractionSession` for core worksheet
data. Issue #150 makes it the default core reader for supported `light`
`.xlsx` / `.xlsm` extraction, with rich extraction sharing the same ZIP.
See [Core OOXML extraction session](ooxml-core-extraction.md) for supported
values, formulas, relationships, tables, lifecycle and limitations.

## Coordinate System

- Rows are 1-based
- Columns are 0-based

## Modes

- light: Skip COM. Supported `.xlsx` / `.xlsm` files use the direct OOXML
  pipeline for cells and pre-com artifacts plus best-effort shapes, connectors
  and charts. Unsupported options or direct-extraction failures restart the
  complete openpyxl compatibility pipeline. `.xls` behavior is unchanged.
- libreoffice: Seed the same OOXML rich baseline as `light`, then try LibreOffice enrichment; on failure, fall back to cells with pre-com artifacts preserved and keep any OOXML rich artifacts already extracted
- standard: Existing behavior (text-bearing shapes, charts if needed)
- verbose: All shapes + sizes, charts with sizes

## Import Boundaries

- Bare package/engine imports and CLI help keep their existing lazy contracts.
  The normal supported `.xlsx` / `.xlsm` light path imports none of openpyxl,
  pandas, SciPy or xlwings, nor concrete COM backends, ExStruct rendering or
  PDFium. Unsupported input, `colors_map` opt-in or legacy overrides may select
  the openpyxl compatibility path.
- Shared cell/map/merged-range types do not import concrete extraction backends.
  Compatibility exports and override call sites resolve the live implementation
  only when requested or executed.
- Openpyxl remains installed for compatibility fallback and paths that select
  it. NumPy remains required by the existing table-clustering contract.
- BIFF xlrd is imported only for `.xls` cell reading. Direct light OOXML table
  heuristics explicitly select Python BFS per call and do not import SciPy.
  Other modes and the openpyxl compatibility pipeline retain the configured
  `EXSTRUCT_BORDER_CLUSTER_BACKEND` behavior.

## Cell Extraction

- Read supported `.xlsx` / `.xlsm` light-mode cached cell values directly from
  OOXML. The compatibility pipeline retains openpyxl reading; `.xls` behavior
  remains unchanged. No production pandas reader is used.
- Ignore blank cells
- Normalize row data into `CellRow`
- Preserve existing numeric coercion (including zero-prefixed numeric strings), boolean/date/time text,
  row gaps, sheet ordering and the former reader's default missing-string tokens.
  Excel error cells are omitted as before. Missing tokens are matched before
  trimming, so padded tokens remain ordinary text.
- Hyperlinks use their external `target` and zero-based numeric-string column
  keys. Links on filtered cells are retained when their row contains another
  emitted value; a row containing only filtered values is not created for links.
- Formula text extraction remains separate from cell extraction. Cell values
  use stored formula results and do not calculate formulas.
- `.xls` direct reading adds no OOXML rich artifacts or COM-independent formula,
  hyperlink, color, merged-cell or table extraction guarantees.

## Table Extraction

- In light OOXML, combine explicit table definitions with the existing
  border-based table-candidate heuristics; preserve candidate output semantics.
- The openpyxl compatibility path retains openpyxl table definitions and border
  clusters.
- Preserve table_candidates even when COM is unavailable
- Path-based detection delegates to worksheet-based detection. Nested border
  scanning and multi-sheet detection must not reopen the workbook when the
  caller supplies a worksheet from the extraction session.
- COM table detection reuses a supplied session only when its resolved file
  path matches the COM workbook path. A different workbook uses standalone
  path-based detection, even when both workbooks contain the same sheet name.
- NumPy remains required. SciPy is optional through `exstruct[fast]` or `[all]`.
  Direct OOXML light table heuristics pass `cluster_backend="python"` per call,
  without reading or changing the environment variable. Other modes and the
  openpyxl compatibility path retain `EXSTRUCT_BORDER_CLUSTER_BACKEND`:
  `auto` attempts SciPy-backed labeling, `python` forces the existing Python
  BFS, and legacy `numpy` also attempts SciPy. Import or execution failure falls
  back to Python; unknown values retain existing `auto` behavior. Sparse-set BFS
  remains an evaluation candidate.

## Shapes / Arrows / SmartArt Extraction

What is extracted:

- Normalization of Type / AutoShapeType (`type` is kept for Shape only)
- Left/Top/Width/Height
- TextFrame2.TextRange.Text
- Arrow direction and connection information
- SmartArt layout/nodes/kids (nested structure)

Mode notes:

- `light` on `.xlsx/.xlsm` uses OOXML drawing parts for best-effort shapes/connectors and does not require COM or LibreOffice runtime
- `libreoffice` may refine geometry/ordering/connector matching beyond the OOXML baseline, but is no longer the only non-COM path to these artifacts

## Chart Extraction

What is extracted:

- ChartType (integer → string via XL_CHART_TYPE_MAP)
- Series / Axis Title / Axis Range
- Chart Title

Mode notes:

- `light` on `.xlsx/.xlsm` uses OOXML chart parts plus worksheet drawing anchors for best-effort chart metadata/placement
- `libreoffice` can still refine chart placement/confidence, but baseline metadata now exists without the LibreOffice runtime

## Print Areas / Auto Page Breaks

- Light OOXML reads print areas from workbook defined names. The compatibility
  path retains openpyxl extraction; COM supplements missing parts as before.
- auto_page_breaks are retrieved via COM only
- The extraction CLI always exposes auto page-break export syntax and validates the required mode/runtime at execution time instead of probing COM during parser construction

## Colors Map

- Prefer COM to include conditional formatting colors
- Overwrite with COM results when COM succeeds
- Use openpyxl results only when COM fails
- An explicit `colors_map` request on the direct light OOXML path selects the
  complete openpyxl compatibility pipeline.

## Error Handling / Fallback

- Direct light OOXML failure or an unsupported construct at any pipeline stage
  discards all partial OOXML data, closes the session, and restarts the complete
  openpyxl compatibility pipeline. The warning and state use
  `FallbackReason.OOXML_COMPATIBILITY = "ooxml_compatibility"` after a successful
  compatibility restart. If compatibility extraction itself takes an existing
  fallback, its later reason remains in state; the initial OOXML warning is
  still logged.
- A per-sheet drawing parse failure remains best-effort and skips affected
  drawing artifacts for that sheet while preserving healthy sheet output.
- COM / LibreOffice runtime failures retain their existing fallback behavior:
  keep safe cell/table and pre-com artifacts, plus rich OOXML artifacts already
  recovered where the current mode contract permits.
- Log fallback reasons uniformly via `FallbackReason`.
