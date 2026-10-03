# Excel Extraction Specification — ExStruct

This document summarizes the current specification for Excel extraction processing.

## Overall Flow

1. `resolve_extraction_inputs` normalizes include_* and mode
2. Pre-com (openpyxl) retrieves cells/print_areas/formulas_map/colors_map/merged_cells
3. `light` uses the OOXML rich backend for best-effort shapes/connectors/charts on `.xlsx/.xlsm`
4. `libreoffice` seeds the same OOXML rich baseline and then uses the LibreOffice backend for optional best-effort enrichment on `.xlsx/.xlsm`
5. `standard` / `verbose` use COM (xlwings) to retrieve shapes/charts/auto_page_breaks
6. When COM succeeds, colors_map is overwritten with COM results
7. When the rich backend fails, cells+table_candidates are preserved, and pre-com artifacts (print_areas / formulas_map / colors_map / merged_cells) are also retained according to their flags

## Workbook Resource Lifetime

- One extraction invocation owns an `OpenpyxlExtractionSession`. Compatible
  OOXML stages share one regular `data_only=True`, `read_only=False` workbook,
  including hyperlink and per-sheet table detection, until final modeling ends.
- An openpyxl formula map opens a second regular `data_only=False` workbook
  lazily. Disabled formula extraction and COM-only `.xls` formula extraction
  must not open that variant.
- Without helper overrides, an OOXML extraction loads one openpyxl workbook, or
  two when openpyxl formula extraction is enabled, independent of sheet count.
  OOXML drawing ZIP reads are separate and are not included in that loader count.
- Standalone path-based helper calls own and close their own workbooks. An
  extraction does not cache workbooks across invocations or share them with
  another concurrent extraction.
- Success, fallback and exceptions release all successfully opened variants.
  Existing helper overrides remain observable and may intentionally replace
  the optimized session path.

## Internal OOXML core alternative

Issue #149 adds a streaming `OoxmlExtractionSession` for core worksheet data
and optional rich extraction through one ZIP. It is available for parity and
future migration work; the pipeline above continues to use openpyxl.
See [Core OOXML extraction session](ooxml-core-extraction.md) for supported
values, formulas, relationships, tables, lifecycle and limitations.

## Coordinate System

- Rows are 1-based
- Columns are 0-based

## Modes

- light: Skip COM entirely; return cells+table_candidates as the base, along with pre-com artifacts according to their flags, and on `.xlsx/.xlsm` emit best-effort OOXML shapes/connectors/charts when present
- libreoffice: Seed the same OOXML rich baseline as `light`, then try LibreOffice enrichment; on failure, fall back to cells with pre-com artifacts preserved and keep any OOXML rich artifacts already extracted
- standard: Existing behavior (text-bearing shapes, charts if needed)
- verbose: All shapes + sizes, charts with sizes

## Import Boundaries

- Bare package/engine imports and CLI help keep their existing lazy contracts.
  Real `.xlsx/.xlsm` light extraction loads neither xlwings nor concrete COM
  backends, COM shape/chart modules, ExStruct rendering or PDFium.
- Shared cell/map/merged-range types do not import concrete extraction backends.
  Compatibility exports and override call sites resolve the live implementation
  only when requested or executed.
- Openpyxl and NumPy remain available when extraction needs them. Openpyxl may
  import optional Pillow itself when installed; light extraction must also work
  without Pillow. This transitive import does not activate ExStruct rendering.
- BIFF xlrd is imported only for `.xls` cell reading. SciPy labeling is imported
  only when accelerated clustering is attempted; explicit Python clustering
  does not import SciPy.

## Cell Extraction

- Read `.xlsx` / `.xlsm` cached cell values directly with openpyxl and `.xls`
  cached values directly with xlrd; no production pandas reader is used.
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

- Merge openpyxl table definitions + border clusters
- Preserve table_candidates even when COM is unavailable
- Path-based detection delegates to worksheet-based detection. Nested border
  scanning and multi-sheet detection must not reopen the workbook when the
  caller supplies a worksheet from the extraction session.
- COM table detection reuses a supplied session only when its resolved file
  path matches the COM workbook path. A different workbook uses standalone
  path-based detection, even when both workbooks contain the same sheet name.
- NumPy remains required. SciPy is optional through `exstruct[fast]` or `[all]`.
  `EXSTRUCT_BORDER_CLUSTER_BACKEND=auto` attempts SciPy-backed labeling, while
  `python` forces the existing Python BFS. The legacy `numpy` value also attempts
  SciPy; import or execution failure falls back to Python. Unknown values retain
  the existing `auto` behavior. Sparse-set BFS remains an evaluation candidate.

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

- print_areas are retrieved via pre-com (openpyxl); COM only supplements missing parts
- auto_page_breaks are retrieved via COM only
- The extraction CLI always exposes auto page-break export syntax and validates the required mode/runtime at execution time instead of probing COM during parser construction

## Colors Map

- Prefer COM to include conditional formatting colors
- Overwrite with COM results when COM succeeds
- Use openpyxl results only when COM fails

## Error Handling / Fallback

- When COM / LibreOffice is unavailable or raises an exception, return cells+table_candidates, preserve the pre-com artifacts (print_areas / formulas_map / colors_map / merged_cells) according to their flags, and keep any OOXML rich artifacts that were already extracted before the failing enrichment step
- Log fallback reasons uniformly via `FallbackReason`
