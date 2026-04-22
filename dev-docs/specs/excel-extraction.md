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

## Coordinate System

- Rows are 1-based
- Columns are 0-based

## Modes

- light: Skip COM entirely; return cells+table_candidates as the base, along with pre-com artifacts according to their flags, and on `.xlsx/.xlsm` emit best-effort OOXML shapes/connectors/charts when present
- libreoffice: Seed the same OOXML rich baseline as `light`, then try LibreOffice enrichment; on failure, fall back to cells with pre-com artifacts preserved and keep any OOXML rich artifacts already extracted
- standard: Existing behavior (text-bearing shapes, charts if needed)
- verbose: All shapes + sizes, charts with sizes

## Cell Extraction

- Load with pandas `read_excel(header=None, dtype=str)`
- Ignore blank cells
- Normalize row data into `CellRow`

## Table Extraction

- Merge openpyxl table definitions + border clusters
- Preserve table_candidates even when COM is unavailable

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
