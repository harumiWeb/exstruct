# ExStruct Test Requirements Specification

Version: 0.5
Status: Required for release

This document summarizes the formal test requirements for all ExStruct functionality. It serves as the foundation for AI agents and human developers to design automated and manual tests.
Overall code coverage must be **80% or higher**.

---

# 1. Coverage Categories

1. Cell extraction
2. Shape extraction
3. Arrow / direction estimation
4. Chart extraction
5. Layout integration
6. Pydantic validation
7. Output (JSON/YAML/TOON)
8. CLI
9. Error handling / fail-safe
10. Regression
11. Performance / memory

---

# 2. Functional Requirements

## 2.1 Cell extraction

- [CEL-01] Exclude empty cells and output only non-empty cells in `c`
- [CEL-02] Row number `r` is 1-based
- [CEL-03] Column keys are 0-based indexes `"0"`, `"1"`, ...
- [CEL-04] Cells containing newlines and tabs can also be read correctly
- [CEL-05] Preserve Unicode (Japanese, emoji, surrogate pairs)
- [CEL-06] Force `dtype=str` when reading with pandas
- [CEL-07] No performance degradation even at full-sheet scale
- [CEL-08] `_coerce_numeric_preserve_format` correctly determines int/float/non-numeric values
- [CEL-09] `detect_tables_openpyxl` detects openpyxl Table objects
- [CEL-10] `CellRow.links` is output when mode=verbose or `include_cell_links=True`
- [CEL-11] detect_tables switches code paths based on the .xlsx/.xls extension and whether openpyxl is available

## 2.1.1 Cell background colors

- [COL-01] Extract `colors_map` only when `include_colors_map=True`
- [COL-02] Do not output `FFFFFF` when `include_default_background=False`
- [COL-03] Exclude target colors when `ignore_colors` is specified (normalize `#` prefixes and letter case)
- [COL-04] When using COM, reference `DisplayFormat.Interior` and retrieve values including conditional formatting
- [COL-05] `_normalize_color_key` / `_normalize_rgb` normalize ARGB/#/auto/theme/indexed

## 2.1.2 Merged cells

- [MRG-01] Extract `merged_cells` only in standard/verbose (`light` yields empty)
- [MRG-02] Output using 1-based row / 0-based column coordinates
- [MRG-03] `v` is the top-left cell value of the merged range (normalize None / empty string to a single space `" "`)
- [MRG-04] Preserve all entries even when multiple ranges exist

## 2.2 Shape extraction

- [SHP-01] Normalize the type of AutoShape
- [SHP-02] Retrieve TextFrame correctly
- [SHP-02a] Keep `type` only for Shape; do not output it for Arrow/SmartArt
- [SHP-03] Fields `w` and `h` are null only when they cannot be retrieved
- [SHP-04] Apply a consistent expansion policy for grouped shapes
- [SHP-05] Retrieve coordinates `l`,`t` as integers, unaffected by zoom
- [SHP-07] Rotation angle matches Excel
- [SHP-09] begin/end_arrow_style matches Excel ENUM
- [SHP-10] Normalize direction to 8 compass directions
- [SHP-11] Shapes without text use text=""
- [SHP-12] Retrieve multi-paragraph text as well

## 2.2.1 SmartArt extraction

- [SHP-SA-01] SmartArt must always output `layout`
- [SHP-SA-02] SmartArt nodes are output to `nodes` as a nested structure
- [SHP-SA-03] Node children are represented by `kids` (do not output level)
- [SHP-SA-04] When SmartArt is present, it is identifiable by `kind="smartart"`

## 2.3 Arrow direction estimation

- [DIR-01] 0° ±22.5° → "E"
- [DIR-02] 45° ±22.5° → "NE"
- [DIR-03] 90° ±22.5° → "N"
- [DIR-04] 135° ±22.5° → "NW"
- [DIR-05] 180° ±22.5° → "W"
- [DIR-06] 225° ±22.5° → "SW"
- [DIR-07] 270° ±22.5° → "S"
- [DIR-08] 315° ±22.5° → "SE"
- [DIR-09] Boundary angles are rounded according to the specification

## 2.4 Chart extraction

- [CH-01] ChartType is normalized by `XL_CHART_TYPE_MAP`
- [CH-02] Retrieve the title (null if absent)
- [CH-03] Retrieve y_axis_title (empty string if absent)
- [CH-04] Axis min/max are float
- [CH-05] Unset axes are empty lists
- [CH-06] Output name_range as a reference formula (example: `=Sheet1!$B$1`)
- [CH-06a] Series names that are string literals are stored in name_literal
- [CH-07] Output x_range as a reference formula
- [CH-08] Output y_range as a reference formula
- [CH-09] Parse major chart types (scatter, bar, etc.)
- [CH-10] On failure, leave a message in `error` and preserve the chart
- [CH-11] Locale-specific semicolon separators can also be parsed

## 2.5 Layout integration

- [LAY-01] Link Shape text to the row it belongs to
- [LAY-02] Simplified column-based linkage (skip; not yet implemented)
- [LAY-03] Preserve order even when multiple shapes are in one row
- [LAY-04] Return an empty list when there are no shapes

---

# 3. Model Validation Requirements

- [MOD-01] All models inherit from `BaseModel`
- [MOD-02] Types match `dev-docs/specs/data-model.md` exactly
- [MOD-03] Optional fields default to None when unspecified
- [MOD-04] Numeric values are normalized to int/float
- [MOD-05] Invalid values in a direction Literal raise ValidationError
- [MOD-06] rows/shapes/charts/tables default to empty lists
- [MOD-07] WorkbookData provides `__getitem__` and ordered iteration
- [MOD-08] PrintArea satisfies row=1-based / column=0-based

---

# 4. Output Requirements (JSON/YAML/TOON)

- [EXP-01] None/empty string/empty list/empty dict are removed by `dict_without_empty_values`
- [EXP-02] JSON output is UTF-8
- [EXP-03] YAML output uses sort_keys=False
- [EXP-04] TOON output is generated correctly
- [EXP-05] No destructive changes in the `WorkbookData` → JSON → `WorkbookData` round trip
- [EXP-06] `export_sheets` outputs files per sheet
- [EXP-07] `to_json` supports pretty/indent
- [EXP-08] `save(path)` determines the format by extension and raises ValueError for unsupported extensions
- [EXP-09] `to_yaml` / `to_toon` raise MissingDependencyError when dependencies are not installed
- [EXP-10] Exclude target fields with include\_\* in OutputOptions, and do not output empty lists
- [EXP-11] Output files per print area with `print_areas_dir` / `save_print_area_views` (do not write if there are no ranges)
- [EXP-12] PrintAreaView keeps only rows within the range and excludes cells/links outside the range
- [EXP-13] PrintAreaView includes only table_candidates fully contained within the range
- [EXP-14] With normalize=True, rebase row and column indexes to the print-area origin
- [EXP-15] When include_print_areas=False, do not output even if `print_areas_dir` is set
- [EXP-16] PrintAreaView includes only shapes/charts intersecting the range, and treats shapes with unknown size as points
- [EXP-17] Chart.w/h are output in verbose; in standard they are controlled by `include_chart_size`
- [EXP-18] Shape.w/h are controlled by `include_shape_size`; the default True applies only in verbose
- [EXP-19] When `auto_page_breaks_dir` is specified, retrieve `auto_print_areas` with include_auto_page_breaks=True (COM required)
- [EXP-20] export_auto_page_breaks raises an exception if auto_print_areas is empty, and writes only when it is present
- [EXP-21] save_auto_page_break_views saves auto_print_areas using unique keys such as Sheet1#auto#1
- [EXP-22] serialize_workbook raises SerializationError for unsupported formats
- [EXP-23] The export/process API correctly outputs when str/Path is passed to output_path/sheets_dir/print_areas_dir/auto_page_breaks_dir
- [EXP-24] `fmt="yml"` is treated as yaml, and the extension becomes `.yaml`
- [EXP-25] `include_merged_cells=False` in OutputOptions excludes `merged_cells`

---

# 5. CLI Requirements

- [CLI-01] `exstruct extract file.xlsx` succeeds
- [CLI-02] `--format json/yaml/toon` works
- [CLI-03] `--image` outputs PNG (Excel COM required, disallowed in `mode=libreoffice`)
- [CLI-04] `--pdf` outputs PDF (Excel COM required, disallowed in `mode=libreoffice`)
- [CLI-05] Exit safely even when an invalid path is provided (do not crash)
- [CLI-06] Output error messages to stdout/stderr
- [CLI-07] `--print-areas-dir` outputs print-area files, and skips when include_print_areas=False
- [CLI-08] stdout output remains UTF-8 even in Windows cp932 environments (e.g., `PYTHONIOENCODING=cp932`)

---

# 6. Error Handling Requirements

- [ERR-01] The process does not crash even on xlwings COM errors
- [ERR-02] Preserve other elements even when shape extraction fails
- [ERR-03] On chart extraction failure, record it in Chart.error
- [ERR-04] Do not raise an exception for broken reference ranges; record null/error
- [ERR-05] Output a message and exit when the Excel file cannot be opened
- [ERR-06] Do not miss openpyxl `_print_area` settings during extraction
- [ERR-07] If auto_print_areas is empty, export_auto_page_breaks raises PrintAreaError (ValueError-compatible)
- [ERR-08] If YAML/TOON dependencies are missing, MissingDependencyError provides installation guidance
- [ERR-09] Raise OutputError on write failure, and preserve the exception in the **cause**

---

# 7. Regression Requirements

- [REG-01] The JSON structure of existing fixtures matches past versions
- [REG-02] Detect model key deletion / renaming as breaking changes
- [REG-03] Detect changes to the direction estimation algorithm
- [REG-04] ChartSeries reference parsing matches past results

---

# 8. Non-functional Requirements

- Performance / memory targets will be added when separately defined

---

# 9. Mode / Integration Requirements

- [MODE-01] CLI `--mode` / API `extract(..., mode=)` accepts light/libreoffice/standard/verbose (default: standard)
- [MODE-02] light: cells + tables only, shapes/charts empty, no COM
- [MODE-03] standard: existing behavior (text-bearing shapes / arrows, and charts if COM is enabled)
- [MODE-04] verbose: all shapes (with size) and charts (with size, except for specific exclusions)
- [MODE-05] process_excel propagates mode when combined with PDF/image options
- [MODE-05a] `process_excel(..., mode="libreoffice")` rejects `pdf` / `image` / `auto_page_breaks_dir` early with `ConfigError`
- [MODE-05b] For `mode="libreoffice"` Python runtime detection, success of the bundled bridge `--probe` is the acceptance condition; incompatible `EXSTRUCT_LIBREOFFICE_PYTHON_PATH` settings are handled early as unavailable/incompatible errors
- [MODE-05c] The required Linux smoke job in GitHub Actions runs `pytest.mark.libreoffice` smoke without skip on `ubuntu-24.04` + `libreoffice` + `python3-uno`
- [MODE-05d] The Windows smoke job in GitHub Actions uses `windows-2025` + `libreoffice-fresh`, prioritizes `soffice.com` for `EXSTRUCT_LIBREOFFICE_PATH` (fallback to `soffice.exe` if not present), and runs `pytest.mark.libreoffice` smoke without skip
- [MODE-06] In standard, existing fixtures do not regress and unnecessary shapes do not increase
- [MODE-07] An invalid mode errors before processing starts
- [INT-01] On COM open failure, fall back to cells + table_candidates
- [INT-02] Preserve print_areas even during COM fallback
- [IO-05] `dict_without_empty_values` removes None/empty list/empty dict and preserves nesting
- [RENDER-01] PDF/PNG smoke tests for Excel+COM+pypdfium2 (env ON/OFF)
- [MODE-08] In light, extract print_areas with openpyxl and exclude them from default output (automatic determination)

## 2.6 Pipeline

- [PIPE-01] build_pre_com_pipeline includes only the required steps according to include_* and mode
- [PIPE-02] build_cells_tables_workbook reflects print_areas conditionally and preserves table_candidates
- [PIPE-03] resolve_extraction_inputs resolves include_* using mode defaults
- [PIPE-04] run_extraction_pipeline attempts COM and falls back to cells+tables on failure
- [PIPE-05] colors_map is overwritten with COM results when COM succeeds, and openpyxl is used only on failure
- [PIPE-06] print_areas preserves openpyxl results, and COM supplements only missing parts
- [PIPE-07] PipelineState holds com_attempted/com_succeeded/fallback_reason
- [PIPE-08] Do not include the COM step for auto_page_breaks when include_auto_page_breaks=False
- [PIPE-09] Do not include the extraction step for merged_cells when include_merged_cells=False
- [PIPE-MOD-01] build_workbook_data builds WorkbookData/SheetData from raw containers
- [PIPE-MOD-02] collect_sheet_raw_data collects extracted data into raw containers

## 2.7 Backend

- [BE-01] OpenpyxlBackend switches cell extraction paths depending on whether include_links is enabled
- [BE-02] OpenpyxlBackend continues with an empty list when table detection fails
- [BE-03] ComBackend returns None when colors_map extraction fails
- [BE-04] OpenpyxlBackend returns None when colors_map extraction fails
- [BE-05] ComBackend continues with an empty map when print_areas extraction fails
- [BE-06] OpenpyxlBackend continues with an empty map when merged_cells extraction fails
- [BE-07] Unimplemented merged_cells in ComBackend raises NotImplementedError

## 2.8 Ranges

- [RNG-01] parse_range_zero_based can normalize sheet-qualified ranges such as "Sheet1!A1:B2"

## 2.9 Table Detection

- [TBL-01] Rectangular merging does not consolidate rectangles in a containment relationship
- [TBL-02] Can generate table candidate range strings from a value matrix

## 2.10 Workbook

- [WB-01] openpyxl_workbook calls close regardless of whether an exception occurs
- [WB-02] openpyxl_workbook sets filters to suppress known openpyxl warnings
- [WB-03] _find_open_workbook tolerates fullname/resolve exceptions and returns None
- [WB-04] _find_open_workbook returns None on a top-level exception
- [WB-05] xlwings_workbook does not start App if an existing workbook is found

## 2.11 Logging

- [LOG-01] log_fallback outputs a warning log including the reason code

## 2.12 Integration/E2E

- [E2E-01] The full flow light extraction → serialize_workbook → export_sheets succeeds
- [E2E-02] Engine.process can output JSON to a stream when output_path=None

---

# 10. COM Test Operations (Local Manual)

Because CI cannot run Excel COM, COM tests are run manually on local environments.
In Codecov, unit and com are separated, and com is maintained with carryforward.

## 10.1 Local execution procedure

- unit (CI equivalent): `task test-unit`
- COM: `task test-com`
- LibreOffice smoke (Linux/Windows CI equivalent): `RUN_LIBREOFFICE_SMOKE=1 pytest tests/core/test_libreoffice_smoke.py -m libreoffice -q`

## 10.2 Codecov manual upload (optional)

Set `CODECOV_TOKEN` and `CODECOV_SHA` for manual upload.

- unit upload: `codecov-cli upload-process -f coverage.xml -F unit -C %CODECOV_SHA% -t %CODECOV_TOKEN%`
- COM upload: `codecov-cli upload-process -f coverage.xml -F com -C %CODECOV_SHA% -t %CODECOV_TOKEN%`
