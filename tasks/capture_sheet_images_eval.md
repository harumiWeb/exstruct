# capture_sheet_images Stability Evaluation

## Purpose

Define a repeatable evaluation set and procedure for MCP `exstruct_capture_sheet_images` (Experimental) before GA.

## Runtime Baseline

- OS: Windows (Excel desktop available)
- Evaluate both runtime profiles:
  - Profile A (current MCP default): `EXSTRUCT_RENDER_SUBPROCESS=1`
  - Profile B (in-process fallback): `EXSTRUCT_RENDER_SUBPROCESS=0`
- Common timeout:
  - `EXSTRUCT_MCP_CAPTURE_SHEET_IMAGES_TIMEOUT_SEC=180`

## Representative Workbook Set

Use the following local files:

1. `sample/basic/sample.xlsx`
2. `sample/formula/formula.xlsx`
3. `sample/flowchart/sample-shape-connector.xlsx`
4. `sample/smartart/sample_smartart.xlsx`
5. `sample/forms_with_many_merged_cells/ja_general_form/ja_form.xlsx`
6. `sample/forms_with_many_merged_cells/en_form_sf425/sample.xlsx`
7. `tests/assets/multiple_print_ranges_4sheets.xlsx`

## Evaluation Cases

For each workbook, run all applicable cases:

1. Full workbook export (`sheet`/`range` omitted)
2. Single sheet export (`sheet` only)
3. Minimal non-empty range export (`sheet` + `range=<one-cell non-empty address>`)
4. Sheet-qualified range export (`'Sheet Name'!A1:B2` when sheet name contains spaces)

Notes:
- Do not fix the minimal range to `A1:A1`; some workbooks can trigger Excel modal dialogs in unattended runs.
- For case 3, choose a one-cell range that is confirmed non-empty from the target sheet.

## Procedure

1. Start MCP server with target profile env.
2. Determine the case-3 cell before capture:
   - Use extraction output (or workbook inspection) and pick the first non-empty cell on the target sheet.
3. For each case, call `exstruct_capture_sheet_images` and record:
   - start/end timestamp
   - elapsed seconds
   - success/failure
   - error message (if failed)
   - generated image count
4. Repeat the whole matrix 3 times.
5. Compute:
   - total runs
   - success rate
   - p50/p95 elapsed time
   - failure type breakdown

Unattended-run rule:
- If Excel modal confirmation appears during a run, mark the run as invalid and re-run after adjusting the minimal range cell.

## Pass/Fail Gate (GA)

- Success rate >= 99% across all runs.
- No opaque timeout failures; all failures include actionable messages.
- No monotonic memory growth that indicates unacceptable leak behavior during repeated runs.

## Reporting Template

- Date:
- Environment:
- Total runs:
- Success rate:
- p50/p95 elapsed:
- Failures by type:
- Notes / next actions:
