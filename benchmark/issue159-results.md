# Issue #159: COM color extraction benchmark

## Method and reproduction

This benchmark compares `cells.extract_sheet_colors_map_com(...)` with the
previous manual path: workbook preparation, preparation of each sheet, and
`_extract_sheet_colors_com(...)`. Both paths use the same already-open Excel
workbook and the same `include_default_background` / `ignore_colors` inputs.
Every pair checks exact `WorkbookColorsMap` equality. Timings cover only the
extraction call and its required display-format preparation.

The runner generates and retains deterministic-shape `.xlsx` fixtures for a
static 300 x 20 grid, sparse conditional formatting limited to 30 cells, dense
conditional formatting, overlapping conditional rules with formulas,
theme/indexed fills, merged cells, and eight populated sheets capped at 30 rows
each. It runs sequentially in one owned,
hidden Excel application, opens one workbook at a time, alternates reference
and hybrid order across repeats, and checks that the Excel process is closed at
the end. Each fixture's SHA-256, each sample, the per-call wrapper counts, the
hybrid logger's DEBUG records, and exact comparison outcomes are retained in
the output JSON. The run covers both values of `include_default_background`
and both `ignore_colors=None` and `ignore_colors={"FF0000"}`.

On Windows, close other Excel instances, then run the benchmark by itself. Do
not overlap it with Excel tests or another COM benchmark.

```powershell
rtk uv run python -m benchmark.issue159_colors --output benchmark/baselines/issue159-colors-windows.json --repeats 3 --rows 300
```

For an initial short run, select the static and sparse-range cases:

```powershell
rtk uv run python -m benchmark.issue159_colors --output benchmark/baselines/issue159-colors-smoke-windows.json --repeats 1 --rows 300 --cases static sparse_cf
```

By default the generated workbooks are kept in a sibling directory named
`issue159-colors-windows-fixtures`. Pass `--fixtures-dir PATH` to choose another
location. Use `--cases` with one or more fixture names to run a subset. The
eight-sheet case uses `min(--rows, 30)` rows per sheet to bound its COM cost.
The script prints flushed fixture, configuration, repeat, and strategy
progress, and writes results as each configuration finishes so an interrupted
run retains completed comparisons.
It also probes `Range("A1:B1").DisplayFormat.Interior.Color` in the overlapping
fixture and records that value beside the two individual cell values.
The probe also records the workbook's COM `Saved` flag immediately before and
after its calculation call.

The wrapper count measures calls through
`cells._get_display_format_color` for both implementations. Legacy measurements
also retain the existing ContextVar counter when available. Historical hybrid
records contain a false zero in `cells_context_display_format_calls`: the
hybrid extractor resets that per-sheet counter, so the final value is not an
aggregate. Do not use that field in existing raw JSON; the historical data is
left unchanged. `wrapper_display_format_calls` and captured hybrid DEBUG
`display_format_calls` remain the authoritative counts and are unchanged. New
hybrid records omit `cells_context_display_format_calls`. Hybrid DEBUG records
are captured verbatim with structured extras and parsed metric fields,
including `used_cells`, `conditional_candidates`, `display_format_calls`, and
any logged fallback duration fields.

The valid final CF fixtures write solid-fill `fgColor` and `bgColor` as the
same color, use OOXML-standard `FormulaRule` expressions without a leading `=`,
and keep Excel `ScreenUpdating` enabled during the runtime measurements.

## Results

### Initial 300-row run: valid non-CF cases

Source: [`issue159-colors-windows.json`](baselines/issue159-colors-windows.json),
recorded 2026-10-04 on Windows 10 build 22631, Python 3.11.13, Excel 16.0,
xlwings 0.37.4. It ran one sample for each of four configurations (both
`include_default_background` values, with and without ignored red). Times below
are the min–max across those four samples, **not medians**. All four exact map
comparisons passed for each listed fixture.

| Fixture | Shape | Reference time (ms) | Hybrid time (ms) | `_get_display_format_color` calls (reference / hybrid) | Exact maps |
| --- | --- | ---: | ---: | ---: | --- |
| Static fills | 300 x 20 | 14,079.4–16,323.6 | 80.3–106.1 | 6,000 / 0 | 4/4 |
| Theme/indexed fills | 300 x 20 | 13,167.0–15,597.6 | 1,353.4–2,321.9 | 6,000 / 649 | 4/4 |
| Merged cells | 300 x 20 | 16,070.4–17,275.6 | 13,143.1–17,053.0 | 6,000 / 6,000 | 4/4 |
| Multiple sheets | 8 sheets x 30 x 20 | 10,136.3–13,358.2 | 109.4–174.0 | 4,800 / 0 | 4/4 |

The eight-sheet input is capped at 30 rows per sheet. The initial run's
sparse-CF, dense-CF, and overlapping-CF results are invalid for color-rendering
evidence: their saved differential-format `bgColor` did not match `fgColor`,
the `FormulaRule` formulas had a leading `=`, and the default-false maps
contained zero cells. A formula-only rerun removed the leading `=` but still
rendered no CF colors ([`issue159-colors-final-smoke-windows.json`](baselines/issue159-colors-final-smoke-windows.json));
that intermediate run is also excluded. Passing fixtures match `bgColor` to
`fgColor`, use the standard formula form, and run with `ScreenUpdating` enabled.
Earlier samples remain in raw JSON but are excluded from timing and equality
conclusions; equality of two empty maps does not establish CF correctness.

### Corrected CF rendering probe: 10 rows

Source: [`issue159-colors-cf-probe-windows.json`](baselines/issue159-colors-cf-probe-windows.json),
Excel 16.0, one repeat, all four configurations per fixture. The corrected
OOXML formulas omit the leading `=`. Each exact comparison passed and the
expected real colors were observed wherever they were not intentionally
ignored.

| Fixture | Rendered color evidence | Exact maps | Hybrid path and wrapper calls | Recorded time (reference / hybrid, ms) |
| --- | --- | --- | --- | ---: |
| Sparse CF, `D2:D10` | 5 red cells; ignored-red cases correctly omit red | 4/4 | Optimized path, no fallback; 200 / 9 | 344.2–906.3 / 33.2–38.2 |
| Dense CF, `A1:T10` | 100 blue cells in all four configurations | 4/4 | Optimized path, no fallback; 200 / 200 | 362.0–928.3 / 423.1–625.6 |
| Overlapping/formula CF, `A1:T10` | 21 red and 37 blue cells; ignoring red leaves 37 blue | 4/4 | Conservative legacy fallback in all configurations; 200 / 200 | 433.3–625.7 / 407.1–894.8 |

This small probe confirms rendered CF color visibility, not 300-row performance.
For the overlap fixture, the workbook's COM `Saved` flag changed from `True`
before calculation to `False` afterward. The hybrid path detected the unsaved
workbook state and used the legacy fallback; these samples therefore validate
fallback correctness and must not be presented as optimized-path timings.

The mixed-range probe returned scalar `0` for
`Range("A1:B1").DisplayFormat.Interior.Color`, while the individual cells
returned `255` and `12611584`. The cell colors differ, and the range result
matches neither as a matrix; this confirms that a mixed `Range.DisplayFormat`
color read cannot supply per-cell color values. Microsoft documents
`Range.DisplayFormat` as a read-only display-format object affected by
conditional formatting ([Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.range.displayformat));
the scalar result and distinct per-cell values above are the direct Excel
runtime evidence.

### Final-source 300-row medians

Source: [`issue159-colors-final-windows.json`](baselines/issue159-colors-final-windows.json),
recorded 2026-10-04 00:55 UTC. This run contains 48 measurements across static
and sparse-CF fixtures: three sequential repeats for each of four configurations
per fixture. All 24 reference/hybrid map comparisons are exact, and all
non-ignored CF visibility checks passed.

| Fixture | Include default | Ignored colors | Reference median (ms) | Hybrid median (ms) | Wrapper calls (reference / hybrid) | CF visibility / map |
| --- | --- | --- | ---: | ---: | ---: | --- |
| Static | False | None | 16,503.2 | 100.4 | 6,000 / 0 | Exact |
| Static | False | `FF0000` | 15,497.6 | 86.1 | 6,000 / 0 | Exact |
| Static | True | None | 15,774.8 | 89.2 | 6,000 / 0 | Exact |
| Static | True | `FF0000` | 18,669.3 | 84.3 | 6,000 / 0 | Exact |
| Sparse CF (`D2:D31`) | False | None | 17,043.7 | 144.2 | 6,000 / 30 | 15 red cells; passed |
| Sparse CF (`D2:D31`) | False | `FF0000` | 16,796.2 | 200.9 | 6,000 / 30 | Red intentionally filtered; exact |
| Sparse CF (`D2:D31`) | True | None | 12,785.3 | 187.5 | 6,000 / 30 | 15 red cells; passed |
| Sparse CF (`D2:D31`) | True | `FF0000` | 15,442.7 | 207.6 | 6,000 / 30 | Red intentionally filtered; exact |

Main reports the full test suite at 1,132 passed, including 28 new regressions
for merged/table/row/column fallback and auto/gradient candidates. Ruff and
mypy for 90 source files also passed.

## Interpretation

Final-source 300-row medians show exact output equality for static and sparse
CF across all four configurations. The hybrid path makes 30 wrapper calls for
the 30-cell sparse CF range, versus 6,000 calls for the reference path, while
still returning the 15 rendered red cells when red is not ignored. Keep the
initial and formula-only CF samples excluded; report overlap-fixture timings
separately because its optimized path correctly falls back when Excel marks the
workbook unsaved after calculation.
