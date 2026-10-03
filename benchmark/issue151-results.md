# Issue #151: COM-first extraction evaluation

## Outcome

Retain public file-first `standard` / `verbose` extraction. The bulk COM
candidate does not meet both performance and output compatibility requirements.
`ComBackend.extract_cells()` and the COM-first runner remain internal experiments.
ADR-0014 records the rationale; no public backend selector is added.

## Method and reproduction

Windows 11 (build 22631), Python 3.11.13, Excel 16.0, xlwings 0.37.4,
openpyxl 3.1.5, NumPy 2.4.6 and Pydantic 2.13.5. Base revision:
`d28341897fbed2b1ff70c385f51d550365d49130` plus this working-tree change.
SciPy was absent during the original timings; both runners used the existing
Python fallback. Installing optional test dependencies later does not change
the recorded run environment.

Generate fixtures in a new directory and run sequentially with no other Excel
instances open. Do not run the real COM tests concurrently with measurements.

```powershell
rtk uv run python -m benchmark.performance --generate-fixtures tasks/issue151-fixtures
rtk uv run python -m benchmark.issue151_com_first --input tasks/issue151-fixtures/small.xlsx tasks/issue151-fixtures/style-heavy.xlsx tasks/issue151-fixtures/large.xlsx tasks/issue151-fixtures/many-sheet.xlsx --output benchmark/baselines/issue151-2026-10-03-windows.json --repeats 3
rtk uv run python -m benchmark.issue151_com_first --input tasks/issue151-fixtures/small.xlsx tasks/issue151-fixtures/style-heavy.xlsx --output benchmark/baselines/issue151-2026-10-03-verbose-windows.json --repeats 3 --modes verbose
rtk uv run python -m benchmark.issue151_com_first --input tasks/issue151-fixtures/large.xlsx tasks/issue151-fixtures/many-sheet.xlsx --output benchmark/baselines/issue151-verbose-no-colors-windows.json --repeats 3 --modes verbose --no-colors
rtk uv run python -m benchmark.issue151_compatibility --output benchmark/baselines/issue151-compatibility-windows.json
rtk uv run python -m benchmark.issue151_report
```

Fixtures: small = 20 x 8 with a formula, merge and chart; medium/style-heavy =
300 x 20; large = 2,000 x 30; many-sheet = 24 sheets of 30 x 10. ZIP hashes
are recorded for reproduction. Default include flags are preserved for standard and the small/medium verbose runs. Large/many-sheet verbose runs explicitly disable colors to isolate the cell architecture; they are reported separately.
The candidate keeps the shared openpyxl table/merged/formula stages and existing
rich COM stages; it replaces cell reading with bounded `Value2` batches and
bulk `ISERROR` masks. It does not use per-cell COM value access.

Cold starts/quits a new app for each extraction. Warm deliberately opens/closes
the workbook in a controlled, running app for both runners. Already-open reuses
the open workbook without closing it. This warm injection is a benchmark
condition, not a new production app-reuse policy. Raw JSON separates acquisition,
cleanup, total time and the remaining extraction/model work. Setup of warm/open
scenarios is outside timed extraction. Excel app IDs are checked after cleanup.
Runner order alternates over three repeats. Timing excludes output hashing.
Medians and output equality below include only successful COM runs; failed runs
remain in raw JSON. If either runner has no successful sample, output equality
is unknown. This avoids mistaking fallback output for COM-reader compatibility.
These are local observations, not statistical confidence intervals or universal
claims about Excel performance. Python imports are not separately cold-started.

## Standard results

| Fixture | Scenario | Current ms | Candidate ms | Output equal | Failed COM |
| --- | --- | ---: | ---: | --- | ---: |
| small | cold | 2670.4 | 2645.2 | False | 0 |
| small | warm | 112.8 | 183.0 | False | 0 |
| small | already-open | 36.7 | 61.2 | False | 0 |
| style-heavy | cold | 2300.8 | 2417.2 | True | 0 |
| style-heavy | warm | 262.0 | 266.5 | True | 0 |
| style-heavy | already-open | 176.1 | 306.4 | True | 0 |
| large | cold | 3663.7 | 3677.5 | True | 1 |
| large | warm | 735.6 | 773.9 | True | 0 |
| large | already-open | 713.1 | 792.8 | True | 0 |
| many-sheet | cold | 3071.5 | 3710.4 | True | 0 |
| many-sheet | warm | 932.4 | 1771.5 | True | 0 |
| many-sheet | already-open | 1010.3 | 1303.7 | True | 0 |

One large/cold candidate run encountered COM RPC failure and returned the clean
file fallback. Its ~773 ms total is not evidence of faster successful extraction.
Excluding that failure removes the apparent cold/large advantage.

## Verbose default results

| Fixture | Scenario | Current ms | Candidate ms | Output equal | Failed COM |
| --- | --- | ---: | ---: | --- | ---: |
| small | cold | 3192.0 | 3308.2 | False | 0 |
| small | warm | 744.6 | 573.2 | False | 0 |
| small | already-open | 690.0 | 617.6 | False | 0 |
| style-heavy | cold | 28834.0 | 27933.2 | True | 0 |
| style-heavy | warm | 22258.6 | 26370.8 | True | 0 |
| style-heavy | already-open | 33104.4 | 22649.0 | True | 0 |

Small formula results differ despite some faster warm samples. Medium timings
are dominated by existing per-cell COM color extraction and fluctuate materially.
The original medium/already-open scenario overlapped an Excel test and was
discarded; the table uses its isolated rerun. Large default verbose was stopped
without retaining an unfinished timing because color extraction dominated.
Large/many-sheet verbose below disables colors explicitly; those timings are
not default verbose timings. A first isolation attempt overlapped another
Excel test and failed with RPC errors; that attempt was discarded and rerun
with no concurrent tests.

## Verbose with colors explicitly disabled

| Fixture | Scenario | Current ms | Candidate ms | Output equal | Failed COM |
| --- | --- | ---: | ---: | --- | ---: |
| large | cold | 3666.5 | 4113.3 | True | 0 |
| large | warm | 1920.5 | 1910.2 | True | 0 |
| large | already-open | 1416.4 | 1355.8 | True | 0 |
| many-sheet | cold | 3091.8 | 4084.2 | True | 0 |
| many-sheet | warm | 936.7 | 1653.4 | True | 0 |
| many-sheet | already-open | 737.9 | 1487.0 | True | 0 |

## Compatibility and reliability

The generated compatibility probe reproduces these differences:

| Case | Saved-file reader | Bulk COM reader |
| --- | --- | --- |
| Date/time | `2024-01-02 03:04:05` | `45293.12783564815` |
| Formula with absent cache | omitted | recalculated `3` |
| Unsaved edit | saved `label` | live `unsaved` |
| External hyperlink | retained | retained |
| Internal location-only link | omitted | omitted |
| Excel error | omitted | omitted |
| Integer matching COM error code | retained | retained |
| Boolean | string `True` | string `True` |

The small fixture differs because Excel supplies a recalculated formula value.
The compatibility fixture also confirms full output inequality while both COM
pipelines succeed. Dates require style-aware interpretation, including workbook
date systems, before Value2 can preserve the current contract. Selecting live
values would require a separate human-approved public contract decision.

Failure-injection tests cover startup, cells, rich extraction and final model
construction. They verify that fallback starts after COM ownership exits and
discards partial artifacts. Lifecycle tests cover open failure, extraction
failure, close failure, quit failure/owned kill, and existing borrowed workbook
preservation. A new app formerly leaked if `books.open` raised before `try`;
that defect is fixed independently of the experiment.

## Validation

- Non-COM suite with `mcp` and `fast` extras: 1,093 passed, 2 skipped,
  12 deselected before the additional batch-boundary and real-COM tests.
- Focused unit tests after batch, cleanup and review fixes: 24 passed.
- Ruff and strict mypy (89 source files) pass.
- Real Excel compatibility test: 1 passed with both pipelines succeeding.
  Pytest emitted Windows RPC exception diagnostics (`0x800706be`) even on
  the sequential rerun; exit code was 0 and no Excel app remained. This is
  output-compatibility evidence, not a clean runtime-stability result.
- Raw baseline JSON and the compatibility probe are retained under
  `benchmark/baselines/`; timing tables can be regenerated with the report runner.

## Reconsideration gate

Retain file-first because the candidate adds COM cost on many warm/many-sheet
paths and changes saved-value semantics. Table detection remains file-based:
the experiment provides no evidence supporting a COM heuristic rewrite.
Future candidates must preserve date/cache/unsaved semantics, verify owned and
borrowed cleanup, report successful and fallback timings separately, and show
repeatable gains across modes and Excel startup conditions before adoption.

## ADR workflow checks

Suggester: recommended (performance strategy retained); drafter: ADR-0014.
Linter: no findings (valid status, required sections, balanced consequences,
evidence triad and no supersession). Reviewer: ready, no design findings.
Reconciler: no policy drift or evidence gaps against ADR-0002/0011/0013,
specs, source and tests. Indexer: README, index.yaml and decision-map synchronized;
14 source paths and domain memberships verified. Remaining limits are local
timing variability and incomplete default large-verbose timings.
