# Issues 145-148 integrated extraction measurements

## Method and compatibility

`baselines/issues145-148-2026-10-03-windows.json` compares base `5bcb0e4` and
integrated extraction code `10f53cf` using the same locked Windows Python 3.11.13
interpreter and six deterministic fixture hashes. Baseline observations were
collected first; final after observations were rerun following the sparse-reader
refinement. Each run has one first observation and three warm repeats, plus
separate fresh-process import/CLI/cold extraction probes and an instrumented
profile. Operations run serially. Source `dirty=true` includes pending
measurement artifacts and smoke-script formatting; production extraction code
is committed at the recorded revision. Local interpreter paths are redacted.

Independent extractions compared exact compact JSON bytes before and after for
all twelve input/mode combinations; every output hash matched. All standard
profiles recorded COM attempted and succeeded with no fallback. Default profiles
load one shared openpyxl workbook after integration; the old implementation
loaded four in light and five in standard, rising to 50/51 for 24 sheets.
Formula-enabled tests require two variants independently of sheet count.
Openpyxl loader counts exclude Excel COM and OOXML drawing ZIP reads.

## Observations

Warm total includes extraction and JSON serialization. Cold total is fresh child
wall time including startup, imports and teardown. RSS is Python lifetime peak
working set after uninstrumented runs, excluding Excel child processes and later
instrumentation. Times are ms and RSS is MiB.

| Input | Mode | Warm before / after | Cold before / after | Peak RSS before / after |
| --- | --- | ---: | ---: | ---: |
| large | light | 1595.9 / 596.3 | 3224 / 1612 | 189.1 / 109.3 |
| large | standard | 5605.5 / 2495.0 | 5848 / 4126 | 187.6 / 165.9 |
| many-sheet | light | 4101.8 / 119.5 | 5120 / 1030 | 133.2 / 75.6 |
| many-sheet | standard | 6133.0 / 2485.3 | 8134 / 4638 | 137.8 / 129.3 |
| small | light | 23.1 / 7.4 | 1553 / 851 | 116.6 / 64.3 |
| small | standard | 1961.2 / 2500.5 | 3359 / 3783 | 122.1 / 121.0 |
| sparse | light | 70.2 / 41.4 | 2173 / 1338 | 126.2 / 71.1 |
| sparse | standard | 1660.3 / 2382.0 | 3716 / 4639 | 130.8 / 130.8 |
| style-heavy | light | 535.4 / 178.4 | 1905 / 1109 | 128.1 / 73.9 |
| style-heavy | standard | 1907.8 / 2002.5 | 4322 / 5224 | 135.3 / 132.0 |
| table-heavy | light | 242.1 / 71.9 | 2301 / 1166 | 129.9 / 74.5 |
| table-heavy | standard | 2325.3 / 2256.9 | 3800 / 3959 | 136.3 / 131.9 |

All six light workloads improved warm/cold times and Python peak RSS in this
run. COM dominates smaller standard workloads; the initial integrated pass
improved small/sparse, whereas the final pass above regressed. Neither pass
establishes a universal standard-mode improvement. A paired five-repeat recheck
is retained in `baselines/issues145-148-standard-recheck-2026-10-03-windows.json`
so unfavorable observations remain visible. No strict shared-runner time gate
or broad statistical claim is made.

The paired recheck measured small standard at 1885.0 / 1894.5 ms (before/after)
and sparse standard at 2121.0 / 2207.2 ms, with COM success throughout. This
reduces the apparent regression relative to the earlier separated observations;
the remaining differences must still be read as local timings, not a guarantee.

The first integrated sparse worksheet measurement regressed from 70 to 154 ms
because regular `iter_rows()` materialized every blank in its rectangular
extent. The final reader scans sorted existing worksheet cell coordinates,
retaining row gaps, column keys and link semantics without inserting blanks.
The regression test asserts both exact output and unchanged sparse storage.
Read-only standalone readers continue to stream `iter_rows()` and reset declared
dimensions, preserving the previous reader's behavior for malformed bounds.

## Verification

- Final non-COM/non-render suite after the backend-alias review fix: 1017 passed,
  1 LibreOffice smoke skipped, 11 deselected; coverage 83.09%, exceeding 80%.
- Actual Excel COM tests: eight passed on this Windows host. Pywin32 emitted
  RPC diagnostics during that run; successful process exit and test results
  are distinct from those diagnostics.
- Actual forced LibreOffice smoke: one passed on this host.
- Actual PDF/PNG render suite: three passed. RPC diagnostics also appeared in
  that run; output PDF/PNG assertions and the test process succeeded.
- Ruff, format and strict mypy passed; session tests cover close on
  success/fallback/error, formula variants, helper overrides and concurrency.
- A built base wheel installed normally into an independent venv without pandas
  or SciPy; bordered light JSON matched for `python`, `auto`, and `numpy`.
  Actual BIFF light extraction also succeeded in that installed environment.
- A separate clean wheel install with `[fast]` installed SciPy and executed real
  accelerated labeling successfully, while pandas remained absent.

Timing datasets target `10f53cf`; subsequent `4952ed9` restores legacy backend
alias overrides and formats the installed smoke. Its full non-COM regression
run passed; the complete performance suite was not rerun for that narrow fix.

SciPy comparison and the user-selected optional-dependency decision are in
`issue148-results.md` and ADR-0012. Sparse BFS remains evaluation-only. These
fixtures are synthetic probes rather than every production workbook.
