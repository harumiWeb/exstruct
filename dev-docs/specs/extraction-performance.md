# Extraction performance measurements

Issue [#143](https://github.com/harumiWeb/exstruct/issues/143), part of #142,
establishes the baseline before production extraction optimizations. The
repository-local `benchmark.performance` module is separate from `bench`'s
correctness/LLM benchmark and needs only the root ExStruct environment.

## Execution contract

Internal runner entry points (not public ExStruct APIs):

- `generate_fixtures(directory: Path) -> list[Path]`
- `run_benchmark(path: Path, mode: Literal["light", "standard"], repeats: int,
  timeout: float, profile_all_features: bool = False) -> dict[str, Any]`
- `worker(path: Path, mode: Literal["light", "standard"], repeats: int,
  profile_all_features: bool = False) -> dict[str, Any]`
- `fresh_process(arguments: list[str], timeout: float) -> tuple[float, str]`
- `source_metadata(timeout: float) -> dict[str, Any]`
- `main(argv: list[str] | None = None) -> int`

The CLI requires at least one repeat, a finite positive timeout and an existing
input; it refuses output paths that resolve to the input itself. The result
dictionary follows JSON schema version 1 described below.

Run from the repository root with the locked Python environment. Each command
benchmarks one input and one mode (`light` or `standard`), preserving the public
`exstruct.extract` defaults. `standard` attempts COM normally; the runner never
sets `SKIP_COM_TESTS`. Environment overrides and the profiled pipeline's COM
attempt/success/fallback reason are recorded. Compare COM-success measurements
only with equivalent COM-success measurements; fallback costs are different.

The supervisor launches fresh processes with the same `sys.executable` for:

1. Python startup (`python -c pass`).
2. Bare `import exstruct`.
3. Extraction CLI help (`python -m exstruct.cli.main --help`). This measures parser
   startup, not runtime COM availability validation or a CLI extraction.
4. A measurement worker: bare import, first extraction, serialization, then
   `--repeats` extractions/serializations in the same process, and a separate profile.
5. One cold extraction + JSON serialization, with no preceding extraction in its
   process. This process imports the bare package before its internal timer;
   its parent wall time also includes Python startup, bare import and teardown.

Fresh-process latency is not filesystem/OS cache coldness. Startup probes are
individual raw observations, not statistical estimates; repeat commands for
startup distributions. No timing subtraction is used to estimate import cost.
Child errors and timeouts return a failing command; timeout is per child, not a
shared end-to-end budget. Both Git metadata subprocesses also use the configured
timeout; a timeout or unavailable Git produces null revision/dirty metadata,
without discarding completed measurements. Successful child stderr is forwarded
to supervisor stderr after its wall-clock measurement ends, so fallback warnings
remain visible without contaminating JSON stdout or inflating child latency.
The worker's entire process duration includes repeats
and profiling and is **not** the cold-extraction metric.

## JSON schema version 1

Every result records UTC timestamp, source revision/dirty state, interpreter,
platform, relevant dependency versions, environment overrides, input name,
size and SHA256. All durations are milliseconds using `perf_counter`.

| Field | Meaning |
| --- | --- |
| `startup.python_process_ms` | Parent wall time for a fresh no-op interpreter |
| `startup.package_import_process_ms` | Parent wall time including bare import |
| `startup.cli_help_process_ms` | Parent wall time including CLI help |
| `startup.cold_extraction_process_ms` | Parent wall time for one cold extraction + serialization |
| `startup.cold_extraction_in_process_ms` | Inside that process, extraction + serialization after bare import |
| `startup.measurement_worker_process_ms` | Whole measurement worker including repeats/profile |
| `import_ms` | Bare package import inside the worker |
| `first` | First extraction, serialization and their total in the worker |
| `repeated` / `repeated_median` | Raw warm observations / per-metric medians |
| `peak_rss_bytes` | Process lifetime high-water RSS sampled after the uninstrumented runs; null if unavailable |
| `modules_after_import` / `modules_after_extraction` | Imported roots from the documented heavyweight module list |
| `output_bytes` | Default compact serialized JSON size in UTF-8 |
| `profile` | Separate warm instrumented run: durations, calls/opens and actual pipeline state |

Peak RSS uses Windows `PeakWorkingSetSize` or POSIX `ru_maxrss` (platform units
normalized to bytes). It includes imports and first/repeated extraction, is not
a memory delta, and excludes the later instrumentation imports and Excel/other
child-process memory. No RSS sampler/psutil dependency is introduced.
The observed module roots are pandas, numpy, scipy, xlwings, openpyxl, PIL and
pypdfium2; this is not an exhaustive import trace.

## Stage attribution

Temporary wrappers instrument backend extraction methods, COM pipeline steps,
table detection and `build_workbook_data`. Openpyxl `load_workbook` and all loaded
aliases (including pandas' reader) are wrapped; `ZipFile.__init__` calls are
counted separately. Wrappers are restored even when extraction raises.
Instrumentation is installed only after normal first/repeated measurements,
and an uninstrumented equivalent output must match the profiled serialized JSON.

Stages cover cells, print areas, formulas, colors, merged cells, table detection,
rich shapes/charts, final workbook model construction and workbook parsing.
Profile serialization and total are recorded separately. Model construction
means final `SheetData`/`WorkbookData` construction, not `CellRow` construction
inside cell extraction. Stage durations are **inclusive**: parsing inside cell
or table extraction overlaps their durations, so summing stage timings is invalid.
COM opening/closing and remaining orchestration occur in extraction total but
are not separate stage metrics. Archive opens are constructor attempts and
openpyxl workbook opens are loader calls; neither includes COM workbook opens.

`calls=0, ms=0` means a stage did not run. Formula/color stages are normally
disabled for these modes. `--profile-all-features` enables formulas, colors and
merged cells in the separate profile only; latency measurements still use mode
defaults. That profile uses `extract_workbook` with explicit include flags and
checks against its own uninstrumented equivalent. Compare profiles with identical
`profile.all_features` settings. Profile times include wrapper overhead and
are for attribution, not substitutes for the uninstrumented latency observations.

For profile measurements, StageRecorder.install enables direct OOXML
instrumentation only for the normal light profile, with direct_ooxml true when
mode is light and all-features profiling is disabled. This instruments the
backend selected by the default request; it does not change backend selection.
For a direct OOXML profile, workbook_parsing reports zero calls because that
stage counts openpyxl load_workbook calls. OOXML XML parsing time is included in
the cells stage duration.

## Fixtures and comparison

`benchmark.performance_fixtures` generates six fixed synthetic categories; see
`benchmark/README.md` for dimensions and commands. ZIP/workbook timestamps are
fixed, and identical generator/dependency versions produce identical hashes.
Existing fixture files are never overwritten.
Generation takes place in a temporary staging directory before any fixture is
published. Exceptions/interruption during generation leave the final fixture
names absent, allowing retry in the same directory. Publication exclusively
creates files; if publication raises, only files created by this call are removed.
Unrelated files and concurrent writers' pre-existing files are preserved. This
is exception recovery, not crash-atomic publication of all six files.
Record and compare input hashes,
locked dependencies, machine/load, source revision, mode and runtime state.
Synthetic fixtures are reproducible workload probes, not claims of production
representativeness; existing sample workbooks can be passed with `--input` too.

Initial observations are stored in `benchmark/baselines/2026-10-03-windows.json`.
The extraction implementation is revision `563da66` (v0.8.2); `dirty=true` records
the added benchmark/doc/test files. No production extraction code is changed.
Results are local observations, not universal performance targets. CI should
test measurement contracts without strict shared-runner wall-clock thresholds.
The committed baseline's `environment.executable` is replaced with `<local-python>`
to omit the identifying local path; local runner output still records its actual
interpreter path. Timing samples and other reproducibility metadata are retained.

The integrated #145-#148 comparison and exact output hashes are recorded in
`benchmark/issues145-148-results.md` and its linked raw baselines. It separates
light improvements from variable Excel COM observations and retains supplemental
paired rechecks. `benchmark/issue148-results.md` records the separate accelerator
comparison and ADR-0012 dependency decision; the candidate sparse BFS remains
evaluation-only.

## Issue #150 light-backend comparison

benchmark/issue150_light.py compares the preserved openpyxl light pipeline with
the direct OOXML pipeline in separate worker processes. It records import and
extraction timings, cold in-process duration, repeated-extraction median, peak
RSS, workbook and ZIP open counts, imported heavy modules, fallback reason, and
serialized output hash. The runner stops if the two output hashes differ.

The final Windows result file is
benchmark/baselines/issue150-2026-10-03-windows.json, recorded on Windows
10.0.22631 with Python 3.11.13. The seven inputs all produced matching
serialized-output hashes. OOXML used one ZIP archive, opened zero openpyxl
workbooks, and loaded none of openpyxl, pandas, SciPy or xlwings. The previous
pipeline opened two ZIP archives, one openpyxl workbook, and loaded openpyxl and
SciPy. See the [Issue #150 benchmark report](../../benchmark/issue150-results.md)
for the run details.

| Input | openpyxl cold / warm median (ms) | OOXML cold / warm median (ms) | openpyxl / OOXML peak RSS (MiB) |
| --- | ---: | ---: | ---: |
| small.xlsx | 619.9 / 7.6 | 267.9 / 3.9 | 64.4 / 39.4 |
| large.xlsx | 955.1 / 481.1 | 645.5 / 407.2 | 109.3 / 71.8 |
| many-sheet.xlsx | 633.2 / 82.6 | 297.6 / 63.4 | 74.5 / 42.6 |
| sparse.xlsx | 644.9 / 39.3 | 256.9 / 15.9 | 74.5 / 39.5 |
| style-heavy.xlsx | 715.4 / 117.5 | 318.8 / 82.6 | 73.4 / 46.0 |
| table-heavy.xlsx | 616.9 / 57.4 | 281.0 / 50.1 | 75.3 / 44.1 |
| sample-shape-connector.xlsx | 588.7 / 13.0 | 244.9 / 5.5 | 64.5 / 39.6 |

OOXML had a lower cold in-process duration, warm median and peak RSS on all
seven inputs in this run. The cold in-process metric includes module imports and
the first extraction, but excludes worker process startup and teardown measured
separately. These are observations from one Windows machine and one recorded
input set, not universal runtime guarantees.
