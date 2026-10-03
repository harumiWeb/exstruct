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
shared end-to-end budget. The worker's entire process duration includes repeats
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

## Fixtures and comparison

`benchmark.performance_fixtures` generates six fixed synthetic categories; see
`benchmark/README.md` for dimensions and commands. ZIP/workbook timestamps are
fixed, and identical generator/dependency versions produce identical hashes.
Existing fixture files are never overwritten. Record and compare input hashes,
locked dependencies, machine/load, source revision, mode and runtime state.
Synthetic fixtures are reproducible workload probes, not claims of production
representativeness; existing sample workbooks can be passed with `--input` too.

Initial observations are stored in `benchmark/baselines/2026-10-03-windows.json`.
The extraction implementation is revision `563da66` (v0.8.2); `dirty=true` records
the added benchmark/doc/test files. No production extraction code is changed.
Results are local observations, not universal performance targets. CI should
test measurement contracts without strict shared-runner wall-clock thresholds.
