# Issue 148 clustering evaluation

Stage one adds an evaluation-only sparse coordinate BFS, regression tests and
a reproducible comparison runner. **Timings have not been collected**: the
coordinator will run the comparison in a serial measurement slot after other
workers finish their heavy verification. This is not completed SciPy
optionalization. Mandatory dependencies, default backend and fallback policy
are unchanged; removing mandatory SciPy requires the user's explicit decision
after reviewing measurements.

## Correctness observations

`baselines/issue148-2026-10-03-correctness.json` records an untimed local run
against base `5bcb0e47984104414956c9232bfca75b5076ddec` plus this change
(`source_dirty=true`). Interpreter: Python 3.11.13 on Windows; NumPy 2.4.6,
SciPy 1.17.1, pandas 3.0.6 and openpyxl 3.1.5.

| Workload | Grid | Occupied cells | Components (`min_size=4`) | Exact bbox/order and light JSON equality |
| --- | --- | ---: | ---: | --- |
| sparse-small | 40 × 20 | 44 | 2 | passed |
| sparse-large | 2000 × 200 | 74 | 2 | passed |
| dense-style-heavy | 300 × 80 | 24000 | 1 | passed |
| many-disconnected | 300 × 100 | 4000 | 1000 | passed |

The maps and workbook hashes are recorded in JSON. The style-heavy workbook
has borders, fills, fonts and number formats. Workbooks use fixed ZIP/core
timestamps. Each workbook includes an extent value at the bottom-right cell;
normal scan limits and table heuristics still apply during full extraction.
Empty timing dictionaries mean **untimed**, not zero latency.

## Reproduce

From the repository root (all commands use the root locked environment):

```powershell
rtk uv sync --locked
rtk uv run --no-sync python -m benchmark.issue148_clustering --check-only --output benchmark/baselines/issue148-correctness.json
# Only after the coordinator grants an exclusive measurement slot:
rtk uv run --no-sync python -m benchmark.issue148_clustering --repeats 7 --output benchmark/baselines/issue148-timed.json
```

The timed command runs both clustering and end-to-end `light` extraction plus
default JSON serialization for every workload and every implementation. Inputs
are generated outside the timer. Each operation has one untimed warm-up followed
by raw `warm_ms` observations and `median_ms`. Clustering includes grid-to-set
conversion; there is no precomputed sparse index. Full extraction patches the
live Python clustering function temporarily, restoring it and the environment
even on failure. This is an evaluation label (`candidate_sparse`), not a new
supported environment backend name.

Output must match the SciPy reference exactly before results can be accepted.
The runner records revision, dirty state, Python/platform, dependency versions,
environment overrides, grid and fixture hashes, component counts and serialized
output hashes/bytes. Missing optional packages are recorded as null. Missing or
timed-out Git metadata is recorded as null. This runner requires SciPy for its
reference comparison. Results are synthetic workload observations, not universal
performance targets; no RSS, process-cold or Excel COM measurement is claimed.
Repeat measurements on equivalent source, dependency versions and machine load.

## SciPy-free runtime verification

Fresh child-interpreter tests block all SciPy imports before importing ExStruct.
An additional actual clean venv, using locked base dependencies except SciPy,
confirmed identical bordered-workbook `light` JSON for `python`, `auto` and
`numpy`. For the latter two, accelerated failure reaches the existing fallback.

Reproduce with a new temporary directory of your choice:

```powershell
rtk uv venv C:/temp/exstruct-issue148-base --python .venv/Scripts/python.exe
rtk uv export --locked --no-dev --no-emit-workspace --no-emit-package scipy --no-hashes --output-file C:/temp/issue148-base-requirements.txt
rtk uv pip install --python C:/temp/exstruct-issue148-base/Scripts/python.exe --no-deps -r C:/temp/issue148-base-requirements.txt
rtk C:/temp/exstruct-issue148-base/Scripts/python.exe -m benchmark.issue148_scipy_free
```

This runs the source tree and proves runtime compatibility, not that the current
published metadata permits installing ExStruct without SciPy. The current
mandatory SciPy dependency remains intact. There are no Excel COM calls.

## Decision still pending

Compare SciPy, current full-grid Python BFS and candidate set BFS before choosing
whether to remove mandatory SciPy or promote the candidate. The set BFS still
scans the input grid via `nonzero`, and its coordinate/set/queue allocations can
cost more on dense grids. The current SciPy loop scans the full label grid per
component, which is a reason to measure many disconnected regions. Neither fact
alone establishes a performance result.

If the user approves optionalization, proposed metadata is
`fast = ["scipy>=1.16.3"]`, keeping `numpy>=2.3.5` in core because other code
continues to require it. Maintain `auto` / `python` / `numpy` names, current
selection, live monkeypatch overrides and accelerated-exception fallback unless
a separate explicit decision changes them. Coordinate shared docs, ADR/index,
CHANGELOG and dependency metadata changes centrally after approval.
