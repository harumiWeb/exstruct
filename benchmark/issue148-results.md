# Issue 148 clustering evaluation

SciPy is now optional through `exstruct[fast]` (also included in `[all]`).
The user selected this policy after reviewing the serial Windows measurements
below on 2026-10-03. NumPy remains mandatory. Existing `auto` / `python` / `numpy`
selection and accelerated-exception fallback are unchanged. Sparse-set BFS is an
evaluation-only candidate and is not selected in production.

## Measurements and correctness

`baselines/issue148-2026-10-03-windows.json` records one warmup and seven timed
repeats on Windows, Python 3.11.13, NumPy 2.4.6, SciPy 1.17.1 and openpyxl 3.1.5.
Source was integrated revision `49f2d66` with documentation/result changes pending.
Timing excludes fixture generation; clustering includes grid-to-set conversion.
Every backend produced identical bounding boxes, ordering and full light JSON.

| Workload | Grid / occupied cells | Clustering median ms: SciPy / Python / candidate | Light extraction + JSON median ms: SciPy / Python / candidate |
| --- | --- | --- | --- |
| sparse-small | 40 × 20 / 44 | 0.031 / 0.077 / 0.031 | 8.8 / 9.6 / 8.5 |
| sparse-large | 2000 × 200 / 74 | 2.511 / 18.658 / 0.618 | 1150 / 1076 / 1099 |
| dense-style-heavy | 300 × 80 / 24000 | 0.217 / 19.959 / 26.997 | 529 / 590 / 524 |
| many-disconnected | 300 × 100 / 4000 | 53.113 / 4.877 / 3.049 | 374 / 299 / 322 |

Dense labeling strongly favors SciPy, while many separated components favor
Python. The candidate improves some isolated clustering cases but does not
establish consistent whole-extraction gains; it is not promoted. These synthetic
sequential observations are not universal performance targets. No RSS,
process-cold or Excel COM measurement is claimed by this comparison. The general
before/after runner measures those separately.

Raw samples, dependency versions, source state, environment, grid/fixture hashes,
component counts and output hashes/bytes are in JSON. The earlier untimed
`baselines/issue148-2026-10-03-correctness.json` is retained as characterization
evidence; its empty timing dictionaries mean untimed rather than zero latency.
Inputs have fixed ZIP/core timestamps. Style-heavy workbooks include borders,
fills, fonts and number formats. Extent markers and normal scan limits apply.
These timings precede the integrated sparse worksheet reader refinement; they
support dependency selection, not final before/after extraction speed claims.

## Reproduce

```powershell
rtk proxy uv sync --locked --all-extras
rtk proxy uv run --no-sync python -m benchmark.issue148_clustering --check-only --output output/issue148-correctness.json
# Run in an exclusive measurement slot, without other tests or Excel jobs:
rtk proxy uv run --no-sync python -m benchmark.issue148_clustering --repeats 7 --output output/issue148-timed.json
```

The reference comparison requires SciPy. Missing optional packages and unavailable
Git metadata are recorded as null. Full extraction temporarily patches the live
Python clustering function and restores overrides/environment even on failure.
`candidate_sparse` is an evaluation label, not a supported environment value.

## Base installed-package verification

The base-install CI jobs install the package without development or acceleration
extras, then run `python -m benchmark.issue148_scipy_free --installed`. This smoke
requires both pandas and SciPy to be absent, refuses source-checkout imports and
compares actual bordered-workbook light JSON for `python`, `auto` and `numpy`.
For `auto` and `numpy`, the existing accelerated failure reaches Python fallback.
Fresh subprocess tests also block SciPy before importing ExStruct and cover
accelerated execution exceptions and live helper overrides.

For a local wheel check:

```powershell
rtk proxy uv build --wheel --out-dir output/base-wheel
rtk proxy uv venv output/base-install --python .venv/Scripts/python.exe
rtk proxy uv pip install --python output/base-install/Scripts/python.exe output/base-wheel/exstruct-0.8.2-py3-none-any.whl
rtk proxy output/base-install/Scripts/python.exe -m benchmark.issue148_scipy_free --installed
```

The environment is separate from the normal development environment, which
includes pandas for characterization and SciPy when installing `fast`/`all`.
See ADR-0012 for the dependency-policy rationale and `docs/api.md` for the
legacy `numpy` name's SciPy-backed meaning.
