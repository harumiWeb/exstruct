# ADR-0012: Optional SciPy Border Clustering

## 状態

`proposed`

## 背景

Issue #148 evaluates whether border clustering justifies a mandatory SciPy
dependency. NumPy is used beyond clustering and remains required. The existing
Python BFS already provides a fallback; removing that implementation or changing
backend names would expand compatibility scope unnecessarily.

The coordinator compared SciPy, existing Python BFS and an evaluation-only sparse
set BFS on four deterministic workloads, with one warmup and seven timed repeats.
All bounding boxes, their order and full light extraction JSON matched exactly.
After reviewing these measurements, the user chose optional SciPy with the
existing Python fallback on 2026-10-03.

## 決定

- Remove SciPy from mandatory runtime dependencies; offer `fast = ["scipy>=1.16.3"]`
  and include SciPy in `all`. Keep NumPy in core.
- Retain the existing `EXSTRUCT_BORDER_CLUSTER_BACKEND` values and selection:
  `auto` attempts SciPy, `python` uses the existing BFS, and the legacy `numpy`
  name attempts SciPy-backed labeling. Import or accelerated execution failure
  still falls back to Python. Do not add a new alias or change defaults.
- Keep sparse-set BFS as an internal evaluation candidate. Do not promote it to
  production selection based on these results.

## 影響

- Base installation becomes smaller and works without SciPy. Users who want the
  previous accelerated installation can choose `exstruct[fast]` or `[all]`.
- Dense clustering remains much faster with SciPy: the measured clustering
  medians were 0.217 ms (SciPy), 19.959 ms (Python), and 26.997 ms (candidate).
  Full extraction plus JSON was 529 ms, 590 ms, and 524 ms respectively;
  worksheet loading dominates much of this workload, and these sequential
  synthetic timings must not be treated as universal speed guarantees.
- Many disconnected components favored Python: clustering medians were 53.113 ms
  (SciPy) and 4.877 ms (Python), with extraction totals of 374 ms and 299 ms.
  Large sparse candidate clustering improved independently, but whole extraction
  did not establish a consistent candidate benefit.
- Keeping the legacy `numpy` name preserves configuration compatibility even
  though it describes an implementation that also uses SciPy. Public docs must
  state this clearly. Base installation still includes NumPy, xlrd and xlwings;
  optional SciPy does not imply that every backend dependency is optional.

## 根拠

- Tests: `tests/core/test_border_clustering_issue148.py` covers exact ordering,
  blocked SciPy imports, accelerated failures and live overrides.
- Code: `src/exstruct/core/cells.py`, `pyproject.toml`, `uv.lock`,
  `benchmark/issue148_clustering.py`, and `benchmark/issue148_scipy_free.py`.
- Related specs: `docs/api.md`, `dev-docs/specs/excel-extraction.md`,
  `dev-docs/specs/extraction-performance.md`, and `benchmark/issue148-results.md`.
- Measurements: `benchmark/baselines/issue148-2026-10-03-windows.json` records
  raw timings, input hashes, output hashes, dependency versions and source state.
- Related decisions: ADR-0010 retains the light OOXML baseline; ADR-0011 treats
  direct cell reading and extraction-scoped resources separately.

## Supersedes

- None

## Superseded by

- None
