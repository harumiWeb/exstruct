# ADR-0014: Retain File-First COM Extraction

## 状態

`proposed`

## 背景

Issue #151 evaluates COM-first `standard` / `verbose` extraction after the
file-reading optimizations in #142. Bulk `Value2` avoids individual cell
value calls, but may still incur more COM overhead than the shared file reader.
Excel exposes live/recalculated values rather than saved cached values.
This performance experiment does not propose a new public mode or output contract.

## 決定

- Retain the current file-first public `standard` / `verbose` pipeline.
  The measured candidate does not satisfy the issue's combined performance
  and compatibility conditions. Keep `ComBackend.extract_cells()` as an
  internal prototype used only by the benchmark experiment.
- Retain shared openpyxl table detection and merged/formula metadata in the
  experiment. Do not replace table heuristics with per-cell COM scanning.
- Reconsider adoption only after preserving date interpretation, saved formula
  caches and saved-file values when a workbook has unsaved edits, and after
  reproducible improvement across cold, warm and many-sheet cases.
- Repair owned Excel cleanup independently: opening failure must still quit
  the newly created app. Failed quit may kill only that owned app. Borrowed
  workbooks/apps remain caller-owned. Preserve ADR-0002's fallback contract.

## 影響

- Existing cell semantics, fallback reasons and public entrypoints remain
  stable. Avoiding a default change prevents date serials and recalculated
  formulas from silently changing consumer output.
- File-first continues to pay preprocessing costs before COM. Some potential
  optimization remains unexplored.
- The prototype and benchmark add maintenance cost but preserve evidence for
  revisiting the question. Process-global patches in the experiment are for
  sequential benchmarks, not concurrent production.
- Local Windows results do not cover every Excel build or every possible
  COM-first design. COM failures must be separated from success-path timings.

## 根拠

- Tests: `tests/backends/test_com_bulk_cells.py`,
  `tests/benchmark/test_issue151_com_first.py`,
  `tests/core/test_workbook_utils.py`.
- Code: `src/exstruct/core/backends/com_backend.py`,
  `src/exstruct/core/workbook.py`, `src/exstruct/core/pipeline.py`,
  `benchmark/issue151_com_first.py`, `benchmark/issue151_compatibility.py`.
- Related specs: `dev-docs/specs/excel-extraction.md`,
  `dev-docs/specs/extraction-performance.md`, `benchmark/issue151-results.md`.
- Related decisions: ADR-0002 rich fallback, ADR-0011 workbook reuse,
  ADR-0013 light-only OOXML selection.
- Classification: `recommended`; evaluate and retain the performance strategy
  without changing public semantics. Primary domain: `performance`;
  other domains: `backend`, `extraction`, `compatibility`.

## Supersedes

- None

## Superseded by

- None
