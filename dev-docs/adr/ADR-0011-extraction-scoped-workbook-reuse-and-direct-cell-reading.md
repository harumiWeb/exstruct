# ADR-0011: Extraction-Scoped Workbook Reuse and Direct Cell Reading

## 状態

`proposed`

## 背景

Issues #145-#147, under #142, address unnecessary DataFrame conversion,
repeated workbook parsing and backend imports on extraction startup. ExStruct
returns typed workbook data rather than DataFrames, so retaining pandas in the
production cell reader adds a substantial dependency without providing a
public abstraction. Independent path-based helpers also parse the same OOXML
workbook repeatedly, including from nested table detection helpers.

ADR-0010 already defines light as a pure-Python rich OOXML baseline. This
decision improves its implementation while retaining that mode boundary,
serialization, coordinate conventions and fallback behavior. It does not
introduce a new public extraction backend or change COM-first/pre-COM ordering.

## 決定

- Read `.xlsx` and `.xlsm` cells directly through openpyxl worksheet iteration;
  build `CellRow` without a pandas DataFrame. Preserve the previous reader's
  observable normalization, empty-cell filtering and cached-formula behavior.
- Retain `.xls` cell reading through a lazily imported direct xlrd reader. This
  preserves the public input format without retaining pandas or requiring Excel
  COM for cell reading. It does not add OOXML artifacts to BIFF workbooks.
- Scope openpyxl workbook ownership to one extraction invocation. Compatible
  stages reuse a regular `data_only=True` workbook. Open a separate
  `data_only=False` workbook only when openpyxl formula extraction is requested.
- Close owned workbooks when the extraction finishes or raises, including
  fallback paths. Retain path-based standalone entrypoints with local ownership.
  Do not cache workbooks globally or across extraction invocations.
- Keep common backend contracts independent of concrete COM implementations;
  import the selected backend only when execution needs it. Light OOXML rich
  extraction must retain shapes and charts while avoiding COM/render imports.
- Preserve compatibility call sites used for overrides and monkeypatches. Import
  isolation must not change which overridden function an entrypoint executes.
- Evaluate clustering accelerators separately in #148. ADR-0012 records the
  user-selected optional SciPy policy after measurements; this ADR does not
  change clustering selection defaults.

## 影響

- Direct cell construction removes production DataFrame allocation and the
  pandas runtime requirement; pandas may remain in development/benchmark tools.
- Per-invocation ownership reduces parsing cost without stale data across calls
  or cross-extraction resource sharing. Formula extraction still needs a second
  workbook because its load semantics differ from cached cell values.
- A regular openpyxl workbook retains more resident objects than a streaming
  read-only reader. Sharing it is preferred because borders, hyperlinks and
  merged-cell extraction require those objects; memory and latency must still be
  measured independently.
- Direct readers must explicitly preserve pandas-era behavior for values such
  as NA tokens, booleans, errors and dates. Characterization regressions are
  needed to prevent dependency removal from changing output semantics.
- xlrd becomes an explicit core dependency for the existing `.xls` input
  capability, though it is imported only on that path. NumPy remains required
  for current table detection; this change does not claim a dependency-free
  light implementation.
- Lazy compatibility wrappers add indirection and require subprocess tests;
  bare package import checks alone cannot prove real extraction isolation.

## 根拠

- Tests: `tests/core/test_cells_reader_parity.py` (frozen previous-reader output,
  actual BIFF and date-mode comparisons),
  `tests/core/test_openpyxl_extraction_session.py` (loader counts, cleanup,
  concurrency, overrides and BIFF routing),
  `tests/core/test_backend_import_isolation.py` (real subprocess extraction),
  `tests/cli/test_cli_lazy_imports.py`, and `tests/benchmark/test_performance.py`.
- Code: `src/exstruct/core/cells.py`, `src/exstruct/core/workbook.py`,
  `src/exstruct/core/pipeline.py`, `src/exstruct/core/openpyxl_session.py`,
  `src/exstruct/core/cell_types.py`, and `src/exstruct/core/backends/`.
- Related specs: `dev-docs/specs/excel-extraction.md`,
  `dev-docs/specs/extraction-performance.md`, and `docs/api.md`.
- Related decisions: ADR-0002 fallback policy, ADR-0003 serialization policy,
  and ADR-0010 light-mode responsibility boundary.

## Supersedes

- None

## Superseded by

- None
