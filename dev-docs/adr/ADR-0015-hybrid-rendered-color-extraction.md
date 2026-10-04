# ADR-0015: Hybrid Rendered Color Extraction

## 状態

`proposed`

## 背景

Issue #159 addresses the cost of reading Excel's rendered background color
through `DisplayFormat` for every cell in a worksheet's used range. Saved cell
fills and conditional-formatting ranges can narrow that work, but saved styles
alone do not represent every rendered fill, and Python must not reproduce
Excel's conditional-format evaluator.

The optimization is limited to `.xlsx` / `.xlsm` workbooks whose saved file can
be tied to the open COM workbook and the extraction's openpyxl session. It does
not assume that saved formatting describes unsaved live workbook changes. The
existing public `colors_map` behavior and the Excel COM rendering authority
remain in place.

## 決定

- Adopt a hybrid strategy for COM-backed `colors_map` extraction when the
  workbook is `.xlsx` / `.xlsm`, the saved file exists, the COM workbook path
  matches the extraction session path, and `workbook.api.Saved` is true.
- Recheck `workbook.api.Saved` after preparing each sheet for rendered-color
  reads. Sheet activation or calculation can run VBA events that change live
  fills after the saved snapshot was opened. If `Saved` is false or cannot be
  read, latch the legacy full scan for that sheet and every remaining sheet;
  do not resume hybrid extraction later in the workbook.
- Read ordinary static fills from the saved worksheet representation. Use
  saved worksheet XML to identify conditional-format ranges, clip those ranges
  to the live COM `UsedRange`, and deduplicate candidate cells. Also send
  static cells with unsupported saved color representations to COM as candidates.
- Traverse only stored openpyxl cells when classifying saved static fills.
  When default backgrounds are requested, emit entries for absent coordinates
  within `UsedRange` without creating openpyxl cells for them.
- Continue to ask Excel for the final color of each candidate cell through
  `DisplayFormat`. Excel remains responsible for all conditional-format rule
  evaluation, including formula semantics, priority, overlap, and
  `stopIfTrue`; Python only identifies candidate coordinates.
- Fall back to the existing full `DisplayFormat` scan for a worksheet whenever
  its XML or formatting constructs make candidate coverage uncertain. This
  includes `extLst`, `AlternateContent`, `pivotTableParts`, unsupported
  conditional-format namespaces/rules, styled row/column definitions,
  table styles, merged-cell effects, and a custom `Normal` named style
  (`builtinId=0`) with a non-empty fill. Missing or unreadable files,
  unsupported formats, an ineligible workbook state, a session path mismatch,
  and raw XML or workbook-load failures also use the legacy scan. Discard
  partial saved-analysis results before scanning that worksheet. If a strict
  candidate `DisplayFormat` read fails, discard that worksheet's partial hybrid
  map and retry its full legacy scan. If a legacy color read also fails, retain
  existing per-cell error suppression: include the default background when
  requested or omit the cell otherwise. Count failed candidate reads and the
  retry scan in display-call instrumentation.
- Reuse a supplied extraction session only for the matching workbook path.
  Standalone path-based calls own and close their session. Retain the current
  COM preparation and full-scan implementation as the correctness fallback
  and internal parity oracle.
- Preserve the existing output model and semantics: color normalization,
  coordinates, per-sheet and workbook maps, `include_default_background`, and
  `ignore_colors`. Do not change public APIs, modes, or serialized schemas.
- Record used-range cell count, conditional-format candidate count, actual
  `DisplayFormat` call count, static/COM/total durations, and fallback reason
  for diagnosis and performance comparison.

## 影響

- Worksheets with ordinary saved fills and no conditional formatting can avoid
  per-cell `DisplayFormat` calls. Sparse conditional-format ranges can reduce
  those calls to the union of clipped candidate cells plus cells whose saved
  static color representation needs Excel resolution.
- Dense conditional formatting and conservative fallbacks retain the legacy
  scan cost. The strategy does not promise a wall-clock speedup for every
  workbook or Excel installation.
- Raw XML validation and style classification add implementation complexity.
  Conservative per-sheet fallback limits the risk of returning incomplete or
  incorrectly rendered maps as the OOXML surface evolves.
- Excel remains the source of truth for rendered colors, so formula and rule
  behavior stays aligned with the installed Excel version. Unsaved live
  formatting is handled by the legacy path rather than inferred from the saved
  file.
- Instrumentation makes candidate reduction and fallback behavior observable;
  deterministic call counts can be compared independently of machine-specific
  timing.

## 根拠

- Tests: `tests/backends/test_colors_map.py` covers default-background and
  ignored-color output; `tests/core/test_cells_color_normalization.py` covers
  color-key normalization; `tests/core/test_openpyxl_extraction_session.py`
  covers shared-session color extraction, standalone/session parity, and
  resource lifecycle on the compatibility path. `tests/core/test_color_hybrid.py`
  uses synthetic workbook data and a fake COM surface to cover static-fill
  parity, candidate clipping/deduplication, conditional-color override,
  unsupported XML/style fallbacks (including `pivotTableParts` and custom
  `Normal` fill), strict candidate-read retry, unsaved/session-mismatch/
  non-`.xlsx` eligibility, post-preparation `Saved` invalidation (false or
  unreadable) for the current and remaining sheets, sparse traversal without
  materializing missing openpyxl cells, missing/corrupt saved files, sheet
  isolation, shared pipeline/session behavior, and session ownership. Automatic and gradient
  fills are synthetic unit cases, not live-Excel fixtures.
- Code: `src/exstruct/core/color_hybrid.py` implements saved-fill and candidate
  analysis; `src/exstruct/core/cells.py` retains the COM color entrypoint and
  legacy per-cell `DisplayFormat` scan;
  `src/exstruct/core/backends/com_backend.py` passes the optional extraction
  session; `src/exstruct/core/openpyxl_session.py` owns the reusable workbook
  session.
- Related specs: `dev-docs/specs/excel-extraction.md`,
  `dev-docs/specs/hybrid-color-extraction.md`, and
  `dev-docs/specs/extraction-performance.md`.
- Related decisions: ADR-0002 defines fallback policy; ADR-0011 defines
  extraction-scoped workbook reuse; ADR-0013 defines the light-mode backend
  boundary; ADR-0014 records the separate file-first cell-extraction decision.
- Benchmark: The final Windows / Excel 16.0 run contains 48 measurements for
  300-row static and sparse-CF fixtures: all 24 reference/hybrid map comparisons
  passed across four configurations and three repeats per fixture. Conditional
  color validation passed. `DisplayFormat` calls fell from 6,000 to 0 for
  static fills and from 6,000 to 30 for sparse CF. See the [final benchmark
  JSON](../../benchmark/baselines/issue159-colors-final-windows.json).
  The separate 10-row sparse, dense, and overlapping-formula probe also passed
  all configurations and verified visible CF colors; its mixed `A1:B1` range
  returned `0` while per-cell reads returned `255` and `12611584`, confirming
  the range property is not a per-cell color matrix. See the [CF probe
  JSON](../../benchmark/baselines/issue159-colors-cf-probe-windows.json).
  Measurement details are in [benchmark results](../../benchmark/issue159-results.md).
  Wall-clock results are environment-specific and do not imply a universal
  timing guarantee.
- Classification: `recommended`; record a performance-strategy change while
  retaining public semantics. Primary domain: `performance`; other domains:
  `backend`, `extraction`, `compatibility`, `fallback`.

## Supersedes

- None

## Superseded by

- None
