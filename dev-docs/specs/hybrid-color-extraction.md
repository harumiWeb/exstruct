# Hybrid Rendered Color Extraction Specification

This specification defines how COM-backed `colors_map` extraction combines
saved worksheet fills with Excel-rendered colors while preserving the existing
output contract.

## Scope and compatibility

- Applies to COM-backed color-map extraction used by `standard` / `verbose` on
  `.xlsx` and `.xlsm` workbooks.
- Does not change the public `colors_map` shape, color-key normalization,
  coordinate convention, `include_default_background`, `ignore_colors`, or
  mode selection.
- Excel `DisplayFormat` remains authoritative for cells affected by
  conditional formatting. Python identifies candidate cells; it does not
  evaluate conditional-format formulas or rule outcomes.
- The existing full COM scan remains available as the fallback and internal
  parity oracle.

## Hybrid eligibility

Use the hybrid path only when all of the following hold:

- The COM workbook is an `.xlsx` or `.xlsm` file.
- Its saved file exists and can be loaded and parsed.
- A usable openpyxl extraction session refers to the same normalized absolute
  path as the COM workbook.
- `workbook.api.Saved` is true.

When a caller supplies a session, reuse it only after confirming the path
match. A path-based standalone call owns a temporary session and closes it
after extraction. A session mismatch, unreadable/missing file, unsupported
format, unreadable saved state, or session load failure selects the legacy
full scan. The hybrid path makes no claim that saved formatting represents
unsaved live changes.

## Saved worksheet analysis

Validate each worksheet's raw XML before using it to limit COM calls. The
supported path must understand the worksheet namespace and conditional-format
markup it encounters. Fall back to a full legacy scan for that worksheet if
any of these conditions prevents a complete, trustworthy candidate set:

- The worksheet contains `extLst`, markup-compatibility `AlternateContent`, or
  `pivotTableParts`.
- Conditional-format markup uses an unrecognized namespace or an unrecognized
  rule representation/type.
- Styled row/column definitions, table styles, or merged-cell effects make the
  rendered background of any used cell ambiguous. The initial implementation
  may conservatively fall back when any row/column style, table, or merged
  range is present.
- A custom `Normal` named style (`builtinId=0`) has a non-empty fill, because
  otherwise unstyled cells could inherit a saved fill that is not explicit on
  the cell.
- Worksheet XML cannot be parsed or its relationships/parts cannot be loaded.

The validator is deliberately conservative. It must not ignore unfamiliar
conditional-format constructs and then return a partially optimized map.

## Static fills and candidate cells

For a worksheet that passes validation:

1. Read the saved worksheet fills and restrict all results to the live COM
   `UsedRange`.
2. Resolve ordinary no-fill cells and solid direct-RGB fills without tint from
   saved data, using the existing color normalization and filtering rules.
3. Add cells with other saved static color forms to the COM candidate set.
   This includes theme, indexed, automatic, patterned, gradient, and tinted
   fills. Excel resolves those colors through `DisplayFormat`.
4. Read every recognized conditional-format `sqref` range, intersect it with
   the live COM `UsedRange`, and deduplicate the resulting cells. The count of
   conditional-format candidates is the size of this clipped union.
5. Query `DisplayFormat.Interior.Color` once per unique candidate cell. Merge
   these normalized COM results over the saved static map at the same
   coordinates.

The live COM used range defines the output boundary. Candidate ranges and
saved static cells outside it must not add coordinates to the result.

## Excel evaluation and output semantics

Excel evaluates every candidate's final rendered color. This includes
condition formulas, relative references, calculation state, priority,
overlapping rules, and `stopIfTrue`. ExStruct does not evaluate or simplify
those rules in Python.

Apply existing behavior after combining the saved and COM colors:

- Normalize color keys with the existing RGB conversion.
- Apply `ignore_colors` using the normalized color key.
- Include or omit the default background according to
  `include_default_background`.
- Keep rows 1-based and columns 0-based in each per-sheet map.
- Return the same `SheetColorsMap` and `WorkbookColorsMap` structures.

COM preparation required by the current rendered-color path remains in place.

## Fallback and atomicity

- Workbook-level eligibility, file, session, XML-container, or load failures
  use the legacy full scan for the workbook.
- An unsupported or ambiguous construct isolated to one worksheet uses the
  legacy full scan for that worksheet while eligible worksheets may still use
  the hybrid path.
- If saved-worksheet parsing, candidate construction, or static-fill
  classification fails after producing partial data for a worksheet, discard
  that worksheet's partial map and recompute it through the legacy full scan.
  Never return a partial hybrid map for that worksheet.
- Candidate rendering uses strict `DisplayFormat` reads. If one fails, discard
  the partial hybrid map and retry a full legacy scan for that worksheet. If
  a color read in that retry also fails, preserve legacy per-cell error
  suppression: include the default background when requested or omit the cell
  otherwise. This does not itself trigger a workbook-level fallback.
- Record a fallback reason for each legacy route. Fallback does not change the
  public result shape or expose a new public mode.

## Resource ownership

- Backend extraction reuses the supplied openpyxl session only when its
  resolved path matches the COM workbook path.
- A standalone path-based call owns and closes its session, including on parse,
  COM, or fallback errors.
- The color extractor does not retain sessions across calls or share them
  across concurrent extractions.

## Instrumentation

Record per worksheet:

- Total cells in the live used-range rectangle.
- Unique conditional-format candidate cells after clipping and deduplication.
- Actual `DisplayFormat` cell calls, including calls for unsupported static
  color representations, failed candidate reads, and legacy retry reads. A
  retry can make this count exceed the used-range cell count.
- Static extraction duration, COM rendered-color duration (including a retry),
  and total color extraction duration.
- Fallback reason, or no fallback.

Use deterministic call counts to compare candidate reduction. Wall-clock
measurements are diagnostic and do not impose a universal CI timing threshold.

## Verification fixtures

`tests/core/test_color_hybrid.py` compares normalized maps against the retained
legacy oracle using synthetic workbook data and a fake COM surface. It covers
static/default behavior, blank cells within the COM used range, conditional
candidate clipping/deduplication, conditional-color override, synthetic
theme/indexed/tinted/pattern/automatic/gradient fill classification, XML
fallbacks (`extLst`, unknown rule, foreign namespace, `pivotTableParts`),
custom `Normal` fill, ambiguous row/column/table/merge cases, transient
candidate-read retry, unsaved/session-mismatch/non-`.xlsx` eligibility,
sheet isolation, shared pipeline/session reuse, and standalone session
ownership, including missing/corrupt saved-file fallback. Automatic and
gradient fills are synthetic unit cases, not live-Excel fixtures.

The repeat benchmark and its current measurements are tracked in
[benchmark/issue159-results.md](../../../benchmark/issue159-results.md). Keep
Excel wall-clock results separate from correctness and deterministic
call-count checks.

## Recorded runtime smoke evidence

The final Windows / Excel 16.0 run has 48 measurements for 300-row static and
sparse-CF fixtures: reference and hybrid maps matched in all 24 comparisons
across four `include_default_background` / `ignore_colors` configurations and
three repeats per fixture. Conditional-color validation passed. Calls fell
from 6,000 to 0 for static fills and from 6,000 to 30 for sparse CF.

The separate Windows / Excel 16.0 CF probe used 10-row sparse, dense, and
overlapping-formula fixtures. All configurations matched exactly and expected
conditional colors were visible. In its mixed `A1:B1` probe,
`Range("A1:B1").DisplayFormat.Interior.Color` returned `0`, while individual
cell reads returned `255` and `12611584`; the range property does not provide a
per-cell color matrix.

See the [final benchmark JSON](../../../benchmark/baselines/issue159-colors-final-windows.json),
[CF probe JSON](../../../benchmark/baselines/issue159-colors-cf-probe-windows.json),
and [benchmark results](../../../benchmark/issue159-results.md) for details.
Wall-clock timings vary by environment and are not a universal guarantee.
