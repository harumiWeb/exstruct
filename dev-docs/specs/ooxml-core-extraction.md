# Core OOXML extraction session

Issue: [#149](https://github.com/harumiWeb/exstruct/issues/149)

## Scope and selection

`exstruct.core.ooxml_session.OoxmlExtractionSession` is the read-only core
reader for the normal `.xlsx` / `.xlsm` `light` path. It shares one ZIP with
the OOXML rich backend and does not change public models. The openpyxl pipeline
remains the complete compatibility path and is selected for unsupported
constructs, direct-stage failures, `colors_map` opt-in, or active legacy
pipeline/helper/workbook overrides. See ADR-0013.

```python
from pathlib import Path
from exstruct.core.ooxml_session import OoxmlExtractionSession
from exstruct.core.backends.ooxml_backend import OoxmlRichBackend

path = Path("book.xlsx")
with OoxmlExtractionSession(path) as session:
    sheets = session.sheets()
    cells = session.extract_cells(include_links=True)
    formulas = session.extract_formulas_map()
    merged = session.extract_merged_cells()
    areas = session.extract_print_areas()
    tables = session.extract_explicit_tables()
    names = session.defined_names()
    charts = OoxmlRichBackend(path, session=session).extract_charts(mode="light")
```

## Resource lifetime and parsing

- One lazy ZIP is owned by the context and shared by core and optional rich
  extraction. Independent sessions have independent caches.
- Worksheet XML and shared strings use `defusedxml.ElementTree.iterparse`.
  Completed cell/row and shared-string elements are removed from their parents.
- Accepted plain/rich text segments retain XML document order; phonetic
  annotations do not contribute cell text.
- Sparse values, formulas, merges, hyperlinks and table references are cached
  only after complete worksheet parsing. Subsequent core methods do not reparse
  worksheets. Memory scales with extracted data and the shared-string dictionary;
  this is not a constant-memory iterator API.
- Small workbook, relationship, style and table parts use safe XML trees.
  Existing drawing parsing retains its separate streaming geometry scan.
- Closing releases the ZIP and caches, including after context-body errors.
  Accessing a closed session raises RuntimeError.
- No openpyxl workbook/worksheet model is created. Pure-Python scalar format,
  date and formula-translation helpers preserve the existing value contract.
  The normal supported `light` path does not import openpyxl, pandas, SciPy or
  xlwings; openpyxl remains installed for compatibility and other paths.
- Core malformed XML, missing required parts and invalid scalar data raise
  errors. Missing optional relationship parts return an empty mapping. No
  partially parsed worksheet is cached. Rich extraction retains its existing
  per-sheet best-effort fallback contract.

## Data contract

- `sheets()` returns workbook order, sheet ID, visibility, resolved part path
  and relationship kind, including chart sheets. Core extraction excludes
  non-worksheet sheets. Empty worksheets remain present in mappings.
- `defined_names()` preserves raw text and optional `localSheetId`; named
  formulas are not evaluated.
- `extract_cells()` returns `CellRow`: rows 1-based, columns 0-based numeric-string
  keys, existing missing-token filtering and numeric-string normalization. Sparse
  gaps do not cause rectangular blank cell materialization.
- Merged follower membership is indexed once per parsed sheet over stored
  coordinates. Row/column bisect limits checks to each merge extent without
  expanding empty rectangles. Hyperlinks use bisect over emitted row keys,
  preserving last-link-wins behavior without scanning unrelated rows.
- Shared strings (`s`), plain/rich inline strings (`inlineStr`), formula string
  caches (`str`), numeric (`n`), boolean (`b`) and ISO date (`d`) values are decoded.
  Formatting/phonetic annotations are omitted. Excel errors (`e`) are omitted,
  matching current extraction.
- Numeric date/time/duration styles and 1900/1904 workbook epochs are honored.
  Styles serve scalar conversion, not border heuristics or colors.
- Formula values use stored caches without calculation. `extract_formulas_map()`
  groups normalized `=`-prefixed formulas at `(row, zero-based column)` coordinates.
  Shared followers translate relative references from the shared anchor.
  Array/dynamic formulas retain explicit anchor text without inventing follower
  formulas. Unresolved shared followers are omitted. Shared translation of
  whole-row/column ranges, unquoted sheet-qualified references, non-ASCII names,
  or references crossing worksheet bounds raises `UnsupportedOoxmlError`;
  the public light pipeline restarts compatibility extraction instead of
  emitting a partially translated formula map. Quoted sheet names and string
  literals are preserved by the ordinary A1 translator.
- Date/duration style classification is cached once per session. Elapsed-time
  formats are case-insensitive, including `[H]:MM:SS`, and preserve both row
  display text and raw merged-anchor values.
- External hyperlink URLs are preserved exactly. Internal `location` links are
  omitted, matching current output. Links on filtered values survive only when
  their row has another emitted value. Range links use zero-based column keys.
- Merged ranges contain cached raw anchor text or a space for an empty anchor;
  covered follower values are excluded from rows.
- Print areas come from `_xlnm.Print_Area`, respecting local scope and quoted
  names including commas/apostrophes. Unsupported non-finite ranges are logged
  and skipped.
- `extract_explicit_tables()` returns referenced table names, display names,
  ranges and column names, without the border heuristics of `table_candidates`.
- `detect_tables(name, mode="light")` also applies the existing border-based
  heuristics to produce `table_candidates` on the direct OOXML path. It passes
  `cluster_backend="python"` per call so the normal path avoids importing SciPy
  without reading or changing `EXSTRUCT_BORDER_CLUSTER_BACKEND`.
- The table worksheet view indexes merged rectangles by row-boundary bands
  and column-boundary segments. Each cell lookup uses two binary searches;
  construction never expands all rows or cells covered by a merge. Overlaps
  retain the first merge in XML order, and anchor values and inherited outer
  borders keep their existing semantics. Index storage depends on boundary
  segments (potentially quadratic in merge count for adversarial overlaps),
  rather than merged-cell area or the table scan extent.

## Pipeline selection and recovery

- The normal `light` pipeline uses one session archive for core worksheet data,
  formulas, hyperlinks, merged cells, print areas, table candidates, shapes and
  charts through final model construction.
- Any uncaught error or unsupported construct in any direct OOXML stage
  discards all partial OOXML results. The session closes in a `finally` path,
  then extraction restarts from the beginning through the complete openpyxl
  compatibility pipeline.
- The warning and extraction state use
  `FallbackReason.OOXML_COMPATIBILITY` (`ooxml_compatibility`). Known
  unsupported options and legacy overrides select compatibility before the
  ZIP is opened where possible.
- Existing per-sheet drawing resilience remains: a drawing failure isolated to
  one sheet may omit that sheet's affected drawing artifacts while preserving
  core data and healthy-sheet drawings.
- Other modes and `.xls` selection/behavior remain unchanged.

## Relationships

`ooxml_package.read_relationships(archive, source_path)` resolves IDs to type,
target and external status. Internal relative paths resolve against the source
directory; absolute targets resolve from the ZIP root. External targets are
preserved and never fetched or read as ZIP parts. These primitives serve
hyperlinks, tables, worksheets, drawings and charts. Legacy drawing helpers
remain available as wrappers.
When drawing extraction shares a core session, workbook enumeration uses the
same resolved officeDocument part as core extraction, including a relocated
workbook. Standalone path-based drawing reads retain their conventional default.

## Review follow-up validation

PR #154 adds regressions for accepted interleaved text segments in shared and
inline strings, a relocated `custom/book.xml` read before any cell extraction,
and sheet-wide sparse merged ranges with overlapping hyperlink overrides.
The relocated fixture includes both a chart and a shape and verifies one ZIP
open per session. Merge coverage indexes only stored coordinates and is reused
across linked/unlinked row extraction.

A local Windows synthetic comparison of the original `807e206` row builder
and the revised builder used 40,000 stored cells (20,000 rows, two columns),
200 single-row merges and 200 links. Exact output equality was asserted:
original 0.684129 s, revised first call 0.124096 s, revised cached call 0.099105 s.
These are row-construction timings from one run, not workbook extraction or
memory measurements.

Follow-up validation: 27 focused tests passed; 1,040 full tests passed with
12 external-runtime tests deselected and 83.57% coverage. Existing eight-workbook
parity, Ruff, strict mypy and diff checks passed. These fixes restore extraction
semantics and improve internal indexing; they introduce no backend-selection,
public mode or fallback policy change and require no new ADR.

## Decision history

Issue #149 added reusable OOXML parsing primitives without changing backend
selection; its original no-ADR assessment applied to that capability-only
scope. Issue #150 changes the default backend and whole-pipeline fallback
contract. ADR-0013 records that policy and applies to this specification.

## Validation

Parity fixtures compare with `OpenpyxlBackend` for both suffixes, link flags,
cells, formulas, merges and print areas. Additional tests cover shared/rich
strings, cached and shared formulas, relationships, tables, archive/read reuse,
sparse sheets, exception cleanup and XML entity rejection. Existing drawing
and pipeline tests retain rich-artifact guarantees.

```powershell
rtk uv run --extra all pytest tests/core/test_ooxml_extraction_session.py tests/core/test_ooxml_package.py tests/core/test_ooxml_drawing.py -q
rtk uv run --extra all pytest -m "not com and not render and not libreoffice" -q --cov=exstruct --cov-report=term --cov-fail-under=80
rtk uv run ruff check . --no-fix
rtk uv run --extra all mypy src/exstruct --strict
```

Local Windows/Python 3.11 verification: 23 focused tests passed; 1,036 tests
passed with 12 external-runtime tests deselected, total coverage 83.51% (80%
gate passed). Ruff and strict mypy passed. Excel COM/render and real LibreOffice
runtime verification are not part of this internal-parser validation.
An additional comparison of all seven tracked `sample/**/*.xlsx` workbooks and
`tests/assets/multiple_print_ranges_4sheets.xlsx` matched cells with links,
formula maps, print areas and merged-range sets against `OpenpyxlBackend`.
