# Core OOXML extraction session

Issue: [#149](https://github.com/harumiWeb/exstruct/issues/149)

## Scope and selection

`exstruct.core.ooxml_session.OoxmlExtractionSession` is an internal, read-only
alternative for `.xlsx` / `.xlsm` core extraction. It does not replace the
openpyxl pipeline or change public models. Default-path migration is separate.

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
- Sparse values, formulas, merges, hyperlinks and table references are cached
  only after complete worksheet parsing. Subsequent core methods do not reparse
  worksheets. Memory scales with extracted data and the shared-string dictionary;
  this is not a constant-memory iterator API.
- Small workbook, relationship, style and table parts use safe XML trees.
  Existing drawing parsing retains its separate streaming geometry scan.
- Closing releases the ZIP and caches, including after context-body errors.
  Accessing a closed session raises RuntimeError.
- No openpyxl workbook/worksheet model is created. Existing value/formula
  normalization and openpyxl scalar format, date and formula-translation helpers
  are reused; openpyxl remains a dependency.
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
  formulas. Unresolved shared followers are omitted; unsupported translation
  is logged and skipped.
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

## Relationships

`ooxml_package.read_relationships(archive, source_path)` resolves IDs to type,
target and external status. Internal relative paths resolve against the source
directory; absolute targets resolve from the ZIP root. External targets are
preserved and never fetched or read as ZIP parts. These primitives serve
hyperlinks, tables, worksheets, drawings and charts. Legacy drawing helpers
remain available as wrappers.

## ADR assessment

- Verdict: `not-needed`; next action: `no-adr`.
- Rationale: internal parsing capabilities are added without changing backend
  selection, priority, mode meanings, public output or fallback policy. A future
  pipeline migration requires a fresh ADR assessment.
- Domains: extraction implementation and resource reuse.
- Existing decisions: ADR-0010 (light OOXML rich baseline), ADR-0002 (rich fallback).
- Evidence triad: this spec and `dev-docs/specs/excel-extraction.md`;
  `src/exstruct/core/ooxml_session.py`, `ooxml_package.py`, `ooxml_drawing.py`,
  `backends/ooxml_backend.py`; `tests/core/test_ooxml_extraction_session.py`,
  `test_ooxml_package.py`, `test_ooxml_drawing.py` and existing pipeline tests.

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
