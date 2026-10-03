# Extraction roadmap

This records the scope already established by the extraction issues. It does
not introduce a new public mode or approve a default-backend migration.

| Milestone | Status | Contract and evidence |
| --- | --- | --- |
| Core OOXML session foundation (Issue #149, PR #154) | Implemented; pending merge | Direct cells, strings, formulas, links, merges, names, print areas and explicit tables; shared ZIP/relationships for optional drawings. See [core OOXML specification](specs/ooxml-core-extraction.md) and `tests/core/test_ooxml_extraction_session.py`. |
| Default-path OOXML migration | Deferred to a separate issue | Current openpyxl pipeline remains selected. Evaluate parity and capability gaps, then assess the migration's ADR and public-contract impact before changing selection. |
| OOXML table heuristics | Outside Issue #149 | Explicit tables are available now; border/style-based heuristics remain on the existing backend. Any extension needs its own scope and parity evidence. |

Current behavior is defined by [Excel extraction](specs/excel-extraction.md),
not by the roadmap status labels.
