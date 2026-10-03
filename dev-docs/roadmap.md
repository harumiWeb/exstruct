# Extraction roadmap

This records the extraction work established by the issues. Current contracts
are defined by the specifications and ADRs linked below.

| Milestone | Status | Contract and evidence |
| --- | --- | --- |
| Core OOXML session foundation (Issue #149, PR #154) | Merged | Direct cells, strings, formulas, links, merges, names, print areas and explicit tables; shared ZIP/relationships for optional drawings. See [core OOXML specification](specs/ooxml-core-extraction.md) and tests/core/test_ooxml_extraction_session.py. |
| Default light OOXML pipeline (Issue #150) | Implemented and locally verified; pending merge | .xlsx / .xlsm light defaults to one shared OOXML session for core and rich output. Unsupported input, colors_map, or legacy overrides restart the complete openpyxl compatibility pipeline. See [ADR-0013](adr/ADR-0013-default-ooxml-backend-for-light-extraction.md), [extraction spec](specs/excel-extraction.md), tests/core/test_light_ooxml_pipeline.py, and [performance results](specs/extraction-performance.md). |
| Light OOXML table heuristics (Issue #150) | Implemented and locally verified; pending merge | Preserve table-candidate behavior with direct OOXML heuristics. The direct light path selects Python BFS per call, without changing the clustering environment variable; see [ADR-0012](adr/ADR-0012-optional-scipy-border-clustering.md). |
| Issue #150 performance comparison | Recorded | benchmark/issue150_light.py compared seven inputs in fresh openpyxl and OOXML workers; all serialized-output hashes matched. See [measured results](specs/extraction-performance.md) and benchmark/baselines/issue150-2026-10-03-windows.json. |

Current behavior is defined by [Excel extraction](specs/excel-extraction.md)
and [core OOXML extraction](specs/ooxml-core-extraction.md).
