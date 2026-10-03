# Issue #150 light backend comparison

Windows / Python 3.11.13, five repeated extractions in separate fresh processes
for each backend. The baseline is the preserved `_run_openpyxl_pipeline`, and
the candidate is the default public light pipeline. Both use the same installed
dependencies and input bytes. Default include flags are used. The full emitted
workbook model SHA-256 matches for every input.

| Fixture | Repeated extraction median, openpyxl / OOXML (ms) | Cold in-process, openpyxl / OOXML (ms) | Peak working set, openpyxl / OOXML (MiB) |
| --- | --- | --- | --- |
| small | 10.8 / 5.0 | 983.4 / 428.1 | 64.4 / 39.6 |
| large (60,000 cells) | 1017.5 / 607.5 | 1626.2 / 1060.1 | 109.3 / 71.8 |
| many-sheet (24 sheets) | 107.7 / 74.1 | 839.9 / 333.2 | 74.5 / 42.7 |
| sparse | 51.3 / 16.9 | 823.0 / 254.0 | 75.3 / 39.7 |
| style-heavy | 136.0 / 110.5 | 902.1 / 460.5 | 73.6 / 46.3 |
| table-heavy (20 tables) | 86.8 / 56.3 | 823.3 / 313.1 | 75.3 / 44.0 |
| tracked shape/connector diagram | 16.2 / 5.9 | 842.9 / 238.9 | 64.6 / 39.6 |

All candidate runs opened one ZIP and zero openpyxl workbooks, compared with
two ZIPs and one openpyxl workbook on the compatibility baseline. Candidate
runs imported none of openpyxl, pandas, SciPy or xlwings and used no fallback.
The default direct table path uses Python clustering even when SciPy is installed;
the compatibility baseline retains its default automatic SciPy selection.

Cold in-process time includes pipeline import, setup and first extraction; it
excludes OS process creation and serialization. `worker_process_ms` in the raw
data includes repeats and instrumentation and is not a cold-start measurement.
Peak memory is the process lifetime working-set high-water mark after repeated
extraction, captured before resource-count instrumentation. These local
sequential measurements establish the acceptance target for these fixtures,
not a universal speed or memory guarantee. Unsupported workbook fallback may
be slower because it runs both readers.

Raw measurements, hashes and environment metadata:
[issue150-2026-10-03-windows.json](baselines/issue150-2026-10-03-windows.json).

## Reproduce

Use a new fixture directory; generation refuses to overwrite existing fixtures.

```powershell
rtk uv run python -m benchmark.performance --generate-fixtures tasks/issue150-fixtures
rtk uv run python -m benchmark.issue150_light --input tasks/issue150-fixtures/small.xlsx tasks/issue150-fixtures/large.xlsx tasks/issue150-fixtures/many-sheet.xlsx tasks/issue150-fixtures/sparse.xlsx tasks/issue150-fixtures/style-heavy.xlsx tasks/issue150-fixtures/table-heavy.xlsx sample/flowchart/sample-shape-connector.xlsx --repeats 5 --output benchmark/baselines/issue150-2026-10-03-windows.json
```

The implementation caches date/duration classification per style and reuses
parsed coordinates. Fully qualified child tags avoid repeated XPath processing
without changing the safe XML parser or cached-value semantics.
