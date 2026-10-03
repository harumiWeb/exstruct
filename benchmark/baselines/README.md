# Initial extraction performance baseline

[2026-10-03-windows.json](2026-10-03-windows.json) contains 12 raw result records:
six generated categories × `light`/`standard`, three warm repeats per record,
default profile flags. Measured on Windows 10.0.22631, Python 3.11.13, ExStruct
0.8.2 / extraction revision `563da66347dc343150b863ad699f8eb8d94da708`.
The working tree was dirty with benchmark/test/docs additions; production source
was unchanged. Dependencies and input hashes are recorded in every result.
The published interpreter path was redacted to `<local-python>` during review;
all other metadata and measured values are preserved.
`SKIP_COM_TESTS` was unset; all standard **profile** runs succeeded with Excel COM.
Runtime state is observed in the separate profile, not in each latency sample.

| Category | light cold process (ms) | light warm extraction median (ms) | standard cold process (ms) | standard warm extraction median (ms) |
| --- | ---: | ---: | ---: | ---: |
| small | 1,610.8 | 19.8 | 4,081.1 | 1,674.0 |
| large | 3,984.2 | 2,141.5 | 6,153.0 | 4,117.8 |
| many-sheet | 8,270.8 | 3,360.0 | 9,329.4 | 7,224.4 |
| sparse | 1,664.6 | 66.5 | 3,740.5 | 2,031.1 |
| style-heavy | 2,141.8 | 331.0 | 4,413.3 | 2,603.2 |
| table-heavy | 2,020.7 | 258.2 | 5,234.9 | 2,682.5 |

Cold process times include interpreter startup, import, one extraction,
serialization and teardown. Warm extraction excludes serialization; raw samples
and serialization totals remain in JSON. Do not subtract these columns to infer
an import or COM cost: they are separate observations.

The many-sheet profile observed 50 openpyxl loader calls in light, 51 in standard.
Other categories observed four/five respectively. These are loader calls, not
COM opens; inclusive parsing and table timings overlap. This records a baseline
for future work and makes no claim about attribution beyond the measured calls.

To reproduce individual records, follow [the runner instructions](../README.md).
For all categories, from the repository root:

```powershell
rtk uv run python -m benchmark.performance --generate-fixtures tasks/performance-inputs
foreach ($mode in @('light', 'standard')) {
    foreach ($inputFile in (Get-ChildItem -LiteralPath tasks/performance-inputs -Filter '*.xlsx')) {
        rtk uv run python -m benchmark.performance --mode $mode --input $inputFile.FullName --output "tasks/performance-results/$mode-$($inputFile.BaseName).json"
    }
}
```

Generated fixtures are not committed: generator and SHA256 values identify them.
Use an unused directory or the same previously generated inputs, and keep the
locked dependencies, runtime and machine conditions comparable. Values here are
single-machine observations, not thresholds or expected runtimes on another host.
