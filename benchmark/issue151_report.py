"""Render recorded COM-first timings, excluding failed COM samples."""

import json
from pathlib import Path
import statistics

from benchmark.issue151_com_first import output_compatible


def table(path: Path) -> str:
    data = json.loads(path.read_text(encoding="utf-8"))
    lines = [
        "| Fixture | Scenario | Current ms | Candidate ms | Output equal | Failed COM |",
        "| --- | --- | ---: | ---: | --- | ---: |",
    ]
    for result in data["results"]:
        records = result["measurements"]
        medians = []
        for backend in ("current", "com-first"):
            values = [
                r["total_ms"]
                for r in records
                if r["backend"] == backend and r["com_succeeded"]
            ]
            medians.append(f"{statistics.median(values):.1f}" if values else "n/a")
        compatible = output_compatible(records)
        lines.append(
            f"| {Path(result['input']).stem} | {result['scenario']} | {' | '.join(medians)} | {compatible if compatible is not None else 'unknown'} | {sum(not r['com_succeeded'] for r in records)} |"
        )
    return "\n".join(lines)


if __name__ == "__main__":
    for mode, filename in (
        ("standard", "issue151-2026-10-03-windows.json"),
        ("verbose", "issue151-2026-10-03-verbose-windows.json"),
        ("verbose (colors disabled)", "issue151-verbose-no-colors-windows.json"),
    ):
        print(f"### {mode}\n\n{table(Path('benchmark/baselines') / filename)}\n")
