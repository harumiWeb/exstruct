"""Compare direct OOXML light extraction with its preserved openpyxl pipeline."""

from __future__ import annotations

import argparse
from datetime import UTC, datetime
import hashlib
from functools import wraps
from importlib.metadata import PackageNotFoundError, version
import json
from pathlib import Path
import platform
import statistics
import subprocess
import sys
from time import perf_counter
from typing import Any


def worker(path: Path, backend: str, repeats: int) -> dict[str, Any]:
    started = perf_counter()
    from exstruct.core import pipeline, workbook
    from exstruct.core.ooxml_session import OoxmlExtractionSession

    import_ms = (perf_counter() - started) * 1000
    from benchmark.performance import peak_rss_bytes

    inputs = pipeline.resolve_extraction_inputs(
        path,
        mode="light",
        include_cell_links=None,
        include_print_areas=None,
        include_auto_page_breaks=False,
        include_colors_map=None,
        include_default_background=False,
        ignore_colors=None,
        include_formulas_map=None,
        include_merged_cells=None,
        include_merged_values_in_rows=True,
    )
    runner = (
        pipeline.run_extraction_pipeline
        if backend == "ooxml"
        else pipeline._run_openpyxl_pipeline
    )
    start = perf_counter()
    result = runner(inputs)
    first_ms = (perf_counter() - start) * 1000
    cold_ms = (perf_counter() - started) * 1000
    samples = []
    for _ in range(repeats):
        start = perf_counter()
        runner(inputs)
        samples.append((perf_counter() - start) * 1000)
    peak = peak_rss_bytes()
    loaded = [
        name
        for name in ("openpyxl", "xlwings", "scipy", "pandas")
        if name in sys.modules
    ]
    encoded = json.dumps(
        result.workbook.model_dump(mode="json"), ensure_ascii=True, sort_keys=True
    )
    loads = 0
    archives = 0
    original_load = workbook.load_workbook
    original_archive = OoxmlExtractionSession.archive
    from zipfile import ZipFile

    original_zip_init = ZipFile.__init__

    @wraps(original_load)
    def load(*args: Any, **kwargs: Any) -> Any:
        nonlocal loads
        loads += 1
        return original_load(*args, **kwargs)

    def zip_init(self: ZipFile, *args: Any, **kwargs: Any) -> None:
        nonlocal archives
        archives += 1
        original_zip_init(self, *args, **kwargs)

    # Count all ZIP constructors, including the previous standalone rich parser.
    workbook.load_workbook = load
    ZipFile.__init__ = zip_init
    try:
        profiled = runner(inputs)
    finally:
        workbook.load_workbook = original_load
        ZipFile.__init__ = original_zip_init
    assert profiled.workbook == result.workbook
    assert OoxmlExtractionSession.archive is original_archive
    return {
        "backend": backend,
        "import_ms": import_ms,
        "first_extract_ms": first_ms,
        "cold_in_process_ms": cold_ms,
        "repeated_extract_ms": samples,
        "median_extract_ms": statistics.median(samples),
        "peak_rss_bytes": peak,
        "workbook_opens": loads,
        "archive_opens": archives,
        "loaded": loaded,
        "fallback_reason": result.state.fallback_reason,
        "output_sha256": hashlib.sha256(encoded.encode()).hexdigest(),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, nargs="+", required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--worker", choices=["ooxml", "openpyxl"])
    args = parser.parse_args()
    if args.repeats < 1:
        parser.error("--repeats must be positive")
    if args.worker:
        print(json.dumps(worker(args.input[0], args.worker, args.repeats)))
        return
    results = []
    for path in args.input:
        measurements = {}
        # Separate processes ensure compatibility imports cannot contaminate OOXML.
        for backend in ("openpyxl", "ooxml"):
            started = perf_counter()
            completed = subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "benchmark.issue150_light",
                    "--worker",
                    backend,
                    "--input",
                    str(path),
                    "--repeats",
                    str(args.repeats),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                check=True,
                timeout=300,
            )
            measurement = json.loads(completed.stdout)
            measurement["worker_process_ms"] = (perf_counter() - started) * 1000
            measurements[backend] = measurement
        if (
            measurements["openpyxl"]["output_sha256"]
            != measurements["ooxml"]["output_sha256"]
        ):
            raise RuntimeError(f"Output mismatch: {path}")
        results.append(
            {
                "input": str(path),
                "input_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                **measurements,
            }
        )
    from benchmark.performance import source_metadata

    dependencies = {}
    for name in ("exstruct", "openpyxl", "numpy", "scipy", "defusedxml", "pydantic"):
        try:
            dependencies[name] = version(name)
        except PackageNotFoundError:
            dependencies[name] = None
    output = (
        json.dumps(
            {
                "recorded_at": datetime.now(UTC).isoformat(),
                "python": sys.version,
                "platform": platform.platform(),
                "source": source_metadata(10),
                "dependencies": dependencies,
                "results": results,
            },
            indent=2,
        )
        + "\n"
    )
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output, encoding="utf-8")
    else:
        print(output, end="")


if __name__ == "__main__":
    main()
