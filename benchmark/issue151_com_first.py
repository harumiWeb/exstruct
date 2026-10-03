"""Internal COM-first experiment; public extraction remains file-first."""

from __future__ import annotations

import argparse
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import replace
from datetime import UTC, datetime
import hashlib
from importlib.metadata import version
import json
import os
from pathlib import Path
import platform
from time import perf_counter
from typing import Any
from unittest.mock import patch

from exstruct.core import pipeline
from exstruct.core.backends.com_backend import ComBackend
from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend
from exstruct.core.modeling import WorkbookRawData, build_workbook_data
from exstruct.core.openpyxl_session import OpenpyxlExtractionSession
from exstruct.errors import FallbackReason


def com_first(inputs: pipeline.ExtractionInputs) -> pipeline.PipelineResult:
    """Attempt COM before file work; discard partial artifacts on failure."""
    if inputs.mode not in {"standard", "verbose"}:
        raise ValueError("Experiment requires standard or verbose")
    if os.getenv("SKIP_COM_TESTS"):
        return pipeline._run_openpyxl_pipeline(inputs)
    attempted = False
    try:
        with pipeline.xlwings_workbook(inputs.file_path) as book:
            attempted = True
            # Tables/merged ranges and file formula semantics stay file-based.
            # Their shared session loads only if the requested stage needs it.
            with OpenpyxlExtractionSession(inputs.file_path) as session:
                file_backend = OpenpyxlBackend(inputs.file_path, session=session)
                artifacts = pipeline.ExtractionArtifacts(openpyxl_session=session)
                artifacts.cell_data = ComBackend(book).extract_cells(
                    include_links=inputs.include_cell_links
                )
                if inputs.include_merged_cells:
                    artifacts.merged_cell_data = file_backend.extract_merged_cells()
                if inputs.include_print_areas:
                    artifacts.print_area_data = file_backend.extract_print_areas()
                if inputs.include_formulas_map and not inputs.use_com_for_formulas:
                    artifacts.formulas_map_data = file_backend.extract_formulas_map()
                pipeline.run_com_pipeline(
                    pipeline.build_com_pipeline(inputs), inputs, artifacts, book
                )
                sheets = pipeline.collect_sheet_raw_data(
                    cell_data=artifacts.cell_data,
                    shape_data=artifacts.shape_data,
                    chart_data=artifacts.chart_data,
                    merged_cell_data=artifacts.merged_cell_data,
                    workbook=book,
                    mode=inputs.mode,
                    include_merged_values_in_rows=inputs.include_merged_values_in_rows,
                    print_area_data=artifacts.print_area_data,
                    auto_page_break_data=artifacts.auto_page_break_data,
                    formulas_map_data=artifacts.formulas_map_data,
                    colors_map_data=artifacts.colors_map_data,
                    openpyxl_session=session,
                )
                return pipeline.PipelineResult(
                    workbook=build_workbook_data(
                        WorkbookRawData(book_name=inputs.file_path.name, sheets=sheets)
                    ),
                    artifacts=artifacts,
                    state=pipeline.PipelineState(
                        com_attempted=True, com_succeeded=True
                    ),
                )
    except Exception as exc:
        pipeline.logger.warning("COM-first experiment failed: %r", exc)
        # Resource scopes have ended before the clean file-only fallback begins.
        with patch.dict(os.environ, {"SKIP_COM_TESTS": "1"}):
            result = pipeline._run_openpyxl_pipeline(inputs)
        result.state.com_attempted = attempted
        result.state.fallback_reason = (
            FallbackReason.COM_PIPELINE_FAILED
            if attempted
            else FallbackReason.COM_UNAVAILABLE
        )
        return result


def inputs_for(path: Path, mode: str) -> pipeline.ExtractionInputs:
    inputs = pipeline.resolve_extraction_inputs(
        path.resolve(),
        mode="standard",
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
    if mode == "verbose":
        inputs = replace(
            inputs,
            mode="verbose",
            include_cell_links=True,
            include_formulas_map=True,
            include_colors_map=True,
        )
    return inputs


def digest(result: pipeline.PipelineResult) -> str:
    encoded = json.dumps(result.workbook.model_dump(mode="json"), sort_keys=True)
    return hashlib.sha256(encoded.encode()).hexdigest()


def output_compatible(records: list[dict[str, Any]]) -> bool | None:
    """Compare successful outputs only; missing success on either side is unknown."""
    hashes = {
        backend: {
            record["output_sha256"]
            for record in records
            if record["backend"] == backend and record["com_succeeded"]
        }
        for backend in ("current", "com-first")
    }
    if not all(hashes.values()):
        return None
    return len(hashes["current"] | hashes["com-first"]) == 1


def _quit_owned_app(owner: Any) -> None:  # noqa: ANN401 - benchmark COM duck type
    try:
        owner.quit()
    except Exception:
        owner.kill()


def measure(
    path: Path, mode: str, scenario: str, repeats: int, *, include_colors: bool = True
) -> dict[str, Any]:
    import xlwings as xw

    inputs = inputs_for(path, mode)
    if not include_colors:
        inputs = replace(inputs, include_colors_map=False)
    owner = None
    book = None
    before = sorted(xw.apps.keys())
    if before:
        raise RuntimeError(
            "Close existing Excel instances before controlled measurements"
        )
    setup_started = perf_counter()
    try:
        if scenario != "cold":
            owner = xw.App(visible=False, add_book=False)
            owner.display_alerts = False
            if scenario == "already-open":
                book = owner.books.open(str(inputs.file_path))
        setup_ms = (perf_counter() - setup_started) * 1000
        excel_version = str(owner.version) if owner is not None else None
        original = pipeline.xlwings_workbook
        records = []
        for repeat in range(repeats):
            # Alternate order to reduce systematic cache/order bias.
            for backend in (
                ("current", "com-first")
                if repeat % 2 == 0
                else ("com-first", "current")
            ):
                phases: dict[str, float] = {}

                @contextmanager
                def timed_workbook(
                    file_path: Path, phases: dict[str, float] = phases
                ) -> Iterator[Any]:
                    start = perf_counter()
                    borrowed = scenario == "already-open"
                    if owner is None or borrowed:
                        with original(file_path) as workbook:
                            nonlocal excel_version
                            excel_version = str(workbook.app.version)
                            phases["com_acquire_ms"] = (perf_counter() - start) * 1000
                            try:
                                yield workbook
                            finally:
                                cleanup_started = perf_counter()
                        phases["com_cleanup_ms"] = (
                            perf_counter() - cleanup_started
                        ) * 1000
                    else:
                        workbook = owner.books.open(str(file_path))
                        phases["com_acquire_ms"] = (perf_counter() - start) * 1000
                        try:
                            yield workbook
                        finally:
                            start = perf_counter()
                            workbook.close()
                            phases["com_cleanup_ms"] = (perf_counter() - start) * 1000

                started = perf_counter()
                with patch.object(pipeline, "xlwings_workbook", timed_workbook):
                    result = (
                        pipeline.run_extraction_pipeline
                        if backend == "current"
                        else com_first
                    )(inputs)
                total = (perf_counter() - started) * 1000
                records.append(
                    {
                        "repeat": repeat,
                        "backend": backend,
                        "total_ms": total,
                        **phases,
                        "non_acquire_cleanup_ms": total - sum(phases.values()),
                        "com_succeeded": result.state.com_succeeded,
                        "fallback_reason": result.state.fallback_reason,
                        "output_sha256": digest(result),
                    }
                )
        return {
            "input": path.as_posix(),
            "input_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            "mode": mode,
            "include_colors_map": inputs.include_colors_map,
            "scenario": scenario,
            "scenario_setup_ms": setup_ms,
            "excel_version": excel_version,
            "measurements": records,
            "output_compatible": output_compatible(records),
        }
    finally:
        try:
            if book is not None:
                book.close()
        finally:
            if owner is not None:
                _quit_owned_app(owner)
        after = sorted(xw.apps.keys())
        if after != before:
            raise RuntimeError(
                f"Excel process cleanup mismatch: before={before}, after={after}"
            )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, nargs="+", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument(
        "--no-colors",
        action="store_true",
        help="Disable the existing per-cell color stage to isolate large-workbook cell architecture",
    )
    parser.add_argument(
        "--modes", nargs="+", choices=["standard", "verbose"], default=["standard"]
    )
    parser.add_argument(
        "--scenarios",
        nargs="+",
        choices=["cold", "warm", "already-open"],
        default=["cold", "warm", "already-open"],
    )
    args = parser.parse_args()
    if args.repeats < 1:
        parser.error("--repeats must be positive")
    if any(path.resolve() == args.output.resolve() for path in args.input):
        parser.error("output must differ from inputs")
    import xlwings as xw

    from benchmark.performance import source_metadata

    results = []
    for path in args.input:
        for mode in args.modes:
            for scenario in args.scenarios:
                print(f"Measuring {path.name} {mode} {scenario}", flush=True)
                results.append(
                    measure(
                        path,
                        mode,
                        scenario,
                        args.repeats,
                        include_colors=not args.no_colors,
                    )
                )
                args.output.parent.mkdir(parents=True, exist_ok=True)
                args.output.write_text(
                    json.dumps(
                        {
                            "recorded_at": datetime.now(UTC).isoformat(),
                            "platform": platform.platform(),
                            "source": source_metadata(10),
                            "xlwings": xw.__version__,
                            "dependencies": {
                                name: version(name)
                                for name in ("openpyxl", "numpy", "pydantic")
                            },
                            "results": results,
                        },
                        indent=2,
                    )
                    + "\n",
                    encoding="utf-8",
                )


if __name__ == "__main__":
    main()
