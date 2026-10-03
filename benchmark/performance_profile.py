"""Temporary benchmark-only instrumentation; no production hooks required."""

from __future__ import annotations

from collections.abc import Callable
from contextlib import ExitStack
from functools import wraps
import sys
from time import perf_counter
from typing import TYPE_CHECKING, Any
from unittest.mock import patch
import zipfile

if TYPE_CHECKING:
    from exstruct.core.pipeline import ExtractionInputs, PipelineResult

STAGES = (
    "cell_extraction",
    "print_area_extraction",
    "formula_extraction",
    "color_extraction",
    "merged_cell_extraction",
    "table_detection",
    "rich_extraction",
    "model_construction",
    "workbook_parsing",
)


class StageRecorder:
    """Record inclusive durations; nested parsing is intentionally not additive."""

    def __init__(self) -> None:
        self.stages = {name: {"ms": 0.0, "calls": 0} for name in STAGES}
        self.archive_opens = 0
        self.state: dict[str, Any] = {}

    def timed(self, name: str, function: Callable[..., Any]) -> Callable[..., Any]:
        @wraps(function)
        def wrapped(*args: object, **kwargs: object) -> object:
            started = perf_counter()
            try:
                return function(*args, **kwargs)
            finally:
                self.stages[name]["ms"] += (perf_counter() - started) * 1000
                self.stages[name]["calls"] += 1

        return wrapped

    def install(self, stack: ExitStack) -> None:
        """Patch live call sites after lazy extraction imports have completed."""
        import openpyxl

        from exstruct.core import integrate, pipeline
        from exstruct.core.backends.ooxml_backend import OoxmlRichBackend
        from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend

        methods = {
            "extract_cells": "cell_extraction",
            "extract_print_areas": "print_area_extraction",
            "extract_formulas_map": "formula_extraction",
            "extract_colors_map": "color_extraction",
            "extract_merged_cells": "merged_cell_extraction",
            "detect_tables": "table_detection",
        }
        for method, stage in methods.items():
            stack.enter_context(
                patch.object(
                    OpenpyxlBackend,
                    method,
                    self.timed(stage, getattr(OpenpyxlBackend, method)),
                )
            )
        functions = {
            "detect_tables": "table_detection",
            "build_workbook_data": "model_construction",
            "step_extract_shapes_com": "rich_extraction",
            "step_extract_charts_com": "rich_extraction",
            "step_extract_print_areas_com": "print_area_extraction",
            "step_extract_formulas_map_com": "formula_extraction",
            "step_extract_colors_map_com": "color_extraction",
        }
        for function, stage in functions.items():
            stack.enter_context(
                patch.object(
                    pipeline, function, self.timed(stage, getattr(pipeline, function))
                )
            )
        for method in ("extract_shapes", "extract_charts"):
            stack.enter_context(
                patch.object(
                    OoxmlRichBackend,
                    method,
                    self.timed("rich_extraction", getattr(OoxmlRichBackend, method)),
                )
            )

        # Match function identity to include imported aliases and pandas' reader.
        original_load = openpyxl.load_workbook
        wrapped_load = self.timed("workbook_parsing", original_load)
        for module in list(sys.modules.values()):
            if module is None:
                continue
            for name, value in list(vars(module).items()):
                if value is original_load:
                    stack.enter_context(patch.object(module, name, wrapped_load))

        original_zip_init = zipfile.ZipFile.__init__

        def archive_init(
            archive: zipfile.ZipFile, *args: object, **kwargs: object
        ) -> None:
            self.archive_opens += 1
            original_zip_init(archive, *args, **kwargs)

        stack.enter_context(patch.object(zipfile.ZipFile, "__init__", archive_init))
        original_pipeline = integrate.run_extraction_pipeline

        def run_pipeline(inputs: ExtractionInputs) -> PipelineResult:
            result = original_pipeline(inputs)
            self.state = {
                "com_attempted": result.state.com_attempted,
                "com_succeeded": result.state.com_succeeded,
                "fallback_reason": result.state.fallback_reason.value
                if result.state.fallback_reason
                else None,
            }
            return result

        stack.enter_context(
            patch.object(integrate, "run_extraction_pipeline", run_pipeline)
        )
