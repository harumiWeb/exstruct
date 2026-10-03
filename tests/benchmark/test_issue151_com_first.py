"""Failure isolation for the non-public COM-first experiment."""

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from benchmark import issue151_com_first as experiment
import pytest

from exstruct.core import pipeline
from exstruct.errors import FallbackReason


@pytest.mark.parametrize("failure", ["startup", "cells", "rich", "model"])
def test_experiment_falls_back_after_cleanup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, failure: str
) -> None:
    from openpyxl import Workbook

    path = tmp_path / "fallback.xlsx"
    book = Workbook()
    book.active["B3"] = "saved"
    book.save(path)
    book.close()
    inputs = experiment.inputs_for(path, "standard")
    monkeypatch.setenv("SKIP_COM_TESTS", "1")
    expected = pipeline._run_openpyxl_pipeline(inputs)
    monkeypatch.delenv("SKIP_COM_TESTS")
    events = []

    @contextmanager
    def com_scope(_: Path) -> Iterator[Any]:
        if failure == "startup":
            raise RuntimeError("startup")
        try:
            yield object()
        finally:
            events.append("closed")

    def cells(*args: object, **kwargs: object) -> dict[str, list[Any]]:
        if failure == "cells":
            raise RuntimeError("cells")
        return {"partial": []}

    def rich(*args: object, **kwargs: object) -> None:
        if failure == "rich":
            raise RuntimeError("rich")

    def model(*args: object, **kwargs: object) -> None:
        raise RuntimeError("model")

    original = pipeline._run_openpyxl_pipeline

    def fallback(resolved: pipeline.ExtractionInputs) -> pipeline.PipelineResult:
        assert events == ([] if failure == "startup" else ["closed"])
        events.append("fallback")
        return original(resolved)

    monkeypatch.setattr(pipeline, "xlwings_workbook", com_scope)
    monkeypatch.setattr(experiment.ComBackend, "extract_cells", cells)
    monkeypatch.setattr(pipeline, "run_com_pipeline", rich)
    monkeypatch.setattr(pipeline, "collect_sheet_raw_data", lambda **kwargs: {})
    monkeypatch.setattr(experiment, "build_workbook_data", model)
    monkeypatch.setattr(pipeline, "_run_openpyxl_pipeline", fallback)
    result = experiment.com_first(inputs)
    assert result.workbook == expected.workbook
    assert result.artifacts.cell_data == expected.artifacts.cell_data
    assert result.state.com_attempted == (failure != "startup")
    assert not result.state.com_succeeded
    assert result.state.fallback_reason == (
        FallbackReason.COM_UNAVAILABLE
        if failure == "startup"
        else FallbackReason.COM_PIPELINE_FAILED
    )
    import os

    assert "SKIP_COM_TESTS" not in os.environ


@pytest.mark.parametrize("candidate_success, expected", [(False, None), (True, True)])
def test_compatibility_requires_success_on_both_sides(
    candidate_success: bool, expected: bool | None
) -> None:
    records = [
        {"backend": "current", "com_succeeded": True, "output_sha256": "saved"},
        {
            "backend": "com-first",
            "com_succeeded": candidate_success,
            "output_sha256": "saved",
        },
        {"backend": "com-first", "com_succeeded": False, "output_sha256": "fallback"},
    ]
    assert experiment.output_compatible(records) is expected
