"""Measurement contracts without machine-dependent timing thresholds."""

from contextlib import ExitStack
import hashlib
from io import StringIO
import json
from pathlib import Path
import platform
import subprocess
import sys
from typing import Any

from benchmark import performance, performance_fixtures
from benchmark.performance_fixtures import CATEGORIES, generate_fixtures
from benchmark.performance_profile import StageRecorder
import pytest


def test_fixtures_are_reproducible_and_refuse_overwrite(tmp_path: Path) -> None:
    from openpyxl import load_workbook

    first = generate_fixtures(tmp_path / "first")
    second = generate_fixtures(tmp_path / "second")
    baseline = json.loads(
        (
            Path(__file__).resolve().parents[2]
            / "benchmark/baselines/2026-10-03-windows.json"
        ).read_text(encoding="utf-8")
    )
    recorded_hashes = {
        record["input"]["name"]: record["input"]["sha256"] for record in baseline
    }
    assert [path.stem for path in first] == list(CATEGORIES)
    for left, right in zip(first, second, strict=True):
        assert left.read_bytes() == right.read_bytes()
        # ZIP bytes may vary across platforms/Python/dependency versions; compare
        # the historical hash only in the recorded generation environment.
        if (
            performance.version("openpyxl")
            == baseline[0]["environment"]["dependencies"]["openpyxl"]
            and sys.version == baseline[0]["environment"]["python"]
            and platform.platform() == baseline[0]["environment"]["platform"]
        ):
            assert (
                hashlib.sha256(left.read_bytes()).hexdigest()
                == recorded_hashes[left.name]
            )
        workbook = load_workbook(left)
        sheets, rows, columns = CATEGORIES[left.stem]
        assert len(workbook.worksheets) == sheets
        assert workbook.worksheets[0].max_row == rows
        # small includes a chart outside the cell grid, which doesn't alter it.
        assert workbook.worksheets[0].max_column == columns
        workbook.close()
    with pytest.raises(FileExistsError):
        generate_fixtures(tmp_path / "first")


def test_recorder_restores_wrappers_after_exception() -> None:
    import openpyxl

    from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend

    original = OpenpyxlBackend.extract_cells
    load = openpyxl.load_workbook
    with pytest.raises(RuntimeError), ExitStack() as stack:
        StageRecorder().install(stack)
        assert OpenpyxlBackend.extract_cells is not original
        assert openpyxl.load_workbook is not load
        raise RuntimeError("test")
    assert OpenpyxlBackend.extract_cells is original
    assert openpyxl.load_workbook is load


def test_timer_records_failure_without_swallowing_it() -> None:
    def fail() -> None:
        raise ValueError("expected")

    recorder = StageRecorder()
    with pytest.raises(ValueError, match="expected"):
        recorder.timed("cell_extraction", fail)()
    assert recorder.stages["cell_extraction"]["calls"] == 1


def test_fresh_worker_measures_light_and_standard_fallback(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from openpyxl import Workbook

    path = tmp_path / "tiny.xlsx"
    workbook = Workbook()
    workbook.active.append(["Name", "Count"])
    workbook.active.append(["Example", 7])
    workbook.save(path)
    workbook.close()
    monkeypatch.setenv("SKIP_COM_TESTS", "1")
    for mode in ("light", "standard"):
        _, output = performance.fresh_process(
            [
                "-m",
                "benchmark.performance",
                "--worker",
                "--input",
                str(path),
                "--mode",
                mode,
                "--repeats",
                "2",
            ],
            60,
        )
        result = json.loads(output)
        assert len(result["repeated"]) == 2
        assert result["modules_after_import"] == []
        assert isinstance(result["modules_after_extraction"], list)
        stages = result["profile"]["stages"]
        assert stages["cell_extraction"]["calls"] == 1
        assert stages["table_detection"]["calls"] == 1
        assert stages["model_construction"]["calls"] == 1
        assert stages["formula_extraction"]["calls"] == 0
        assert result["profile"]["openpyxl_workbook_opens"] == (
            0 if mode == "light" else 1
        )
        # Direct light shares one ZIP; other paths also count standalone
        # drawing reads. Workbook loaders and ZIP constructors are distinct.
        assert result["profile"]["archive_opens"] >= 1
        state = result["profile"]["pipeline_state"]
        assert state["com_succeeded"] is False
        assert state["fallback_reason"] == (
            "skip_com_tests" if mode == "standard" else None
        )
        assert result["output_bytes"] > 0
        assert result["peak_rss_bytes"] is None or result["peak_rss_bytes"] > 0


@pytest.mark.parametrize(
    "arguments",
    [
        ["--repeats", "0"],
        ["--timeout", "nan"],
        ["--timeout", "inf"],
        ["--timeout", "0"],
        ["--input", "missing.xlsx"],
    ],
)
def test_invalid_arguments(arguments: list[str]) -> None:
    with pytest.raises(SystemExit) as error:
        performance.main(arguments)
    assert error.value.code == 2


def test_subprocess_errors_are_reported(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    path = tmp_path / "input.xlsx"
    path.write_bytes(b"placeholder")

    def fail(*args: Any, **kwargs: Any) -> dict[str, Any]:  # noqa: ANN401
        raise subprocess.TimeoutExpired("python", 1)

    monkeypatch.setattr(performance, "run_benchmark", fail)
    assert performance.main(["--input", str(path)]) == 1
    assert "Benchmark failed" in capsys.readouterr().err


def test_optional_profile_and_supervisor_report(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from openpyxl import Workbook

    path = tmp_path / "optional.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "=1+2"
    workbook.save(path)
    workbook.close()
    monkeypatch.setenv("SKIP_COM_TESTS", "1")
    result = performance.run_benchmark(path, "light", 1, 60, profile_all_features=True)
    assert result["schema_version"] == 1
    assert result["mode"] == "light"
    assert len(result["input"]["sha256"]) == 64
    assert result["environment"]["SKIP_COM_TESTS"] == "1"
    assert result["profile"]["all_features"] is True
    assert result["profile"]["openpyxl_workbook_opens"] == 2
    for stage in ("formula_extraction", "color_extraction", "merged_cell_extraction"):
        assert result["profile"]["stages"][stage]["calls"] == 1
    assert result["startup"]["cold_extraction_process_ms"] > 0
    with pytest.raises(SystemExit):
        performance.main(["--input", str(path), "--output", str(path)])


def test_successful_child_stderr_does_not_pollute_json(
    capsys: pytest.CaptureFixture[str],
) -> None:
    _, output = performance.fresh_process(
        [
            "-c",
            "import sys; print('fallback warning', file=sys.stderr); print('{\"ok\": true}')",
        ],
        30,
    )
    assert json.loads(output) == {"ok": True}
    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err == "fallback warning\n"


def test_forwarding_stderr_is_outside_child_timing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    elapsed = 0.0

    class SlowStderr(StringIO):
        def write(self, value: str) -> int:
            nonlocal elapsed
            elapsed += 100
            return super().write(value)

    def completed(*args: object, **kwargs: object) -> subprocess.CompletedProcess[str]:
        nonlocal elapsed
        elapsed += 2
        return subprocess.CompletedProcess("python", 0, stdout="{}", stderr="warning")

    monkeypatch.setattr(performance, "perf_counter", lambda: elapsed)
    monkeypatch.setattr(performance.subprocess, "run", completed)
    monkeypatch.setattr(performance.sys, "stderr", SlowStderr())
    measured_ms, output = performance.fresh_process(["-c", "pass"], 30)
    assert measured_ms == 2000
    assert output == "{}"
    assert elapsed > 2


@pytest.mark.parametrize("timeout_call", [1, 2])
def test_git_metadata_timeouts_return_nulls(
    timeout_call: int, monkeypatch: pytest.MonkeyPatch
) -> None:
    calls: list[object] = []

    def git(*args: object, **kwargs: object) -> subprocess.CompletedProcess[str]:
        calls.append(args[0])
        assert kwargs["timeout"] == 0.5
        if len(calls) == timeout_call:
            raise subprocess.TimeoutExpired("git", 0.5)
        return subprocess.CompletedProcess("git", 0, stdout="revision\n")

    monkeypatch.setattr(performance.subprocess, "run", git)
    assert performance.source_metadata(0.5) == {"revision": None, "dirty": None}
    assert len(calls) == timeout_call


def test_generation_interruption_leaves_directory_retryable(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    directory = tmp_path / "fixtures"
    directory.mkdir()
    marker = directory / "keep.txt"
    marker.write_text("existing", encoding="utf-8")
    original = performance_fixtures._write_fixture

    def interrupted(path: Path, *dimensions: int) -> None:
        if path.stem == "large":
            raise KeyboardInterrupt("generation interrupted")
        original(path, *dimensions)

    with monkeypatch.context() as patcher:
        patcher.setattr(performance_fixtures, "_write_fixture", interrupted)
        with pytest.raises(KeyboardInterrupt):
            generate_fixtures(directory)
    assert list(directory.iterdir()) == [marker]
    assert len(generate_fixtures(directory)) == 6
    assert marker.read_text(encoding="utf-8") == "existing"


def test_publication_collision_preserves_other_writers_files(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    directory = tmp_path / "fixtures"
    collision = directory / "large.xlsx"

    def concurrent_writer(path: Path, *dimensions: int) -> None:
        path.write_bytes(b"generated")
        if path.stem == "table-heavy":
            collision.write_bytes(b"other writer")

    monkeypatch.setattr(performance_fixtures, "_write_fixture", concurrent_writer)
    with pytest.raises(FileExistsError):
        generate_fixtures(directory)
    assert list(directory.iterdir()) == [collision]
    assert collision.read_bytes() == b"other writer"


def test_published_baseline_has_no_identifying_interpreter_path() -> None:
    path = (
        Path(__file__).resolve().parents[2]
        / "benchmark/baselines/2026-10-03-windows.json"
    )
    records = json.loads(path.read_text(encoding="utf-8"))
    assert len(records) == 12
    assert all(
        record["environment"]["executable"] == "<local-python>" for record in records
    )
