from __future__ import annotations

import builtins
from collections.abc import Callable
import json
import logging
from pathlib import Path
import shutil
import subprocess
import sys
import threading
import time
from types import ModuleType, SimpleNamespace
from typing import Any, cast

import pytest
import xlwings as xw

from exstruct.errors import MissingDependencyError, RenderError
import exstruct.render as render


class FakeSheet:
    """Minimal sheet stub with a name attribute."""

    def __init__(self, name: str) -> None:
        self.name = name
        self.api = FakeSheetApi()


class FakeSheetApi:
    """Stub of xlwings Sheet.api for PDF export."""

    def ExportAsFixedFormat(self, file_format: int, output_path: str) -> None:
        _ = file_format
        Path(output_path).write_bytes(b"%PDF-1.4")


class FakeBookApi:
    """Stub of xlwings Book.api for PDF export."""

    def ExportAsFixedFormat(self, file_format: int, output_path: str) -> None:
        _ = file_format
        Path(output_path).write_bytes(b"%PDF-1.4")

    def SaveAs(self, output_path: str) -> None:
        Path(output_path).write_bytes(b"XLSX")


class FakeBook:
    """Stub of xlwings Book."""

    def __init__(self, sheet_names: list[str]) -> None:
        self.sheets = [FakeSheet(name) for name in sheet_names]
        self.api = FakeBookApi()
        self.closed = False

    def close(self) -> None:
        """Mark the book as closed."""
        self.closed = True


class FakeBooks:
    """Stub of xlwings Books collection."""

    def __init__(self, sheet_names: list[str], raise_on_open: bool) -> None:
        self._sheet_names = sheet_names
        self._raise_on_open = raise_on_open

    def open(self, path: str) -> FakeBook:
        """Return a fake book or raise to simulate failures."""
        _ = path
        if self._raise_on_open:
            raise ValueError("open failed")
        return FakeBook(self._sheet_names)


class FakeApp:
    """Stub of xlwings App."""

    def __init__(self, sheet_names: list[str], raise_on_open: bool) -> None:
        self.books = FakeBooks(sheet_names, raise_on_open)
        self.display_alerts = True
        self.quit_called = False

    def quit(self) -> None:
        """Mark the app as quit."""
        self.quit_called = True


class FakeBitmap:
    """Stub of a rendered bitmap."""

    def to_pil(self) -> FakeImage:
        """Return a fake image object."""
        return FakeImage()


class FakePage:
    """Stub of a PDF page."""

    def render(self, scale: float) -> FakeBitmap:
        _ = scale
        return FakeBitmap()


class FakePdfDocument:
    """Stub of pdfium.PdfDocument."""

    def __init__(self, path: str) -> None:
        self._path = path
        self._page_count = 2 if "sheet_01" in path else 1

    def __enter__(self) -> FakePdfDocument:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: object | None,
    ) -> bool | None:
        _ = exc_type
        _ = exc
        _ = tb
        return None

    def __getitem__(self, index: int) -> FakePage:
        _ = index
        return FakePage()

    def __len__(self) -> int:
        return self._page_count


class ExplodingPdfDocument:
    """PdfDocument stub that raises on enter."""

    def __init__(self, path: str) -> None:
        _ = path

    def __enter__(self) -> ExplodingPdfDocument:
        raise RuntimeError("boom")

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: object | None,
    ) -> bool | None:
        _ = exc_type
        _ = exc
        _ = tb
        return None


class FakeImage:
    """Stub of a PIL image with a save method."""

    def save(
        self,
        path: Path,
        image_format: str | None = None,
        dpi: tuple[int, int] | None = None,
        **kwargs: object,
    ) -> None:
        _ = image_format
        _ = dpi
        _ = kwargs
        path.write_bytes(b"PNG")


def _fake_app_factory(
    sheet_names: list[str], raise_on_open: bool = False
) -> Callable[..., FakeApp]:
    """Return a factory that mimics xlwings.App."""

    def _factory(*args: object, **kwargs: object) -> FakeApp:
        _ = args
        _ = kwargs
        return FakeApp(sheet_names, raise_on_open)

    return _factory


def test_require_excel_app_success(monkeypatch: pytest.MonkeyPatch) -> None:
    """_require_excel_app returns the constructed app instance."""
    fake_app = FakeApp(["Sheet1"], raise_on_open=False)
    monkeypatch.setattr(xw, "App", lambda *a, **k: fake_app)

    assert render._require_excel_app() is fake_app


def test_require_excel_app_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    """_require_excel_app wraps unexpected errors as RenderError."""

    def _raise(*args: object, **kwargs: object) -> None:
        _ = args
        _ = kwargs
        raise RuntimeError("boom")

    monkeypatch.setattr(xw, "App", _raise)

    with pytest.raises(RenderError, match="Excel \\(COM\\) is not available"):
        render._require_excel_app()


def test_export_pdf_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """export_pdf writes a PDF and returns sheet names."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    output_pdf = tmp_path / "out.pdf"
    sheet_names = ["Sheet1", "Summary"]
    monkeypatch.setattr(xw, "App", _fake_app_factory(sheet_names))

    result = render.export_pdf(xlsx, output_pdf)

    assert result == sheet_names
    assert output_pdf.exists()


def test_export_pdf_wraps_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_pdf wraps unexpected errors as RenderError."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    output_pdf = tmp_path / "out.pdf"
    monkeypatch.setattr(xw, "App", _fake_app_factory(["Sheet1"], raise_on_open=True))

    with pytest.raises(RenderError, match="Failed to export PDF for"):
        render.export_pdf(xlsx, output_pdf)


def test_export_pdf_missing_output_raises(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_pdf raises RenderError when the output file is missing."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    output_pdf = tmp_path / "out.pdf"
    monkeypatch.setattr(xw, "App", _fake_app_factory(["Sheet1"]))

    real_copy = shutil.copy

    def _copy(
        src: Path | str, dst: Path | str, *args: object, **kwargs: object
    ) -> Path:
        _ = args
        _ = kwargs
        if Path(dst) == output_pdf:
            return Path(dst)
        return Path(real_copy(src, dst))

    monkeypatch.setattr(shutil, "copy", _copy)

    with pytest.raises(RenderError, match="Failed to export PDF to"):
        render.export_pdf(xlsx, output_pdf)


def test_require_pdfium_missing_dependency(monkeypatch: pytest.MonkeyPatch) -> None:
    """_require_pdfium raises MissingDependencyError when import fails."""
    real_import = builtins.__import__

    def _fake_import(
        name: str,
        globals_dict: dict[str, object] | None,
        locals_dict: dict[str, object] | None,
        fromlist: tuple[str, ...],
        level: int,
    ) -> object:
        if name == "pypdfium2":
            raise ImportError("missing")
        return real_import(name, globals_dict, locals_dict, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _fake_import)

    with pytest.raises(MissingDependencyError, match="pypdfium2"):
        render._require_pdfium()


def test_export_sheet_images_success(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_sheet_images renders images with sanitized names."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    out_dir = tmp_path / "images"
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "0")

    fake_pdfium = SimpleNamespace(PdfDocument=FakePdfDocument)
    monkeypatch.setattr(render, "_require_pdfium", lambda: fake_pdfium)
    fake_app = FakeApp(["Sheet/1", "  "], False)
    monkeypatch.setattr(render, "_require_excel_app", lambda: fake_app)

    written = render.export_sheet_images(xlsx, out_dir, dpi=144)

    assert written[0].name == "01_Sheet_1.png"
    assert written[1].name == "02_Sheet_1.png"
    assert written[2].name == "03_sheet.png"
    assert all(path.exists() for path in written)
    assert fake_app.display_alerts is False


def test_export_sheet_images_propagates_render_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_sheet_images re-raises RenderError from _require_excel_app."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    out_dir = tmp_path / "images"
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "0")

    fake_pdfium = SimpleNamespace(PdfDocument=FakePdfDocument)
    monkeypatch.setattr(render, "_require_pdfium", lambda: fake_pdfium)
    monkeypatch.setattr(
        render, "_require_excel_app", lambda: (_ for _ in ()).throw(RenderError("boom"))
    )

    with pytest.raises(RenderError, match="boom"):
        render.export_sheet_images(xlsx, out_dir)


def test_export_sheet_images_wraps_unknown_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_sheet_images wraps unexpected errors as RenderError."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    out_dir = tmp_path / "images"
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "0")

    fake_pdfium = SimpleNamespace(PdfDocument=ExplodingPdfDocument)
    monkeypatch.setattr(render, "_require_pdfium", lambda: fake_pdfium)
    monkeypatch.setattr(
        render, "_require_excel_app", lambda: FakeApp(["Sheet1"], False)
    )

    with pytest.raises(RenderError, match="Failed to export sheet images"):
        render.export_sheet_images(xlsx, out_dir)


def test_export_sheet_images_uses_subprocess_when_enabled(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_sheet_images delegates to subprocess rendering when enabled."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    out_dir = tmp_path / "images"

    calls: list[tuple[Path, Path, int, str, int]] = []

    def _fake_subprocess(
        pdf_path: Path,
        output_dir: Path,
        sheet_index: int,
        safe_name: str,
        dpi: int,
    ) -> list[Path]:
        calls.append((pdf_path, output_dir, sheet_index, safe_name, dpi))
        return [output_dir / f"{sheet_index + 1:02d}_{safe_name}.png"]

    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "1")
    monkeypatch.setattr(
        render, "_require_excel_app", lambda: FakeApp(["SheetA", "SheetB"], False)
    )
    monkeypatch.setattr(render, "_require_pdfium", lambda: SimpleNamespace())
    monkeypatch.setattr(render, "_render_pdf_pages_subprocess", _fake_subprocess)

    written = render.export_sheet_images(xlsx, out_dir, dpi=144)

    assert len(calls) == 2
    assert written[0].name == "01_SheetA.png"
    assert written[1].name == "02_SheetB.png"


def test_export_sheet_images_with_sheet_only(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Export only the target sheet when `sheet` is specified."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    out_dir = tmp_path / "images"
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "0")

    fake_pdfium = SimpleNamespace(PdfDocument=FakePdfDocument)
    monkeypatch.setattr(render, "_require_pdfium", lambda: fake_pdfium)
    monkeypatch.setattr(
        render, "_require_excel_app", lambda: FakeApp(["SheetA", "SheetB"], False)
    )

    written = render.export_sheet_images(xlsx, out_dir, dpi=144, sheet="SheetB")

    assert len(written) >= 1
    assert all("SheetB" in path.name for path in written)
    assert all("SheetA" not in path.name for path in written)


def test_export_sheet_images_with_sheet_and_range(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Forward normalized range to plan builder when range is specified."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    out_dir = tmp_path / "images"
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "0")

    fake_pdfium = SimpleNamespace(PdfDocument=FakePdfDocument)
    monkeypatch.setattr(render, "_require_pdfium", lambda: fake_pdfium)
    monkeypatch.setattr(
        render, "_require_excel_app", lambda: FakeApp(["SheetA", "SheetB"], False)
    )

    captured: dict[str, object] = {}
    original = render._build_sheet_export_plan

    def _capturing_plan(
        wb: xw.Book, *, sheet: str | None = None, a1_range: str | None = None
    ) -> list[tuple[str, render._SheetApiProtocol, str | None]]:
        captured["sheet"] = sheet
        captured["a1_range"] = a1_range
        return original(wb, sheet=sheet, a1_range=a1_range)

    monkeypatch.setattr(render, "_build_sheet_export_plan", _capturing_plan)
    written = render.export_sheet_images(
        xlsx,
        out_dir,
        dpi=144,
        sheet="SheetA",
        a1_range="a1:b2",
    )

    assert captured == {"sheet": "SheetA", "a1_range": "A1:B2"}
    assert len(written) >= 1
    assert written[0].name.startswith("01_SheetA")


def test_export_sheet_images_rejects_invalid_range(tmp_path: Path) -> None:
    """Reject invalid A1 range before render execution."""
    with pytest.raises(ValueError, match="Invalid range reference"):
        render.export_sheet_images(
            tmp_path / "input.xlsx",
            tmp_path / "images",
            sheet="Sheet1",
            a1_range="A1",
        )


def test_use_render_subprocess_env_toggle(monkeypatch: pytest.MonkeyPatch) -> None:
    """_use_render_subprocess respects the env toggle."""
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "1")
    assert render._use_render_subprocess() is True
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "0")
    assert render._use_render_subprocess() is False


class FakeWorkerProcess:
    """Stub subprocess.Popen-like process for worker tests."""

    def __init__(
        self,
        *,
        returncode: int | None = None,
        stderr: str = "",
        wait_timeout: bool = False,
    ) -> None:
        self.returncode = returncode
        self.stderr = stderr
        self.wait_timeout = wait_timeout
        self.wait_calls: list[float | None] = []
        self.terminate_called = False
        self.kill_called = False

    def poll(self) -> int | None:
        """Return current worker exit code.

        Returns:
            Worker exit code when finished, otherwise None.
        """
        return self.returncode

    def wait(self, timeout: float | None = None) -> int:
        """Wait for worker completion and return exit code.

        Args:
            timeout: Optional maximum wait time in seconds.

        Returns:
            Worker exit code.

        Raises:
            subprocess.TimeoutExpired: Raised when `wait_timeout` is enabled.
        """
        self.wait_calls.append(timeout)
        if self.wait_timeout:
            raise subprocess.TimeoutExpired(cmd="worker", timeout=timeout or 0.0)
        if self.returncode is None:
            self.returncode = 0
        return self.returncode

    def terminate(self) -> None:
        """Simulate graceful process termination.

        Returns:
            None.
        """
        self.terminate_called = True
        self.returncode = 0

    def kill(self) -> None:
        """Simulate forced process termination.

        Returns:
            None.
        """
        self.kill_called = True
        self.returncode = -9

    def communicate(
        self,
        timeout: float | None = None,
    ) -> tuple[str | None, str | None]:
        """Return stubbed worker stdio.

        Args:
            timeout: Optional communication timeout in seconds.

        Returns:
            Tuple of `(stdout, stderr)` strings.
        """
        _ = timeout
        return ("", self.stderr)


def test_render_pdf_pages_subprocess_success(tmp_path: Path) -> None:
    """_render_pdf_pages_subprocess returns paths when worker succeeds."""
    output_dir = tmp_path / "images"

    def _fake_runner(
        pdf_path: Path,
        output_dir_arg: Path,
        sheet_index: int,
        safe_name: str,
        dpi: int,
        *,
        startup_timeout_seconds: float,
        result_timeout_seconds: float,
        join_timeout_seconds: float,
    ) -> render._RenderWorkerResult:
        _ = pdf_path
        _ = output_dir_arg
        _ = sheet_index
        _ = safe_name
        _ = dpi
        _ = startup_timeout_seconds
        _ = result_timeout_seconds
        _ = join_timeout_seconds
        return render._RenderWorkerResult.success([str(output_dir / "01_Sheet1.png")])

    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")

    with pytest.MonkeyPatch.context() as monkeypatch:
        monkeypatch.setattr(render, "_run_render_worker_subprocess", _fake_runner)
        result = render._render_pdf_pages_subprocess(
            pdf_path, output_dir, 0, "Sheet1", 144
        )

    assert result == [output_dir / "01_Sheet1.png"]


def test_render_pdf_pages_subprocess_emits_stage_logs(
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """_render_pdf_pages_subprocess emits start/done stage logs."""
    output_dir = tmp_path / "images"

    def _fake_runner(
        pdf_path: Path,
        output_dir_arg: Path,
        sheet_index: int,
        safe_name: str,
        dpi: int,
        *,
        startup_timeout_seconds: float,
        result_timeout_seconds: float,
        join_timeout_seconds: float,
    ) -> render._RenderWorkerResult:
        _ = pdf_path
        _ = output_dir_arg
        _ = sheet_index
        _ = safe_name
        _ = dpi
        _ = startup_timeout_seconds
        _ = result_timeout_seconds
        _ = join_timeout_seconds
        return render._RenderWorkerResult.success([str(output_dir / "01_Sheet1.png")])

    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")

    with pytest.MonkeyPatch.context() as monkeypatch:
        monkeypatch.setattr(render, "_run_render_worker_subprocess", _fake_runner)
        caplog.set_level(logging.INFO, logger=render.__name__)
        render._render_pdf_pages_subprocess(pdf_path, output_dir, 0, "Sheet1", 144)

    log_text = caplog.text
    assert "render-stage=subprocess.start" in log_text
    assert "render-stage=subprocess.done" in log_text


def test_render_pdf_pages_subprocess_error(tmp_path: Path) -> None:
    """_render_pdf_pages_subprocess raises when worker reports error."""

    def _fake_runner(
        pdf_path: Path,
        output_dir: Path,
        sheet_index: int,
        safe_name: str,
        dpi: int,
        *,
        startup_timeout_seconds: float,
        result_timeout_seconds: float,
        join_timeout_seconds: float,
    ) -> render._RenderWorkerResult:
        _ = pdf_path
        _ = output_dir
        _ = sheet_index
        _ = safe_name
        _ = dpi
        _ = startup_timeout_seconds
        _ = result_timeout_seconds
        _ = join_timeout_seconds
        return render._RenderWorkerResult.failure("boom")

    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"

    with pytest.MonkeyPatch.context() as monkeypatch:
        monkeypatch.setattr(render, "_run_render_worker_subprocess", _fake_runner)
        with pytest.raises(RenderError, match="stage=worker boom"):
            render._render_pdf_pages_subprocess(pdf_path, output_dir, 0, "Sheet1", 144)


def test_run_render_worker_subprocess_success_when_join_timeout(
    tmp_path: Path,
) -> None:
    """Return success when result is received even if join wait times out."""
    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"
    process = FakeWorkerProcess(returncode=None, wait_timeout=True)

    with pytest.MonkeyPatch.context() as monkeypatch:
        monkeypatch.setattr(render, "_start_render_worker_process", lambda _: process)
        monkeypatch.setattr(
            render,
            "_wait_for_worker_startup",
            lambda proc, *, started_path, timeout_seconds: None,
        )
        monkeypatch.setattr(
            render,
            "_wait_for_worker_result",
            lambda proc,
            *,
            result_path,
            join_timeout_deadline,
            join_timeout_seconds,
            post_exit_timeout_seconds: render._RenderWorkerResult.success(
                [str(output_dir / "01_Sheet1.png")]
            ),
        )
        result = render._run_render_worker_subprocess(
            pdf_path,
            output_dir,
            0,
            "Sheet1",
            144,
            startup_timeout_seconds=1.0,
            result_timeout_seconds=1.0,
            join_timeout_seconds=0.1,
        )

    assert result.paths == [str(output_dir / "01_Sheet1.png")]
    assert process.terminate_called is True


def test_run_render_worker_subprocess_starts_join_budget_after_startup(
    tmp_path: Path,
) -> None:
    """Start join timeout budget after startup wait has completed."""
    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"
    process = FakeWorkerProcess(returncode=0)
    captured: dict[str, float] = {}

    with pytest.MonkeyPatch.context() as monkeypatch:
        timestamps = iter([100.0, 200.0, 200.02])
        monkeypatch.setattr(
            "exstruct.render.time.perf_counter",
            lambda: next(timestamps, 200.02),
        )
        monkeypatch.setattr(render, "_start_render_worker_process", lambda _: process)

        def _fake_startup(
            proc: render._WorkerProcessProtocol,
            *,
            started_path: Path,
            timeout_seconds: float,
        ) -> None:
            _ = proc
            _ = started_path
            _ = timeout_seconds
            _ = time.perf_counter()

        monkeypatch.setattr(render, "_wait_for_worker_startup", _fake_startup)

        def _capture_deadline(
            proc: render._WorkerProcessProtocol,
            *,
            result_path: Path,
            join_timeout_deadline: float,
            join_timeout_seconds: float,
            post_exit_timeout_seconds: float,
        ) -> render._RenderWorkerResult:
            _ = proc
            _ = result_path
            _ = join_timeout_seconds
            _ = post_exit_timeout_seconds
            captured["deadline"] = join_timeout_deadline
            return render._RenderWorkerResult.success(
                [str(output_dir / "01_Sheet1.png")]
            )

        monkeypatch.setattr(render, "_wait_for_worker_result", _capture_deadline)
        monkeypatch.setattr(
            render,
            "_wait_for_worker_join",
            lambda proc, *, timeout_seconds: True,
        )

        render._run_render_worker_subprocess(
            pdf_path,
            output_dir,
            0,
            "Sheet1",
            144,
            startup_timeout_seconds=5.0,
            result_timeout_seconds=1.0,
            join_timeout_seconds=0.1,
        )

    assert captured["deadline"] == pytest.approx(200.1, abs=1e-6)


def test_run_render_worker_subprocess_uses_single_join_budget(
    tmp_path: Path,
) -> None:
    """Use remaining join budget after result wait instead of full timeout twice."""
    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"
    process = FakeWorkerProcess(returncode=0)
    captured: dict[str, float] = {}

    with pytest.MonkeyPatch.context() as monkeypatch:
        timestamps = iter([100.0, 100.08])
        monkeypatch.setattr(
            "exstruct.render.time.perf_counter",
            lambda: next(timestamps, 100.08),
        )
        monkeypatch.setattr(render, "_start_render_worker_process", lambda _: process)
        monkeypatch.setattr(
            render,
            "_wait_for_worker_startup",
            lambda proc, *, started_path, timeout_seconds: None,
        )
        monkeypatch.setattr(
            render,
            "_wait_for_worker_result",
            lambda proc,
            *,
            result_path,
            join_timeout_deadline,
            join_timeout_seconds,
            post_exit_timeout_seconds: render._RenderWorkerResult.success(
                [str(output_dir / "01_Sheet1.png")]
            ),
        )

        def _capture_join_timeout(
            proc: render._WorkerProcessProtocol, *, timeout_seconds: float
        ) -> bool:
            _ = proc
            captured["timeout"] = timeout_seconds
            return True

        monkeypatch.setattr(render, "_wait_for_worker_join", _capture_join_timeout)

        render._run_render_worker_subprocess(
            pdf_path,
            output_dir,
            0,
            "Sheet1",
            144,
            startup_timeout_seconds=1.0,
            result_timeout_seconds=1.0,
            join_timeout_seconds=0.1,
        )

    assert captured["timeout"] == pytest.approx(0.02, abs=1e-6)


def test_run_render_worker_subprocess_startup_error_is_actionable(
    tmp_path: Path,
) -> None:
    """Emit startup-stage error with stderr snippet on bootstrap failure."""
    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"
    process = FakeWorkerProcess(returncode=1, stderr="bootstrap failed")

    with pytest.MonkeyPatch.context() as monkeypatch:
        monkeypatch.setattr(render, "_start_render_worker_process", lambda _: process)
        with pytest.raises(RenderError, match="stage=startup"):
            render._run_render_worker_subprocess(
                pdf_path,
                output_dir,
                0,
                "Sheet1",
                144,
                startup_timeout_seconds=0.1,
                result_timeout_seconds=0.1,
                join_timeout_seconds=0.1,
            )


def test_wait_for_worker_result_timeout_has_join_stage(tmp_path: Path) -> None:
    """Timeout while worker is still running should report join stage."""
    process = FakeWorkerProcess(returncode=None)
    result_path = tmp_path / "result.json"

    with pytest.raises(RenderError, match="stage=join timed out"):
        render._wait_for_worker_result(
            process,
            result_path=result_path,
            join_timeout_deadline=time.perf_counter() + 0.01,
            join_timeout_seconds=0.01,
            post_exit_timeout_seconds=0.01,
        )


def test_wait_for_worker_result_allows_longer_than_post_exit_timeout(
    tmp_path: Path,
) -> None:
    """Do not fail early while worker is alive even when post-exit timeout is short."""
    process = FakeWorkerProcess(returncode=None)
    result_path = tmp_path / "result.json"
    start_writing = threading.Event()
    writer_released = threading.Event()
    real_sleep = time.sleep

    def _write_later() -> None:
        if not start_writing.wait(1.0):
            return
        result_path.write_text(
            json.dumps({"paths": [str(tmp_path / "images" / "01_Sheet1.png")]}),
            encoding="utf-8",
        )
        writer_released.set()

    sleep_calls = 0

    def _sleep_hook(seconds: float) -> None:
        nonlocal sleep_calls
        sleep_calls += 1
        if sleep_calls == 1:
            start_writing.set()
        real_sleep(seconds)

    thread = threading.Thread(target=_write_later)
    thread.start()
    try:
        with pytest.MonkeyPatch.context() as monkeypatch:
            monkeypatch.setattr("exstruct.render.time.sleep", _sleep_hook)
            result = render._wait_for_worker_result(
                process,
                result_path=result_path,
                join_timeout_deadline=time.perf_counter() + 0.20,
                join_timeout_seconds=0.20,
                post_exit_timeout_seconds=0.01,
            )
    finally:
        start_writing.set()
        thread.join(timeout=1.0)

    assert writer_released.is_set() is True
    assert result.paths == [str(tmp_path / "images" / "01_Sheet1.png")]


def test_wait_for_worker_result_reports_result_stage_after_exit(
    tmp_path: Path,
) -> None:
    """When worker has exited and no result arrives, report result stage."""
    process = FakeWorkerProcess(returncode=1, stderr="missing result")
    result_path = tmp_path / "result.json"

    with pytest.raises(RenderError, match="stage=result worker exited without result"):
        render._wait_for_worker_result(
            process,
            result_path=result_path,
            join_timeout_deadline=time.perf_counter() + 0.20,
            join_timeout_seconds=0.20,
            post_exit_timeout_seconds=0.01,
        )


def test_read_worker_result_prioritizes_error_field(tmp_path: Path) -> None:
    """Treat payload as failure when `error` is present.

    Args:
        tmp_path: Temporary directory fixture.

    Returns:
        None.
    """
    result_path = tmp_path / "result.json"
    result_path.write_text(
        json.dumps({"paths": [], "error": "worker failed"}),
        encoding="utf-8",
    )

    result = render._read_worker_result(result_path)

    assert result.error == "worker failed"
    assert result.paths == []


def test_get_render_subprocess_timeout_seconds(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Subprocess timeout readers validate env values."""
    monkeypatch.delenv("EXSTRUCT_RENDER_SUBPROCESS_STARTUP_TIMEOUT_SEC", raising=False)
    monkeypatch.delenv("EXSTRUCT_RENDER_SUBPROCESS_JOIN_TIMEOUT_SEC", raising=False)
    monkeypatch.delenv("EXSTRUCT_RENDER_SUBPROCESS_RESULT_TIMEOUT_SEC", raising=False)
    assert render._get_render_subprocess_startup_timeout_seconds() == 5.0
    assert render._get_render_subprocess_join_timeout_seconds() == 120.0
    assert render._get_render_subprocess_result_timeout_seconds() == 5.0

    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_STARTUP_TIMEOUT_SEC", "3")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_JOIN_TIMEOUT_SEC", "30")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_RESULT_TIMEOUT_SEC", "7")
    assert render._get_render_subprocess_startup_timeout_seconds() == 3.0
    assert render._get_render_subprocess_join_timeout_seconds() == 30.0
    assert render._get_render_subprocess_result_timeout_seconds() == 7.0

    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_STARTUP_TIMEOUT_SEC", "0")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_JOIN_TIMEOUT_SEC", "0")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_RESULT_TIMEOUT_SEC", "0")
    assert render._get_render_subprocess_startup_timeout_seconds() == 5.0
    assert render._get_render_subprocess_join_timeout_seconds() == 120.0
    assert render._get_render_subprocess_result_timeout_seconds() == 5.0

    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_STARTUP_TIMEOUT_SEC", "NaN")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_JOIN_TIMEOUT_SEC", "NaN")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_RESULT_TIMEOUT_SEC", "NaN")
    assert render._get_render_subprocess_startup_timeout_seconds() == 5.0
    assert render._get_render_subprocess_join_timeout_seconds() == 120.0
    assert render._get_render_subprocess_result_timeout_seconds() == 5.0

    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_STARTUP_TIMEOUT_SEC", "inf")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_JOIN_TIMEOUT_SEC", "inf")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_RESULT_TIMEOUT_SEC", "inf")
    assert render._get_render_subprocess_startup_timeout_seconds() == 5.0
    assert render._get_render_subprocess_join_timeout_seconds() == 120.0
    assert render._get_render_subprocess_result_timeout_seconds() == 5.0

    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_STARTUP_TIMEOUT_SEC", "-inf")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_JOIN_TIMEOUT_SEC", "-inf")
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS_RESULT_TIMEOUT_SEC", "-inf")
    assert render._get_render_subprocess_startup_timeout_seconds() == 5.0
    assert render._get_render_subprocess_join_timeout_seconds() == 120.0
    assert render._get_render_subprocess_result_timeout_seconds() == 5.0


def test_sanitize_sheet_filename() -> None:
    """_sanitize_sheet_filename replaces invalid characters and defaults."""
    assert render._sanitize_sheet_filename("Sheet/1") == "Sheet_1"
    assert render._sanitize_sheet_filename("  ") == "sheet"


def test_split_csv_respecting_quotes() -> None:
    """Split CSV-like PrintArea strings while honoring quotes."""
    raw = "'Sheet 1'!A1:B2,'Sheet,2'!C3:D4,'O''Brien'!E1:F2"
    parts = render._split_csv_respecting_quotes(raw)
    assert parts == ["'Sheet 1'!A1:B2", "'Sheet,2'!C3:D4", "'O''Brien'!E1:F2"]


def test_extract_print_areas_with_page_setup() -> None:
    """Parse PrintArea from a PageSetup stub."""

    class _PageSetup:
        PrintArea = "'Sheet 1'!A1:B2,'Sheet 1'!C3:D4"

    class _SheetApi:
        PageSetup = _PageSetup()

    areas = render._extract_print_areas(cast(render._SheetApiProtocol, _SheetApi()))
    assert areas == ["'Sheet 1'!A1:B2", "'Sheet 1'!C3:D4"]


def test_extract_print_areas_empty_print_area() -> None:
    """Return empty list when PrintArea is empty."""

    class _PageSetup:
        PrintArea = ""

    class _SheetApi:
        PageSetup = _PageSetup()

    assert (
        render._extract_print_areas(cast(render._SheetApiProtocol, _SheetApi())) == []
    )


def test_extract_print_areas_handles_exception() -> None:
    """Return empty list when PrintArea access raises."""

    class _PageSetup:
        @property
        def PrintArea(self) -> str:
            """
            Simulate accessing a worksheet's PrintArea and always raise an error to emulate a failure.

            Raises:
                RuntimeError: Always raised to simulate an error when retrieving the PrintArea.
            """
            raise RuntimeError("boom")

    class _SheetApi:
        PageSetup = _PageSetup()

    assert (
        render._extract_print_areas(cast(render._SheetApiProtocol, _SheetApi())) == []
    )


def test_iter_sheet_apis_prefers_worksheets_collection() -> None:
    """Prefer the Worksheets collection when iterating COM sheets."""

    class _WsApi:
        def __init__(self, name: str) -> None:
            """
            Initialize the FakeSheet with the given Excel sheet name.

            Parameters:
                name (str): The sheet's name to assign to the object's `Name` attribute.
            """
            self.Name = name

    class _Worksheets:
        def __init__(self) -> None:
            """
            Initialize the fake PDF document stub.

            Sets the `Count` attribute to 2 to emulate a document with two pages.
            """
            self.Count = 2

        def Item(self, index: int) -> _WsApi:
            """
            Return a worksheet API stub for the sheet at the given index.

            Parameters:
                index (int): One-based index of the worksheet within the workbook.

            Returns:
                _WsApi: A worksheet API stub corresponding to the sheet at `index`.
            """
            return _WsApi(f"Sheet{index}")

    class _Api:
        Worksheets = _Worksheets()

    class _Wb:
        api = _Api()
        sheets: list[Any] = []

    result = render._iter_sheet_apis(_Wb())
    assert result[0][1] == "Sheet1"
    assert result[1][1] == "Sheet2"


def test_export_pdf_propagates_render_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    def _raise() -> xw.App:
        """
        Always raises a RenderError to simulate failure when obtaining an Excel application.

        Raises:
            RenderError: Always raised with the message "boom".
        """
        raise RenderError("boom")

    monkeypatch.setattr(render, "_require_excel_app", _raise)
    with pytest.raises(RenderError, match="boom"):
        render.export_pdf(tmp_path / "in.xlsx", tmp_path / "out.pdf")


def test_require_pdfium_success(monkeypatch: pytest.MonkeyPatch) -> None:
    """_require_pdfium returns the imported module when available."""
    fake_pdfium = ModuleType("pypdfium2")
    sys.modules["pypdfium2"] = fake_pdfium
    try:
        assert render._require_pdfium() is fake_pdfium
    finally:
        sys.modules.pop("pypdfium2", None)


def test_build_sheet_export_plan_handles_multiple_areas(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Expand multiple print areas into separate export plan rows."""

    class _SheetApi:
        pass

    def _fake_iter(_: xw.Book) -> list[tuple[int, str, _SheetApi]]:
        """
        Return a single-item list that mimics iterating workbook sheets for tests.

        Returns:
            A list with one tuple (index, sheet name, sheet API stub): (0, "Sheet1", _SheetApi()).
        """
        return [(0, "Sheet1", _SheetApi())]

    def _fake_extract(_: _SheetApi) -> list[str]:
        """
        Provide two fake print-area ranges for testing.

        Parameters:
            _ (_SheetApi): Ignored sheet API placeholder.

        Returns:
            list[str]: Two print-area ranges: "A1:B2" and "C3:D4".
        """
        return ["A1:B2", "C3:D4"]

    monkeypatch.setattr(render, "_iter_sheet_apis", _fake_iter)
    monkeypatch.setattr(render, "_extract_print_areas", _fake_extract)

    plan = render._build_sheet_export_plan(cast(xw.Book, object()))
    assert [item[0] for item in plan] == ["Sheet1", "Sheet1"]
    assert [item[2] for item in plan] == ["A1:B2", "C3:D4"]


def test_page_index_from_suffix_default() -> None:
    """Default to zero when no suffix exists."""
    assert render._page_index_from_suffix("sheet") == 0


def test_page_index_from_suffix_non_digit() -> None:
    """Default to zero when suffix is not numeric."""
    assert render._page_index_from_suffix("sheet_pxx") == 0


def test_export_sheet_pdf_skips_invalid_print_area(tmp_path: Path) -> None:
    """Skip restoring PrintArea when setter fails."""

    class _BadPageSetup:
        @property
        def PrintArea(self) -> str:
            """
            Represents the worksheet's PrintArea setting as an Excel range string.

            Returns:
                str: The PrintArea range (e.g., "A1:B2").
            """
            return "A1:B2"

        @PrintArea.setter
        def PrintArea(self, _value: object) -> None:
            """
            Simulated setter for PrintArea that always fails.

            Parameters:
                _value (object): Ignored; the provided value is not used because the setter always raises.

            Raises:
                RuntimeError: Always raised with the message "bad".
            """
            raise RuntimeError("bad")

    class _SheetApi:
        PageSetup = _BadPageSetup()

        def ExportAsFixedFormat(
            self, _file_format: int, _output_path: str, *args: object, **kwargs: object
        ) -> None:
            """
            Simulate exporting a workbook/sheet to a fixed-format file by writing a minimal fake PDF header to the given path.

            Parameters:
                _file_format (int): Ignored numeric format indicator.
                _output_path (str): Filesystem path where the fake export file will be written.
                *args, **kwargs: Additional arguments are accepted and ignored.
            """
            _ = args
            _ = kwargs

    render._export_sheet_pdf(
        cast(render._SheetApiProtocol, _SheetApi()),
        tmp_path / "out.pdf",
        ignore_print_areas=False,
        print_area="A1:B2",
    )


def test_render_sheet_images_requires_pdfium(tmp_path: Path) -> None:
    """Raise RenderError when pdfium is missing."""
    with pytest.raises(RenderError, match="pypdfium2 is required"):
        render._render_sheet_images(
            None,
            tmp_path / "sheet.pdf",
            tmp_path,
            0,
            "Sheet1",
            144,
            False,
        )


def test_export_sheet_images_with_app_retries_on_empty(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Retry export when rendering returns empty results."""
    calls: list[int] = []

    def _fake_render(
        _pdfium: ModuleType | None,
        _pdf_path: Path,
        output_dir: Path,
        sheet_index: int,
        safe_name: str,
        _dpi: int,
        _use_subprocess: bool,
    ) -> list[Path]:
        """
        Simulates rendering a PDF sheet to image files for tests.

        On the first invocation this function returns an empty list to simulate a transient empty render result; on subsequent invocations it returns a single Path inside output_dir named "{sheet_index+1:02d}_{safe_name}.png".

        Parameters:
            _pdfium: Ignored in the fake implementation (kept for signature compatibility).
            _pdf_path: Ignored in the fake implementation (kept for signature compatibility).
            output_dir (Path): Directory where the fake image path is located.
            sheet_index (int): Zero-based index of the sheet; used to build the filename prefix.
            safe_name (str): Sanitized sheet name used in the filename.
            _dpi: Ignored in the fake implementation (kept for signature compatibility).
            _use_subprocess: Ignored in the fake implementation (kept for signature compatibility).

        Returns:
            list[Path]: Empty list on the first call, otherwise a list containing one Path pointing to the fake PNG file.
        """
        calls.append(1)
        if len(calls) == 1:
            return []
        return [output_dir / f"{sheet_index + 1:02d}_{safe_name}.png"]

    monkeypatch.setattr(render, "_render_sheet_images", _fake_render)
    monkeypatch.setattr(
        render, "_require_excel_app", lambda: FakeApp(["Sheet1"], False)
    )
    monkeypatch.setattr(render, "_export_sheet_pdf", lambda *a, **k: None)
    monkeypatch.setattr(
        render,
        "_build_sheet_export_plan",
        lambda _wb, *, sheet=None, a1_range=None: [
            ("Sheet1", cast(render._SheetApiProtocol, object()), None)
        ],
    )

    result = render._export_sheet_images_with_app(
        tmp_path / "in.xlsx",
        tmp_path / "out",
        tmp_path / "tmp",
        144,
        False,
        None,
        None,
        None,
    )
    assert len(calls) == 2
    assert result


def test_export_sheet_images_with_app_skips_retry_for_targeted_range(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Do not run print-area fallback for targeted range exports.

    Args:
        tmp_path: Temporary directory fixture.
        monkeypatch: Pytest monkeypatch fixture.

    Returns:
        None.
    """
    render_calls: list[int] = []
    export_calls: list[bool] = []

    def _fake_render(
        _pdfium: ModuleType | None,
        _pdf_path: Path,
        _output_dir: Path,
        _sheet_index: int,
        _safe_name: str,
        _dpi: int,
        _use_subprocess: bool,
    ) -> list[Path]:
        render_calls.append(1)
        return []

    def _fake_export(
        _sheet_api: render._SheetApiProtocol,
        _pdf_path: Path,
        *,
        ignore_print_areas: bool,
        print_area: str | None = None,
    ) -> None:
        _ = print_area
        export_calls.append(ignore_print_areas)

    monkeypatch.setattr(render, "_render_sheet_images", _fake_render)
    monkeypatch.setattr(
        render, "_require_excel_app", lambda: FakeApp(["Sheet1"], False)
    )
    monkeypatch.setattr(render, "_export_sheet_pdf", _fake_export)
    monkeypatch.setattr(
        render,
        "_build_sheet_export_plan",
        lambda _wb, *, sheet=None, a1_range=None: [
            ("Sheet1", cast(render._SheetApiProtocol, object()), None)
        ],
    )

    result = render._export_sheet_images_with_app(
        tmp_path / "in.xlsx",
        tmp_path / "out",
        tmp_path / "tmp",
        144,
        False,
        None,
        "Sheet1",
        "A1:B2",
    )

    assert result == []
    assert len(render_calls) == 1
    assert export_calls == [False]


def test_page_index_from_suffix_handles_multi_digits() -> None:
    """Support multi-digit page suffixes."""
    assert render._page_index_from_suffix("sheet_01") == 0
    assert render._page_index_from_suffix("sheet_01_p01") == 0
    assert render._page_index_from_suffix("sheet_01_p10") == 9
    assert render._page_index_from_suffix("sheet_01_p100") == 99
    assert render._page_index_from_suffix("sheet_01_p0") == 0


def test_export_sheet_pdf_does_not_swallow_export_errors(tmp_path: Path) -> None:
    """Propagate export errors even if restore fails."""

    class _FlakyPageSetup:
        def __init__(self) -> None:
            """
            Initialize a PageSetup-like test stub with a default print area and a setter call counter.

            The instance starts with `_print_area` set to "A1" and `_set_calls` set to 0 to track how many times the print area setter has been invoked.
            """
            self._print_area: object = "A1"
            self._set_calls = 0

        @property
        def PrintArea(self) -> object:
            """
            Retrieve the current PrintArea value from the PageSetup stub.

            Returns:
                print_area (object): The stored PrintArea value (typically a string) or whatever was set on the stub.
            """
            return self._print_area

        @PrintArea.setter
        def PrintArea(self, value: object) -> None:
            """
            Set the PrintArea value on this stub PageSetup instance.

            Parameters:
                value (object): The print area value to assign.

            Raises:
                RuntimeError: If the setter is invoked more than once (simulates a restore failure).
            """
            if self._set_calls >= 1:
                raise RuntimeError("restore failed")
            self._print_area = value
            self._set_calls += 1

    class _ExplodingSheetApi:
        PageSetup: render._PageSetupProtocol = cast(
            render._PageSetupProtocol, _FlakyPageSetup()
        )

        def ExportAsFixedFormat(
            self, file_format: int, output_path: str, *args: object, **kwargs: object
        ) -> None:
            """
            Simulate exporting to a fixed format; this stub always raises an export error.

            Raises:
                RuntimeError: with message "export failed" when invoked.
            """
            _ = file_format
            _ = output_path
            _ = args
            _ = kwargs
            raise RuntimeError("export failed")

    pdf_path = tmp_path / "out.pdf"
    with pytest.raises(RuntimeError, match="export failed"):
        render._export_sheet_pdf(
            _ExplodingSheetApi(),
            pdf_path,
            ignore_print_areas=False,
            print_area="A1:B2",
        )
