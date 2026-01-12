from __future__ import annotations

import builtins
from collections.abc import Callable
from pathlib import Path
import shutil
import sys
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
    monkeypatch.setattr(
        render, "_require_excel_app", lambda: FakeApp(["Sheet/1", "  "], False)
    )

    written = render.export_sheet_images(xlsx, out_dir, dpi=144)

    assert written[0].name == "01_Sheet_1.png"
    assert written[1].name == "01_Sheet_1_p02.png"
    assert written[2].name == "02_sheet.png"
    assert all(path.exists() for path in written)


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


def test_use_render_subprocess_env_toggle(monkeypatch: pytest.MonkeyPatch) -> None:
    """_use_render_subprocess respects the env toggle."""
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "1")
    assert render._use_render_subprocess() is True
    monkeypatch.setenv("EXSTRUCT_RENDER_SUBPROCESS", "0")
    assert render._use_render_subprocess() is False


class FakeQueue:
    """Stub queue for subprocess tests."""

    def __init__(self) -> None:
        self.payload: dict[str, list[str] | str] | None = None

    def put(self, payload: dict[str, list[str] | str]) -> None:
        self.payload = payload

    def get(self, timeout: float | None = None) -> dict[str, list[str] | str]:
        _ = timeout
        if self.payload is None:
            raise TimeoutError("timeout")
        return self.payload

    def empty(self) -> bool:
        return self.payload is None


class FakeProcess:
    """Stub process for subprocess tests."""

    def __init__(
        self,
        queue: FakeQueue,
        exitcode: int,
        payload: dict[str, list[str] | str] | None = None,
    ) -> None:
        self._queue = queue
        self.exitcode = exitcode
        if payload is not None:
            self._queue.put(payload)

    def start(self) -> None:
        if self._queue.payload is None:
            self._queue.put({"paths": ["dummy"]})

    def join(self) -> None:
        return None


class FakeContext:
    """Stub multiprocessing context for subprocess tests."""

    def __init__(self, queue: FakeQueue, process: FakeProcess) -> None:
        self._queue = queue
        self._process = process

    def Queue(self) -> FakeQueue:
        return self._queue

    def Process(self, target: object, args: tuple[object, ...]) -> FakeProcess:
        _ = target
        _ = args
        return self._process


def test_render_pdf_pages_subprocess_success(tmp_path: Path) -> None:
    """_render_pdf_pages_subprocess returns paths when worker succeeds."""
    queue = FakeQueue()
    process = FakeProcess(
        queue,
        exitcode=0,
        payload={"paths": [str(tmp_path / "images" / "01_Sheet1.png")]},
    )
    context = FakeContext(queue, process)
    render_mp = cast(Any, render).mp

    def _get_context(_: str) -> FakeContext:
        return context

    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"

    with pytest.MonkeyPatch.context() as monkeypatch:
        monkeypatch.setattr(render_mp, "get_context", _get_context)
        result = render._render_pdf_pages_subprocess(
            pdf_path, output_dir, 0, "Sheet1", 144
        )

    assert result == [output_dir / "01_Sheet1.png"]


def test_render_pdf_pages_subprocess_error(tmp_path: Path) -> None:
    """_render_pdf_pages_subprocess raises when worker reports error."""
    queue = FakeQueue()
    process = FakeProcess(queue, exitcode=0, payload={"error": "boom"})
    context = FakeContext(queue, process)
    render_mp = cast(Any, render).mp

    def _get_context(_: str) -> FakeContext:
        return context

    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"

    with pytest.MonkeyPatch.context() as monkeypatch:
        monkeypatch.setattr(render_mp, "get_context", _get_context)
        with pytest.raises(RenderError, match="boom"):
            render._render_pdf_pages_subprocess(pdf_path, output_dir, 0, "Sheet1", 144)


def test_get_subprocess_result_timeout() -> None:
    """_get_subprocess_result returns an error payload on timeout."""
    queue = FakeQueue()
    result = render._get_subprocess_result(cast(Any, queue))

    error = cast(str, result["error"])
    assert error.startswith("subprocess did not return results")


def test_render_pdf_pages_worker_success(tmp_path: Path) -> None:
    """_render_pdf_pages_worker writes images and returns paths."""
    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"
    queue = FakeQueue()
    fake_pdfium = cast(Any, ModuleType("pypdfium2"))
    fake_pdfium.PdfDocument = FakePdfDocument

    sys.modules["pypdfium2"] = fake_pdfium
    try:
        render._render_pdf_pages_worker(
            pdf_path, output_dir, 0, "Sheet1", 144, cast(Any, queue)
        )
    finally:
        sys.modules.pop("pypdfium2", None)

    assert queue.payload == {
        "paths": [
            str(output_dir / "01_Sheet1.png"),
            str(output_dir / "01_Sheet1_p02.png"),
        ]
    }
    assert (output_dir / "01_Sheet1.png").exists()
    assert (output_dir / "01_Sheet1_p02.png").exists()


def test_render_pdf_pages_worker_error(tmp_path: Path) -> None:
    """_render_pdf_pages_worker reports errors via queue."""
    pdf_path = tmp_path / "sheet_01.pdf"
    pdf_path.write_bytes(b"%PDF-1.4")
    output_dir = tmp_path / "images"
    queue = FakeQueue()

    fake_pdfium = cast(Any, ModuleType("pypdfium2"))
    fake_pdfium.PdfDocument = ExplodingPdfDocument
    sys.modules["pypdfium2"] = fake_pdfium
    try:
        render._render_pdf_pages_worker(
            pdf_path, output_dir, 0, "Sheet1", 144, cast(Any, queue)
        )
    finally:
        sys.modules.pop("pypdfium2", None)

    assert queue.payload == {"error": "boom"}


def test_sanitize_sheet_filename() -> None:
    """_sanitize_sheet_filename replaces invalid characters and defaults."""
    assert render._sanitize_sheet_filename("Sheet/1") == "Sheet_1"
    assert render._sanitize_sheet_filename("  ") == "sheet"
