from __future__ import annotations

import builtins
from collections.abc import Callable
from pathlib import Path
from types import SimpleNamespace

import pytest

from exstruct.errors import MissingDependencyError, RenderError
import exstruct.render as render


class FakeSheet:
    """Minimal sheet stub with a name attribute."""

    def __init__(self, name: str) -> None:
        self.name = name


class FakeBookApi:
    """Stub of xlwings Book.api for PDF export."""

    def ExportAsFixedFormat(self, file_format: int, output_path: str) -> None:
        _ = file_format
        Path(output_path).write_bytes(b"%PDF-1.4")


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
    monkeypatch.setattr(render.xw, "App", lambda *a, **k: fake_app)

    assert render._require_excel_app() is fake_app


def test_require_excel_app_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    """_require_excel_app wraps unexpected errors as RenderError."""

    def _raise(*args: object, **kwargs: object) -> None:
        _ = args
        _ = kwargs
        raise RuntimeError("boom")

    monkeypatch.setattr(render.xw, "App", _raise)

    with pytest.raises(RenderError, match="Excel \\(COM\\) is not available"):
        render._require_excel_app()


def test_export_pdf_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """export_pdf writes a PDF and returns sheet names."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    output_pdf = tmp_path / "out.pdf"
    sheet_names = ["Sheet1", "Summary"]
    monkeypatch.setattr(render.xw, "App", _fake_app_factory(sheet_names))

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
    monkeypatch.setattr(
        render.xw, "App", _fake_app_factory(["Sheet1"], raise_on_open=True)
    )

    with pytest.raises(RenderError, match="Failed to export PDF for"):
        render.export_pdf(xlsx, output_pdf)


def test_export_pdf_missing_output_raises(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_pdf raises RenderError when the output file is missing."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    output_pdf = tmp_path / "out.pdf"
    monkeypatch.setattr(render.xw, "App", _fake_app_factory(["Sheet1"]))

    real_copy = render.shutil.copy

    def _copy(
        src: Path | str, dst: Path | str, *args: object, **kwargs: object
    ) -> Path:
        _ = args
        _ = kwargs
        if Path(dst) == output_pdf:
            return Path(dst)
        return Path(real_copy(src, dst))

    monkeypatch.setattr(render.shutil, "copy", _copy)

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

    def _fake_export_pdf(excel_path: Path, output_pdf: Path) -> list[str]:
        _ = excel_path
        output_pdf.write_bytes(b"%PDF-1.4")
        return ["Sheet/1", "  "]

    fake_pdfium = SimpleNamespace(PdfDocument=FakePdfDocument)
    monkeypatch.setattr(render, "_require_pdfium", lambda: fake_pdfium)
    monkeypatch.setattr(render, "export_pdf", _fake_export_pdf)

    written = render.export_sheet_images(xlsx, out_dir, dpi=144)

    assert written[0].name == "01_Sheet_1.png"
    assert written[1].name == "02_sheet.png"
    assert all(path.exists() for path in written)


def test_export_sheet_images_propagates_render_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_sheet_images re-raises RenderError from export_pdf."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    out_dir = tmp_path / "images"

    def _fake_export_pdf(excel_path: Path, output_pdf: Path) -> list[str]:
        _ = excel_path
        _ = output_pdf
        raise RenderError("boom")

    fake_pdfium = SimpleNamespace(PdfDocument=FakePdfDocument)
    monkeypatch.setattr(render, "_require_pdfium", lambda: fake_pdfium)
    monkeypatch.setattr(render, "export_pdf", _fake_export_pdf)

    with pytest.raises(RenderError, match="boom"):
        render.export_sheet_images(xlsx, out_dir)


def test_export_sheet_images_wraps_unknown_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """export_sheet_images wraps unexpected errors as RenderError."""
    xlsx = tmp_path / "input.xlsx"
    xlsx.write_bytes(b"dummy")
    out_dir = tmp_path / "images"

    def _fake_export_pdf(excel_path: Path, output_pdf: Path) -> list[str]:
        _ = excel_path
        output_pdf.write_bytes(b"%PDF-1.4")
        return ["Sheet1"]

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

    fake_pdfium = SimpleNamespace(PdfDocument=ExplodingPdfDocument)
    monkeypatch.setattr(render, "_require_pdfium", lambda: fake_pdfium)
    monkeypatch.setattr(render, "export_pdf", _fake_export_pdf)

    with pytest.raises(RenderError, match="Failed to export sheet images"):
        render.export_sheet_images(xlsx, out_dir)


def test_sanitize_sheet_filename() -> None:
    """_sanitize_sheet_filename replaces invalid characters and defaults."""
    assert render._sanitize_sheet_filename("Sheet/1") == "Sheet_1"
    assert render._sanitize_sheet_filename("  ") == "sheet"
