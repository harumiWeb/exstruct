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
    assert written[1].name == "02_Sheet_1.png"
    assert written[2].name == "03_sheet.png"
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
        lambda _wb: [("Sheet1", cast(render._SheetApiProtocol, object()), None)],
    )

    result = render._export_sheet_images_with_app(
        tmp_path / "in.xlsx",
        tmp_path / "out",
        tmp_path / "tmp",
        144,
        False,
        None,
    )
    assert len(calls) == 2
    assert result


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