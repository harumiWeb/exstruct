"""Tests for extraction mode behavior and output paths."""

import os
from pathlib import Path
import subprocess
import sys
from typing import Never, cast

from _pytest.capture import CaptureFixture
from _pytest.monkeypatch import MonkeyPatch
from openpyxl import Workbook
import pytest
import xlwings as xw

from exstruct import ConfigError, ExStructEngine, ExtractionMode, extract, process_excel
from exstruct.core.integrate import extract_workbook
from exstruct.core.ooxml_drawing import (
    DrawingShapeRef,
    OoxmlChartInfo,
    OoxmlShapeInfo,
    SheetDrawingData,
)
from exstruct.models import Arrow, Chart, Shape


def _make_basic_book(path: Path) -> None:
    """
    Create and save a simple Excel workbook with one sheet named "Sheet1" containing "v1" in A1 and "v2" in B1.

    Parameters:
        path (Path): Filesystem path where the workbook will be saved.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "v1"
    ws["B1"] = "v2"
    wb.save(path)


def _ensure_excel() -> None:
    """
    Ensure Excel COM is available for tests and skip the current test if it is not.

    If the SKIP_COM_TESTS environment variable is set, this function skips the test. Otherwise it tries to start a hidden xlwings App and quits it; if starting the App fails, the function skips the test due to unavailable Excel COM.
    """
    if os.getenv("SKIP_COM_TESTS"):
        pytest.skip("SKIP_COM_TESTS is set; skipping Excel-dependent test.")
    try:
        app = xw.App(add_book=False, visible=False)
        app.quit()
    except Exception:
        pytest.skip("Excel COM is unavailable; skipping Excel-dependent test.")


def _make_shapes_book(path: Path) -> None:
    """Create shapes book for tests."""

    app = xw.App(add_book=False, visible=False)
    try:
        wb = app.books.add()
        sht = wb.sheets[0]
        sht.name = "Sheet1"
        rect = sht.api.Shapes.AddShape(1, 50, 50, 120, 60)
        rect.TextFrame2.TextRange.Text = "rect"
        _ = sht.api.Shapes.AddShape(5, 200, 50, 80, 40)  # no text
        line = sht.api.Shapes.AddLine(10, 10, 110, 10)
        line.Line.EndArrowheadStyle = 3
        wb.save(str(path))
        wb.close()
    finally:
        try:
            app.quit()
        except Exception:
            pass


def test_lightモードではCOMに触れずセルとテーブルのみ(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Test that light mode avoids COM and still emits OOXML rich artifacts."""

    path = tmp_path / "book.xlsx"
    _make_basic_book(path)

    def _boom(*_a: object, **_k: object) -> Never:
        """Raise if the COM path is accessed during the test."""

        raise AssertionError("COM should not be accessed in light mode")

    monkeypatch.setattr("exstruct.core.pipeline.xlwings_workbook", _boom)
    monkeypatch.setattr(
        "exstruct.core.backends.ooxml_backend.read_sheet_drawings",
        lambda _path: {
            "Sheet1": SheetDrawingData(
                shapes=[
                    OoxmlShapeInfo(
                        ref=DrawingShapeRef(
                            drawing_id=1,
                            name="Shape 1",
                            kind="shape",
                            left=12,
                            top=24,
                            width=80,
                            height=36,
                        ),
                        text="shape",
                        shape_type="AutoShape-Rectangle",
                    )
                ],
                charts=[
                    OoxmlChartInfo(
                        name="Chart 1",
                        chart_type="Line",
                        title="title",
                        y_axis_title="Y",
                        y_axis_range=[],
                        series=[],
                        anchor_left=48,
                        anchor_top=72,
                        anchor_width=120,
                        anchor_height=90,
                    )
                ],
            )
        },
    )
    data = extract(path, mode="light")
    sheet = next(iter(data.sheets.values()))
    assert len(sheet.shapes) == 1
    assert sheet.shapes[0].provenance == "python_ooxml"
    assert len(sheet.charts) == 1
    assert sheet.charts[0].provenance == "python_ooxml"


@pytest.mark.com  # type: ignore[misc]
def test_standardモードはテキストなし図形を除外する(tmp_path: Path) -> None:
    """Test that standard mode excludes textless non-arrow shapes."""

    _ensure_excel()
    path = tmp_path / "shapes.xlsx"
    _make_shapes_book(path)

    data = extract(path, mode="standard")
    shapes = data.sheets["Sheet1"].shapes
    texts = [s.text for s in shapes]
    assert "rect" in texts
    # textless shapes that remain must be矢印系のみ
    for s in shapes:
        if s.text != "":
            continue
        assert isinstance(s, Arrow)
        assert s.direction is not None or s.begin_arrow_style is not None


@pytest.mark.com  # type: ignore[misc]
def test_verboseモードでは全図形と幅高さが出力される(tmp_path: Path) -> None:
    """Test that verbose mode emits all shapes with width and height."""

    _ensure_excel()
    path = tmp_path / "shapes.xlsx"
    _make_shapes_book(path)

    data = extract(path, mode="verbose")
    shapes = data.sheets["Sheet1"].shapes
    texts = [s.text for s in shapes]
    assert "" in texts  # textless oval included
    for s in shapes:
        assert s.w is not None
        assert s.h is not None


def test_process_excelにモードが伝搬する(tmp_path: Path) -> None:
    """Test that process_excel forwards the selected mode."""

    path = tmp_path / "book.xlsx"
    out = tmp_path / "out.json"
    _make_basic_book(path)
    process_excel(path, out, mode="light")
    assert out.exists()


def test_process_excel_can_enable_backend_metadata(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that process_excel configures backend metadata inclusion."""

    captured: dict[str, object] = {}

    def _fake_process(
        self: ExStructEngine,
        file_path: Path,
        output_path: Path | None = None,
        **_kwargs: object,
    ) -> None:
        captured["file_path"] = file_path
        captured["output_path"] = output_path
        captured["include_backend_metadata"] = (
            self.output.filters.include_backend_metadata
        )

    monkeypatch.setattr("exstruct.ExStructEngine.process", _fake_process)
    path = tmp_path / "book.xlsx"
    out = tmp_path / "out.json"
    _make_basic_book(path)
    process_excel(path, out, mode="light", include_backend_metadata=True)
    assert captured["file_path"] == path
    assert captured["output_path"] == out
    assert captured["include_backend_metadata"] is True


def test_process_excel_keeps_print_area_filter_on_engine_auto_default(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify that process_excel leaves print-area inclusion on the engine default."""

    captured: dict[str, object] = {}

    def _fake_process(
        self: ExStructEngine,
        file_path: Path,
        output_path: Path | None = None,
        **_kwargs: object,
    ) -> None:
        captured["file_path"] = file_path
        captured["output_path"] = output_path
        captured["include_print_areas"] = self.output.filters.include_print_areas

    monkeypatch.setattr("exstruct.ExStructEngine.process", _fake_process)
    path = tmp_path / "book.xlsx"
    out = tmp_path / "out.json"
    _make_basic_book(path)
    process_excel(path, out, mode="light")
    assert captured["file_path"] == path
    assert captured["output_path"] == out
    assert captured["include_print_areas"] is None


def test_libreofficeモードを受け付ける(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Test that LibreOffice mode is accepted by extraction."""

    path = tmp_path / "book.xlsx"
    _make_basic_book(path)

    class _Backend:
        """Backend test double used by pipeline and mode tests."""

        def extract_shapes(self, *, mode: str) -> dict[str, list[Shape]]:
            """Provide the shape-extraction behavior for this test double."""

            _ = mode
            return {"Sheet1": [Shape(id=1, text="s", l=0, t=0)]}

        def extract_charts(self, *, mode: str) -> dict[str, list[Chart]]:
            """Provide the chart-extraction behavior for this test double."""

            _ = mode
            return {
                "Sheet1": [
                    Chart(
                        name="Chart 1",
                        chart_type="Line",
                        title="title",
                        y_axis_title="",
                        y_axis_range=[],
                        series=[],
                        l=0,
                        t=0,
                    )
                ]
            }

    monkeypatch.setattr(
        "exstruct.core.pipeline.resolve_rich_backend",
        lambda **_kwargs: _Backend(),
    )
    data = extract(path, mode="libreoffice")
    sheet = next(iter(data.sheets.values()))
    assert len(sheet.shapes) == 1
    assert len(sheet.charts) == 1


def test_invalidモードはエラーになる(tmp_path: Path) -> None:
    """Test that invalid extraction modes raise errors."""

    path = tmp_path / "book.xlsx"
    _make_basic_book(path)
    with pytest.raises(ValueError):
        extract(path, mode=cast(ExtractionMode, "invalid"))

    out = tmp_path / "out.json"
    with pytest.raises(ValueError):
        process_excel(path, out, mode=cast(ExtractionMode, "invalid"))


def test_libreofficeモードはxlsを拒否する(tmp_path: Path) -> None:
    """Test that LibreOffice mode rejects .xls inputs."""

    path = tmp_path / "book.xls"
    path.write_bytes(b"not-a-real-xls")
    with pytest.raises(ValueError, match="not supported in libreoffice mode"):
        extract(path, mode="libreoffice")


def test_CLIのmode引数バリデーション(tmp_path: Path) -> None:
    """Test CLI mode argument validation."""

    path = tmp_path / "book.xlsx"
    _make_basic_book(path)
    cmd = [
        sys.executable,
        "-m",
        "exstruct.cli.main",
        str(path),
        "--mode",
        "invalid",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode != 0


def test_process_excel_rejects_rendering_in_libreoffice_mode(tmp_path: Path) -> None:
    """Verify that process_excel rejects rendering in LibreOffice mode."""

    path = tmp_path / "book.xlsx"
    out = tmp_path / "out.json"
    _make_basic_book(path)

    with pytest.raises(ConfigError, match="does not support PDF/PNG rendering"):
        process_excel(path, out, mode="libreoffice", pdf=True)


def test_process_excel_rejects_auto_page_breaks_in_libreoffice_mode(
    tmp_path: Path,
) -> None:
    """Verify that process_excel rejects auto page-break export in LibreOffice mode."""

    path = tmp_path / "book.xlsx"
    out = tmp_path / "out.json"
    _make_basic_book(path)

    with pytest.raises(ConfigError, match="does not support auto page-break export"):
        process_excel(
            path,
            out,
            mode="libreoffice",
            auto_page_breaks_dir=tmp_path / "auto",
        )


def test_process_excel_rejects_rendering_and_auto_page_breaks_in_libreoffice_mode(
    tmp_path: Path,
) -> None:
    """Verify that process_excel rejects rendering and auto page-break export together."""

    path = tmp_path / "book.xlsx"
    out = tmp_path / "out.json"
    _make_basic_book(path)

    with pytest.raises(
        ConfigError,
        match="does not support PDF/PNG rendering or auto page-break export",
    ):
        process_excel(
            path,
            out,
            mode="libreoffice",
            pdf=True,
            auto_page_breaks_dir=tmp_path / "auto",
        )


def test_extract_workbook_rejects_auto_page_breaks_in_libreoffice_mode(
    tmp_path: Path,
) -> None:
    """Verify that extract workbook rejects auto page breaks in LibreOffice mode."""

    with pytest.raises(ConfigError, match="does not support auto page-break export"):
        extract_workbook(
            tmp_path / "book.xlsx",
            mode="libreoffice",
            include_auto_page_breaks=True,
        )


def _make_multi_sheet_book(path: Path) -> None:
    """Create multi sheet book for tests."""

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1["A1"] = "v1"
    ws2 = wb.create_sheet("Data 02")
    ws2["A1"] = "v2"
    wb.save(path)


def test_process_excel_defaults_to_stdout(
    tmp_path: Path, capsys: CaptureFixture[str]
) -> None:
    """Verify that process_excel writes to stdout when no output path is given."""

    path = tmp_path / "book.xlsx"
    _make_basic_book(path)

    process_excel(path, output_path=None, mode="light", pretty=True)
    captured = capsys.readouterr().out
    assert '"book_name": "book.xlsx"' in captured
    assert "Sheet1" in captured


def test_process_excel_sheets_dir_output(tmp_path: Path) -> None:
    """Verify that process_excel writes one JSON file per sheet."""

    path = tmp_path / "book.xlsx"
    _make_multi_sheet_book(path)
    sheets_dir = tmp_path / "sheets"

    process_excel(path, output_path=None, mode="light", sheets_dir=sheets_dir)

    files = list(sheets_dir.glob("*.json"))
    assert len(files) == 2
    names = {f.stem for f in files}
    assert "Sheet1" in names
    assert "Data 02" in names


def test_process_excel_light_print_areas_dir_output(tmp_path: Path) -> None:
    """Verify that light-mode process_excel writes print-area files."""

    path = tmp_path / "book.xlsx"
    _make_basic_book(path)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "v1"
    ws["B1"] = "v2"
    ws.print_area = "A1:B1"
    wb.save(path)
    wb.close()

    areas_dir = tmp_path / "areas"
    process_excel(path, output_path=None, mode="light", print_areas_dir=areas_dir)

    files = list(areas_dir.glob("*.json"))
    assert len(files) == 1
    assert files[0].stem == "Sheet1_area1_r1-1_c0-1"


def test_CLI_defaults_to_stdout(tmp_path: Path) -> None:
    """Verify that the CLI writes JSON to stdout by default."""

    path = tmp_path / "book.xlsx"
    _make_basic_book(path)
    cmd = [
        sys.executable,
        "-m",
        "exstruct.cli.main",
        str(path),
        "--mode",
        "light",
        "--pretty",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    assert '"book_name": "book.xlsx"' in result.stdout
