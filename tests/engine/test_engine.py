import json
from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch

from exstruct.engine import (
    DestinationOptions,
    ExStructEngine,
    FilterOptions,
    OutputOptions,
    StructOptions,
)
from exstruct.models import (
    CellRow,
    Chart,
    ChartSeries,
    MergedCells,
    PrintArea,
    Shape,
    SheetData,
    WorkbookData,
)


def test_engine_extract_uses_mode(monkeypatch: MonkeyPatch, tmp_path: Path) -> None:
    called: dict[str, object] = {}

    def fake_extract(
        path: Path,
        mode: str,
        include_cell_links: bool = False,
        include_print_areas: bool = True,
        include_auto_page_breaks: bool = False,
        include_colors_map: bool = False,
        include_default_background: bool = False,
        ignore_colors: set[str] | None = None,
        include_formulas_map: bool | None = None,
        include_merged_cells: bool | None = None,
        include_merged_values_in_rows: bool = True,
    ) -> WorkbookData:
        """
        Test helper that simulates workbook extraction for unit tests.
        
        Records the received `mode` and `include_print_areas` into the outer `called` mapping and returns a minimal WorkbookData whose `book_name` is the input path's filename and whose `sheets` is empty.
        
        Parameters:
            path (Path): Path to the workbook; its filename is used for the returned WorkbookData.book_name.
            mode (str): Extraction mode passed through and recorded.
            include_print_areas (bool): Whether print areas were requested; the value is recorded in `called`.
        
        Returns:
            WorkbookData: A WorkbookData instance with `book_name` set to path.name and an empty `sheets` mapping.
        """
        called["mode"] = mode
        called["include_print_areas"] = include_print_areas
        return WorkbookData(book_name=path.name, sheets={})

    monkeypatch.setattr("exstruct.engine.extract_workbook", fake_extract)
    engine = ExStructEngine(options=StructOptions(mode="standard"))
    engine.extract(tmp_path / "book.xlsx", mode="verbose")
    assert called["mode"] == "verbose"


def _sample_workbook() -> WorkbookData:
    shape = Shape(id=1, text="x", l=0, t=0, w=10, h=10, type="Rect")
    chart = Chart(
        name="c1",
        chart_type="Line",
        title=None,
        y_axis_title="",
        y_axis_range=[],
        series=[ChartSeries(name="s1")],
        l=0,
        t=0,
        error=None,
    )
    sheet = SheetData(
        rows=[CellRow(r=1, c={"0": "v"}, links={"0": "http://example.com"})],
        shapes=[shape],
        charts=[chart],
        table_candidates=["A1:B2"],
        print_areas=[PrintArea(r1=1, c1=0, r2=3, c2=2)],
        merged_cells=MergedCells(items=[(1, 0, 1, 1, "merged")]),
    )
    return WorkbookData(book_name="book.xlsx", sheets={"Sheet1": sheet})


def test_engine_serialize_filters_shapes(tmp_path: Path) -> None:
    wb = _sample_workbook()
    engine = ExStructEngine(
        output=OutputOptions(filters=FilterOptions(include_shapes=False))
    )
    text = engine.serialize(wb, fmt="json")
    assert '"shapes"' not in text


def test_engine_serialize_filters_tables(tmp_path: Path) -> None:
    wb = _sample_workbook()
    engine = ExStructEngine(
        output=OutputOptions(filters=FilterOptions(include_tables=False))
    )
    text = engine.serialize(wb, fmt="json")
    assert "table_candidates" not in text


def test_engine_serialize_filters_merged_cells() -> None:
    wb = _sample_workbook()
    engine = ExStructEngine(
        output=OutputOptions(filters=FilterOptions(include_merged_cells=False))
    )
    text = engine.serialize(wb, fmt="json")
    assert "merged_cells" not in text


def test_engine_include_cell_links_toggle() -> None:
    wb = _sample_workbook()
    # By default links remain (already present)
    engine = ExStructEngine()
    text = engine.serialize(wb, fmt="json")
    data = json.loads(text)
    # Navigate workbook -> sheets -> sheet1 -> rows -> row[0] -> links -> cell '0'
    assert data["sheets"]["Sheet1"]["rows"][0]["links"]["0"] == "http://example.com"

    engine_no_links = ExStructEngine(
        output=OutputOptions(
            filters=FilterOptions(
                include_rows=True,
                include_shapes=True,
                include_charts=True,
                include_tables=True,
            )
        )
    )
    # overwrite output options to drop links by excluding rows manually would drop links, but links live inside rows; not filtered here.
    # Explicitly reserialize after removing links at row level
    wb_no_links = WorkbookData(
        book_name=wb.book_name,
        sheets={
            "Sheet1": SheetData(
                rows=[CellRow(r=1, c={"0": "v"}, links=None)],
                shapes=[],
                charts=[],
                table_candidates=[],
            )
        },
    )
    text2 = engine_no_links.serialize(wb_no_links, fmt="json")
    assert "links" not in text2


def test_engine_export_respects_sheets_dir(tmp_path: Path) -> None:
    wb = _sample_workbook()
    sheets_dir = tmp_path / "sheets"
    engine = ExStructEngine(
        output=OutputOptions(destinations=DestinationOptions(sheets_dir=sheets_dir))
    )
    out = tmp_path / "out.json"
    engine.export(wb, output_path=out)
    assert out.exists()
    assert sheets_dir.exists()
    files = list(sheets_dir.glob("*.json"))
    assert len(files) == 1


def test_engine_export_print_areas_dir(tmp_path: Path) -> None:
    wb = _sample_workbook()
    areas_dir = tmp_path / "areas"
    engine = ExStructEngine(
        output=OutputOptions(destinations=DestinationOptions(print_areas_dir=areas_dir))
    )
    out = tmp_path / "out.json"
    engine.export(wb, output_path=out)
    files = list(areas_dir.glob("*.json"))
    assert out.exists()
    assert files
    content = files[0].read_text(encoding="utf-8")
    assert '"area"' in content
    assert '"sheet_name": "Sheet1"' in content


def test_engine_export_print_areas_respects_include_flag(tmp_path: Path) -> None:
    wb = _sample_workbook()
    areas_dir = tmp_path / "areas"
    engine = ExStructEngine(
        output=OutputOptions(
            filters=FilterOptions(include_print_areas=False),
            destinations=DestinationOptions(print_areas_dir=areas_dir),
        )
    )
    out = tmp_path / "out.json"
    engine.export(wb, output_path=out)
    assert out.exists()
    # When print areas are excluded, no per-area files should be written.
    assert not areas_dir.exists() or not list(areas_dir.glob("*"))


def test_engine_export_print_areas_light_mode_skips_shapes_and_charts(
    tmp_path: Path,
) -> None:
    wb = _sample_workbook()
    areas_dir = tmp_path / "areas"
    engine = ExStructEngine(
        options=StructOptions(mode="light"),
        output=OutputOptions(
            destinations=DestinationOptions(print_areas_dir=areas_dir)
        ),
    )
    out = tmp_path / "out.json"
    engine.export(wb, output_path=out, fmt="json")
    assert out.exists()
    # light mode should not emit per-area files (print areas are absent in light extraction)
    assert not areas_dir.exists() or not list(areas_dir.glob("*"))


def test_engine_export_accepts_string_paths(tmp_path: Path) -> None:
    wb = _sample_workbook()
    sheets_dir = tmp_path / "sheets"
    areas_dir = tmp_path / "areas"
    out = tmp_path / "out.json"
    engine = ExStructEngine(
        output=OutputOptions(
            destinations=DestinationOptions(
                sheets_dir=str(sheets_dir),
                print_areas_dir=str(areas_dir),
            )
        )
    )
    engine.export(wb, output_path=str(out), fmt="json")

    assert out.exists()
    assert sheets_dir.exists()
    assert list(sheets_dir.glob("*.json"))
    assert areas_dir.exists()
    assert list(areas_dir.glob("*.json"))


def test_engine_process_normalizes_string_paths(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    calls: dict[str, object] = {}

    def fake_extract(
        self: ExStructEngine, file_path: Path, *, mode: str | None = None
    ) -> WorkbookData:
        calls["extract_path"] = file_path
        return _sample_workbook()

    def fake_export(
        self: ExStructEngine,
        wb: WorkbookData,
        *,
        output_path: Path | None = None,
        **_: object,
    ) -> None:
        calls["export_output_path"] = output_path
        return None

    def fake_pdf(file_path: Path, pdf_path: Path) -> None:
        calls["pdf_path"] = pdf_path
        calls["pdf_input_path"] = file_path

    def fake_images(file_path: Path, images_dir: Path, *, dpi: int) -> None:
        calls["images_dir"] = images_dir
        calls["images_input_path"] = file_path
        calls["dpi"] = dpi

    monkeypatch.setattr(ExStructEngine, "extract", fake_extract, raising=True)
    monkeypatch.setattr(ExStructEngine, "export", fake_export, raising=True)
    monkeypatch.setattr("exstruct.engine.export_pdf", fake_pdf, raising=True)
    monkeypatch.setattr(
        "exstruct.engine.export_sheet_images", fake_images, raising=True
    )

    engine = ExStructEngine()
    input_path = tmp_path / "input.xlsx"
    input_path.write_text("", encoding="utf-8")
    output_path = tmp_path / "out.json"

    engine.process(
        str(input_path),
        output_path=str(output_path),
        pdf=True,
        image=True,
        dpi=144,
        out_fmt="json",
    )

    assert isinstance(calls["extract_path"], Path)
    assert isinstance(calls["pdf_input_path"], Path)
    assert isinstance(calls["images_input_path"], Path)
    assert isinstance(calls["export_output_path"], Path)
    assert isinstance(calls["pdf_path"], Path)
    assert calls["pdf_path"].suffix == ".pdf"
    assert isinstance(calls["images_dir"], Path)
    assert calls["images_dir"].name.endswith("_images")
    assert calls["dpi"] == 144