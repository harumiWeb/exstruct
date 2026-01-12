from pathlib import Path

from openpyxl import Workbook
import pytest

from exstruct import process_excel

pytestmark = pytest.mark.render


def test_render_smoke_pdf_and_png(tmp_path: Path) -> None:
    # create a tiny workbook
    xlsx = tmp_path / "sample.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "hello"
    wb.save(xlsx)

    out_json = tmp_path / "out.json"
    process_excel(
        xlsx,
        output_path=out_json,
        out_fmt="json",
        pdf=True,
        image=True,
        dpi=72,
        mode="standard",
        pretty=True,
    )
    pdf_path = out_json.with_suffix(".pdf")
    images_dir = out_json.parent / f"{out_json.stem}_images"
    assert out_json.exists()
    assert pdf_path.exists()
    assert images_dir.exists()
    assert any(images_dir.glob("*.png"))


def test_render_multiple_print_ranges_images(tmp_path: Path) -> None:
    xlsx = (
        Path(__file__).resolve().parents[1]
        / "assets"
        / "multiple_print_ranges_4sheets.xlsx"
    )
    out_json = tmp_path / "out.json"
    process_excel(
        xlsx,
        output_path=out_json,
        out_fmt="json",
        image=True,
        dpi=72,
        mode="standard",
        pretty=True,
    )
    images_dir = out_json.parent / f"{out_json.stem}_images"
    images = list(images_dir.glob("*.png"))
    assert images_dir.exists()
    assert len(images) == 4
