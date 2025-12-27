from pathlib import Path

from openpyxl import Workbook

from exstruct import extract


def _make_book_with_print_area(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "x"
    ws.print_area = "A1:B2"
    wb.save(path)
    wb.close()


def test_light_mode_includes_print_areas_without_com(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _make_book_with_print_area(path)

    wb_data = extract(path, mode="light")
    areas = wb_data.sheets["Sheet1"].print_areas
    assert len(areas) == 1
    area = areas[0]
    assert (area.r1, area.c1, area.r2, area.c2) == (1, 0, 2, 1)
