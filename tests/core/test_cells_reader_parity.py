"""Cell contracts captured against the pandas reader before its replacement."""

from datetime import date, datetime, time, timedelta
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from openpyxl.worksheet.hyperlink import Hyperlink
import pytest

from exstruct.core import cells
from exstruct.models import CellRow


def make_parity_book(path: Path) -> None:
    """Cover reader conversions, missing tokens, cached formulas and row gaps."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Values"
    ws.append([123, 1.5, "00123", "1.50", True, False, "   ", " x "])
    ws.append([0, -5, 1e20, "1e2", "+001", "-.50", "1.", " NA "])
    ws.append(
        [
            date(2020, 1, 2),
            datetime(2020, 1, 2, 3, 4, 5, 123000),
            time(3, 4, 5, 123000),
            timedelta(days=2, hours=3),
        ]
    )
    tokens = [
        "",
        "#N/A",
        "#N/A N/A",
        "#NA",
        "-1.#IND",
        "-1.#QNAN",
        "-NaN",
        "-nan",
        "1.#IND",
        "1.#QNAN",
        "<NA>",
        "N/A",
        "NA",
        "NULL",
        "NaN",
        "None",
        "n/a",
        "nan",
        "null",
    ]
    for col, token in enumerate(tokens, start=1):
        ws.cell(4, col, token).data_type = "s"
    for col, error in enumerate(
        ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!"], start=1
    ):
        ws.cell(5, col, error)
    ws["D9"] = "sparse"
    ws["A10"] = "=1+1"
    ws["B10"] = "=TRUE()"
    ws["C10"] = '= "cached"'
    ws["D10"] = "=1/0"
    ws["E10"] = "=1+2"
    ws["A11"] = "visible"
    ws["B11"].hyperlink = "https://example.test/na"
    ws["B11"] = "NA"
    ws["A12"].hyperlink = "https://example.test/only"
    ws["A12"] = "NA"
    ws["C11"].hyperlink = Hyperlink(ref="C11", location="Values!A1")
    ws["C11"] = "internal"
    ws["A14"] = "first"
    ws.merge_cells("A14:B14")
    wb.create_sheet("Empty")
    wb.create_sheet("Last")["C3"] = "last"
    wb.save(path)
    wb.close()
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    sheet = parts["xl/worksheets/sheet1.xml"].decode().replace("<v />", "<v></v>")
    sheet = sheet.replace("<f>1+1</f><v></v>", "<f>1+1</f><v>2</v>")
    sheet = sheet.replace("<f>TRUE()</f><v></v>", "<f>TRUE()</f><v>1</v>")
    sheet = sheet.replace('r="B10"', 'r="B10" t="b"')
    sheet = sheet.replace('<f>= "cached"</f><v></v>', '<f>= "cached"</f><v>cached</v>')
    # The formula text is stored without its initial equals sign.
    sheet = sheet.replace('<f> "cached"</f><v></v>', '<f> "cached"</f><v>cached</v>')
    sheet = sheet.replace('r="C10"', 'r="C10" t="str"')
    sheet = sheet.replace("<f>1/0</f><v></v>", "<f>1/0</f><v>#DIV/0!</v>")
    sheet = sheet.replace('r="D10"', 'r="D10" t="e"')
    parts["xl/worksheets/sheet1.xml"] = sheet.encode()
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)


def dump_rows(data: dict[str, list[CellRow]]) -> dict[str, list[dict[str, object]]]:
    return {
        name: [row.model_dump(mode="json") for row in rows]
        for name, rows in data.items()
    }


@pytest.mark.parametrize("suffix", [".xlsx", ".xlsm"])
def test_cells_match_frozen_old_reader(tmp_path: Path, suffix: str) -> None:
    path = tmp_path / f"parity{suffix}"
    make_parity_book(path)
    expected = json.loads(
        (Path(__file__).parent / "cells_reader_expected.json").read_text(
            encoding="utf-8"
        )
    )
    assert dump_rows(cells.extract_sheet_cells(path)) == expected["xlsx"]
    assert dump_rows(cells.extract_sheet_cells_with_links(path)) == expected["links"]


def test_real_xls_matches_frozen_old_reader() -> None:
    path = Path(__file__).parents[1] / "assets" / "sample.xls"
    expected = json.loads(
        (Path(__file__).parent / "cells_reader_expected.json").read_text(
            encoding="utf-8"
        )
    )
    assert dump_rows(cells.extract_sheet_cells(path)) == expected["xls"]


def test_cells_match_live_pandas_oracle(tmp_path: Path) -> None:
    pd = pytest.importorskip("pandas")
    path = tmp_path / "oracle.xlsx"
    make_parity_book(path)
    expected: dict[str, list[CellRow]] = {}
    for name, frame in pd.read_excel(
        path, header=None, sheet_name=None, dtype=str
    ).items():
        rows = []
        for row_number, row in enumerate(
            frame.fillna("").itertuples(index=False, name=None), start=1
        ):
            values = {
                str(col): cells._coerce_numeric_preserve_format(str(value))
                for col, value in enumerate(row)
                if str(value).strip()
            }
            if values:
                rows.append(CellRow(r=row_number, c=values))
        expected[name] = rows
    assert cells.extract_sheet_cells(path) == expected
