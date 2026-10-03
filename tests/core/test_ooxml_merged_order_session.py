"""Merged-range ordering parity for the dependency-free OOXML session."""

from pathlib import Path

from openpyxl import Workbook

from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend
from exstruct.core.ooxml_session import OoxmlExtractionSession


def test_merged_ranges_keep_openpyxl_set_iteration_order(tmp_path: Path) -> None:
    path = tmp_path / "merged-order.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet"
    ranges = [
        "F55:K55",
        "A47:B53",
        "I52:J53",
        "F51:G52",
        "I54:K54",
        "A10:B12",
        "C10:K10",
        "D7:E8",
        "H3:K3",
        "H6:H7",
        "G7:G8",
        "A22:A33",
    ]
    for index, reference in enumerate(ranges):
        anchor = reference.split(":", 1)[0]
        sheet[anchor] = f"merge-{index}"
        sheet.merge_cells(reference)
    workbook.save(path)
    workbook.close()

    expected = OpenpyxlBackend(path).extract_merged_cells()["Sheet"]
    with OoxmlExtractionSession(path) as session:
        actual = session.extract_merged_cells()["Sheet"]
    assert actual == expected
