from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch
from openpyxl import Workbook

from exstruct.core.backends.com_backend import ComBackend
from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend
from exstruct.core.ranges import parse_range_zero_based


def test_openpyxl_backend_extract_cells_switches_link_mode(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    calls: list[str] = []

    def fake_cells(file_path: Path) -> dict[str, list[object]]:
        calls.append("cells")
        return {}

    def fake_cells_links(file_path: Path) -> dict[str, list[object]]:
        calls.append("links")
        return {}

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.extract_sheet_cells",
        fake_cells,
    )
    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.extract_sheet_cells_with_links",
        fake_cells_links,
    )

    backend = OpenpyxlBackend(tmp_path / "book.xlsx")
    backend.extract_cells(include_links=False)
    backend.extract_cells(include_links=True)

    assert calls == ["cells", "links"]


def test_openpyxl_backend_detect_tables_handles_failure(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    def fake_detect(file_path: Path, sheet_name: str) -> list[str]:
        raise RuntimeError("boom")

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.detect_tables_openpyxl",
        fake_detect,
    )

    backend = OpenpyxlBackend(tmp_path / "book.xlsx")
    assert backend.detect_tables("Sheet1") == []


def test_com_backend_extract_colors_map_returns_none_on_failure(
    monkeypatch: MonkeyPatch,
) -> None:
    def fake_colors_map(
        workbook: object,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        raise RuntimeError("boom")

    monkeypatch.setattr(
        "exstruct.core.backends.com_backend.extract_sheet_colors_map_com",
        fake_colors_map,
    )

    class DummyWorkbook:
        pass

    backend = ComBackend(DummyWorkbook())
    assert (
        backend.extract_colors_map(include_default_background=False, ignore_colors=None)
        is None
    )


def test_openpyxl_backend_extract_print_areas(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["A", "B"])
    ws.append([1, 2])
    ws.print_area = "A1:B2"
    file_path = tmp_path / "print_area.xlsx"
    wb.save(file_path)
    wb.close()

    backend = OpenpyxlBackend(file_path)
    areas = backend.extract_print_areas()
    assert "Sheet1" in areas
    assert areas["Sheet1"]
    assert areas["Sheet1"][0].r1 == 1
    assert areas["Sheet1"][0].c1 == 0


def test_parse_range_zero_based_parses_sheet_prefix() -> None:
    bounds = parse_range_zero_based("Sheet1!A1:B2")
    assert bounds is not None
    assert bounds.r1 == 0
    assert bounds.c1 == 0
    assert bounds.r2 == 1
    assert bounds.c2 == 1
