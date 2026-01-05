from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch

from exstruct.core.pipeline import (
    ExtractionArtifacts,
    ExtractionInputs,
    build_cells_tables_workbook,
    build_com_pipeline,
    build_pre_com_pipeline,
    resolve_extraction_inputs,
)
from exstruct.models import CellRow, PrintArea


def test_build_pre_com_pipeline_respects_flags(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_merged_cells=False,
    )
    steps = build_pre_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert step_names == ["step_extract_cells"]


def test_build_pre_com_pipeline_includes_colors_map_for_light(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="light",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_merged_cells=True,
    )
    steps = build_pre_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert step_names == [
        "step_extract_cells",
        "step_extract_print_areas_openpyxl",
        "step_extract_colors_map_openpyxl",
        "step_extract_merged_cells_openpyxl",
    ]


def test_build_pre_com_pipeline_skips_merged_cells_when_disabled(
    tmp_path: Path,
) -> None:
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_merged_cells=False,
    )
    steps = build_pre_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert "step_extract_merged_cells_openpyxl" not in step_names


def test_build_com_pipeline_respects_flags(tmp_path: Path) -> None:
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=True,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_merged_cells=False,
    )
    steps = build_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert step_names == [
        "step_extract_shapes_com",
        "step_extract_charts_com",
        "step_extract_auto_page_breaks_com",
    ]


def test_build_com_pipeline_excludes_auto_page_breaks_when_disabled(
    tmp_path: Path,
) -> None:
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=False,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_merged_cells=False,
    )
    steps = build_com_pipeline(inputs)
    step_names = [step.__name__ for step in steps]
    assert "step_extract_auto_page_breaks_com" not in step_names


def test_build_com_pipeline_empty_for_light(tmp_path: Path) -> None:
    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="light",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=True,
        include_colors_map=True,
        include_default_background=False,
        ignore_colors=None,
        include_merged_cells=False,
    )
    steps = build_com_pipeline(inputs)
    assert steps == []


def test_resolve_extraction_inputs_defaults(tmp_path: Path) -> None:
    inputs = resolve_extraction_inputs(
        tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=None,
        include_print_areas=None,
        include_auto_page_breaks=False,
        include_colors_map=None,
        include_default_background=True,
        ignore_colors=None,
        include_merged_cells=None,
    )
    assert inputs.include_cell_links is False
    assert inputs.include_print_areas is True
    assert inputs.include_colors_map is False
    assert inputs.include_default_background is False
    assert inputs.include_merged_cells is True


def test_build_cells_tables_workbook_uses_print_areas(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    def fake_detect_tables(_: Path, __: str) -> list[str]:
        return ["A1:B2"]

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.detect_tables_openpyxl",
        fake_detect_tables,
    )

    inputs = ExtractionInputs(
        file_path=tmp_path / "book.xlsx",
        mode="standard",
        include_cell_links=False,
        include_print_areas=True,
        include_auto_page_breaks=False,
        include_colors_map=False,
        include_default_background=False,
        ignore_colors=None,
        include_merged_cells=True,
    )
    artifacts = ExtractionArtifacts(
        cell_data={"Sheet1": [CellRow(r=1, c={"0": "v"})]},
        print_area_data={"Sheet1": [PrintArea(r1=1, c1=0, r2=1, c2=0)]},
        merged_cell_data={"Sheet1": []},
    )
    wb = build_cells_tables_workbook(
        inputs=inputs,
        artifacts=artifacts,
        reason="test",
    )
    sheet = wb.sheets["Sheet1"]
    assert sheet.print_areas
    assert sheet.table_candidates == ["A1:B2"]
    assert sheet.merged_cells == []
