"""Fresh-process import boundaries exercised by real extraction, not mocks."""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
from typing import cast

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Border, Side
import pytest

_SHAPE_SAMPLE = (
    Path(__file__).resolve().parents[2] / "sample/flowchart/sample-shape-connector.xlsx"
)


def _probe(code: str) -> dict[str, object]:
    result = subprocess.run(
        [sys.executable, "-c", code], capture_output=True, text=True, timeout=60
    )
    assert result.returncode == 0, result.stderr or result.stdout
    return cast(dict[str, object], json.loads(result.stdout))


@pytest.mark.parametrize("suffix", [".xlsx", ".xlsm"])
@pytest.mark.parametrize("without_pillow", [False, True])
def test_real_light_extraction_isolates_backends(
    tmp_path: Path, suffix: str, without_pillow: bool
) -> None:
    path = tmp_path / f"rich{suffix}"
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = "Data"
    for row in [("Name", "Value"), ("A", 12), ("B", 24)]:
        sheet.append(row)
    border = Border(*(Side(style="thin") for _ in range(4)))
    for row in sheet:
        for cell in row:
            cell.border = border
    sheet.print_area = "A1:B3"
    chart = BarChart()
    chart.add_data(
        Reference(sheet, min_col=2, min_row=1, max_row=3), titles_from_data=True
    )
    sheet.add_chart(chart, "D1")
    book.save(path)
    book.close()
    # This finder simulates an absent optional dependency only inside this probe.
    # Production extraction does not modify module caches or import machinery.
    payload = _probe(
        f"""
import importlib.abc
import json
import os
import sys

class MissingHeavyDependencies(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname.split(".", 1)[0] in {{"openpyxl", "xlwings", "scipy", "pandas"}}:
            raise ModuleNotFoundError("Heavy dependency unavailable", name=fullname)
sys.meta_path.insert(0, MissingHeavyDependencies())

if {without_pillow!r}:
    class MissingPillow(importlib.abc.MetaPathFinder):
        def find_spec(self, fullname, path=None, target=None):
            if fullname == "PIL" or fullname.startswith("PIL."):
                raise ModuleNotFoundError("Pillow unavailable", name=fullname)
    sys.meta_path.insert(0, MissingPillow())

os.environ.pop("EXSTRUCT_BORDER_CLUSTER_BACKEND", None)
import exstruct
workbook = exstruct.extract({str(path)!r}, mode="light")
sheet = workbook.sheets["Data"]
assert sheet.rows[1].c == {{"0": "A", "1": 12}}
assert sheet.table_candidates
assert sheet.print_areas
assert sheet.charts and sheet.charts[0].provenance == "python_ooxml"
# Existing real OOXML shapes/connectors must survive the import split too.
diagram = exstruct.extract({str(_SHAPE_SAMPLE)!r}, mode="light")
assert any(s.shapes for s in diagram.sheets.values())
for s in diagram.sheets.values():
    assert all(shape.provenance == "python_ooxml" for shape in s.shapes)
forbidden = ["openpyxl", "xlwings", "scipy", "pandas", "xlrd", "exstruct.core.backends.com_backend",
             "exstruct.core.backends.libreoffice_backend", "exstruct.core.charts",
             "exstruct.core.shapes", "exstruct.render", "pypdfium2"]
print(json.dumps({{
    "loaded": [name for name in forbidden if any(
        m == name or m.startswith(name + ".") for m in sys.modules)],
    "pillow": "PIL" in sys.modules,
}}))
"""
    )
    assert payload["loaded"] == []
    if without_pillow:
        assert payload["pillow"] is False


def test_shared_backend_contracts_do_not_load_concrete_dependencies() -> None:
    payload = _probe("""
import json
import sys
from exstruct.core.backends.base import Backend, RichBackend
import typing
assert typing.get_type_hints(Backend.extract_merged_cells)
assert typing.get_type_hints(RichBackend.extract_shapes)
roots = ["numpy", "pandas", "openpyxl", "xlwings", "scipy", "PIL"]
print(json.dumps({"loaded": [name for name in roots if name in sys.modules]}))
""")
    assert payload["loaded"] == []


def test_runtime_annotations_remain_resolvable() -> None:
    payload = _probe("""
import json
import typing
import exstruct
from exstruct.core import cells, integrate, pipeline, workbook
from exstruct.core.backends import Backend, OpenpyxlBackend
assert cells.MergedCellRange is pipeline.MergedCellRange
for function in (exstruct.extract, integrate.extract_workbook,
                 pipeline.resolve_rich_backend, pipeline.step_extract_shapes_com,
                 workbook.xlwings_workbook, cells.detect_tables,
                 OpenpyxlBackend.extract_cells, Backend.extract_cells):
    assert typing.get_type_hints(function)
assert typing.get_type_hints(pipeline.PipelinePlan)
print(json.dumps({"resolved": True}))
""")
    assert payload == {"resolved": True}


@pytest.mark.parametrize("backend", ["python", "numpy"])
def test_scipy_loads_only_for_accelerated_clustering(backend: str) -> None:
    payload = _probe(
        f"""
import json
import os
import sys
import numpy as np
os.environ["EXSTRUCT_BORDER_CLUSTER_BACKEND"] = {backend!r}
from exstruct.core.cells import detect_border_clusters
assert "scipy" not in sys.modules
assert detect_border_clusters(np.ones((2, 2), dtype=bool)) == [(0, 0, 1, 1)]
print(json.dumps({{"scipy": "scipy" in sys.modules}}))
"""
    )
    assert payload == {"scipy": backend == "numpy"}


def test_compatibility_wrappers_observe_live_module_overrides() -> None:
    payload = _probe("""
import json
from types import SimpleNamespace
from exstruct.core import charts, pipeline, shapes, workbook
import openpyxl
import xlwings
original_app = xlwings.App
workbook.xw.App = "live module override"
assert xlwings.App == "live module override"
workbook.xw.App = original_app
shapes.get_shapes_with_position = lambda workbook, mode: {"live": []}
charts.get_charts = lambda sheet, mode: []
assert pipeline.get_shapes_with_position(object()) == {"live": []}
assert pipeline.get_charts(object()) == []
openpyxl.load_workbook = lambda *args, **kwargs: "live loader"
assert workbook.load_workbook("ignored.xlsx") == "live loader"
workbook.load_workbook = lambda *args, **kwargs: SimpleNamespace(close=lambda: None)
with workbook.openpyxl_workbook("ignored.xlsx", data_only=True, read_only=False) as book:
    assert hasattr(book, "close")
print(json.dumps({"live_overrides": True}))
""")
    assert payload == {"live_overrides": True}


def test_pipeline_backend_alias_overrides_avoid_concrete_imports() -> None:
    payload = _probe("""
import json
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
from exstruct.core import pipeline
inputs = SimpleNamespace(mode='standard', file_path=Path('unused.xlsx'),
                         include_default_background=False, ignore_colors=None)
workbook = object()
com_rich = Mock(return_value='COM_OVERRIDE')
pipeline.ComRichBackend = com_rich
assert pipeline.resolve_rich_backend(inputs=inputs, workbook=workbook) == 'COM_OVERRIDE'
com_rich.assert_called_once_with(workbook)
inputs.mode = 'libreoffice'
lo_rich = Mock(return_value='LO_OVERRIDE')
pipeline.LibreOfficeRichBackend = lo_rich
assert pipeline.resolve_rich_backend(inputs=inputs) == 'LO_OVERRIDE'
lo_rich.assert_called_once_with(inputs.file_path)
backend = SimpleNamespace(extract_print_areas=Mock(return_value={'p': []}),
                          extract_auto_page_breaks=Mock(return_value={'a': []}),
                          extract_formulas_map=Mock(return_value={'f': {}}),
                          extract_colors_map=Mock(return_value={'c': {}}))
factory = Mock(return_value=backend)
pipeline.ComBackend = factory
artifacts = pipeline.ExtractionArtifacts()
pipeline.step_extract_print_areas_com(inputs, artifacts, workbook)
pipeline.step_extract_auto_page_breaks_com(inputs, artifacts, workbook)
pipeline.step_extract_formulas_map_com(inputs, artifacts, workbook)
pipeline.step_extract_colors_map_com(inputs, artifacts, workbook)
assert artifacts.print_area_data == {'p': []}
assert artifacts.auto_page_break_data == {'a': []}
assert artifacts.formulas_map_data == {'f': {}}
assert artifacts.colors_map_data == {'c': {}}
assert factory.call_count == 4
for call in factory.call_args_list:
    assert call.args == (workbook,)
backend.extract_colors_map.assert_called_once_with(include_default_background=False,
                                                  ignore_colors=None)
for name in ('xlwings', 'exstruct.core.backends.com_backend',
             'exstruct.core.backends.libreoffice_backend'):
    assert name not in sys.modules
print(json.dumps({'overrides': True}))
""")
    assert payload == {"overrides": True}
