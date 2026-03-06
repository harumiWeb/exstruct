from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
import subprocess
from typing import cast
from xml.etree import ElementTree

from _pytest.monkeypatch import MonkeyPatch
import pytest

from exstruct.core.backends.libreoffice_backend import LibreOfficeRichBackend
from exstruct.core.libreoffice import (
    LibreOfficeSession,
    LibreOfficeSessionConfig,
    LibreOfficeUnavailableError,
)
from exstruct.core.ooxml_drawing import _parse_connector_node
from exstruct.models import Arrow


class _DummySession:
    def __enter__(self) -> _DummySession:
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        _ = exc_type
        _ = exc
        _ = tb

    def load_workbook(self, file_path: Path) -> object:
        return {"file_path": str(file_path)}

    def close_workbook(self, workbook: object) -> None:
        _ = workbook


def _dummy_session_factory() -> LibreOfficeSession:
    return cast(LibreOfficeSession, _DummySession())


def test_libreoffice_backend_extracts_connector_graph_from_sample() -> None:
    backend = LibreOfficeRichBackend(
        Path("sample/flowchart/sample-shape-connector.xlsx"),
        session_factory=_dummy_session_factory,
    )
    shape_data = backend.extract_shapes(mode="libreoffice")
    assert shape_data
    first_sheet = next(iter(shape_data.values()))
    connectors = [shape for shape in first_sheet if isinstance(shape, Arrow)]
    resolved = [
        connector
        for connector in connectors
        if connector.begin_id is not None and connector.end_id is not None
    ]
    assert len(connectors) >= 10
    assert len(resolved) >= 10
    assert all(connector.provenance == "libreoffice_uno" for connector in resolved)


def test_libreoffice_backend_extracts_chart_from_sample() -> None:
    backend = LibreOfficeRichBackend(
        Path("sample/basic/sample.xlsx"),
        session_factory=_dummy_session_factory,
    )
    chart_data = backend.extract_charts(mode="libreoffice")
    charts = chart_data["Sheet1"]
    assert len(charts) >= 1
    chart = charts[0]
    assert chart.title == "売上データ"
    assert len(chart.series) == 3
    assert chart.l is not None and chart.t is not None
    assert chart.w is not None and chart.h is not None
    assert chart.provenance == "libreoffice_uno"
    assert chart.approximation_level == "partial"


def test_ooxml_connector_tail_end_maps_to_begin_arrow_style() -> None:
    node = ElementTree.fromstring(
        """
        <xdr:cxnSp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xdr:nvCxnSpPr>
            <xdr:cNvPr id="7" name="Connector 1" />
            <xdr:cNvCxnSpPr />
          </xdr:nvCxnSpPr>
          <xdr:spPr>
            <a:xfrm>
              <a:off x="0" y="0" />
              <a:ext cx="12700" cy="12700" />
            </a:xfrm>
            <a:ln>
              <a:tailEnd type="triangle" />
            </a:ln>
          </xdr:spPr>
        </xdr:cxnSp>
        """
    )
    connector = _parse_connector_node(node)
    assert connector is not None
    assert connector.begin_arrow_style == 2
    assert connector.end_arrow_style is None


def test_ooxml_connector_head_end_maps_to_end_arrow_style() -> None:
    node = ElementTree.fromstring(
        """
        <xdr:cxnSp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xdr:nvCxnSpPr>
            <xdr:cNvPr id="7" name="Connector 1" />
            <xdr:cNvCxnSpPr />
          </xdr:nvCxnSpPr>
          <xdr:spPr>
            <a:xfrm>
              <a:off x="0" y="0" />
              <a:ext cx="12700" cy="12700" />
            </a:xfrm>
            <a:ln>
              <a:headEnd type="triangle" />
            </a:ln>
          </xdr:spPr>
        </xdr:cxnSp>
        """
    )
    connector = _parse_connector_node(node)
    assert connector is not None
    assert connector.begin_arrow_style is None
    assert connector.end_arrow_style == 2


def test_libreoffice_session_cleans_temp_profile_on_enter_failure(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    removed_paths: list[Path] = []
    created_dir = tmp_path / "lo-profile"
    soffice_path = tmp_path / "soffice.exe"
    soffice_path.write_text("", encoding="utf-8")

    def _fake_mkdtemp(*, prefix: str, temp_dir: str | None = None) -> str:
        _ = prefix
        _ = temp_dir
        created_dir.mkdir(parents=True, exist_ok=True)
        return str(created_dir)

    def _fake_rmtree(path: Path | str, *, ignore_errors: bool) -> None:
        _ = ignore_errors
        removed_paths.append(Path(path))

    def _fake_run(*_args: object, **_kwargs: object) -> object:
        raise subprocess.TimeoutExpired(cmd="soffice --version", timeout=1.0)

    monkeypatch.setattr(
        "exstruct.core.libreoffice.mkdtemp",
        cast(Callable[..., str], _fake_mkdtemp),
    )
    monkeypatch.setattr("exstruct.core.libreoffice.shutil.rmtree", _fake_rmtree)
    monkeypatch.setattr("exstruct.core.libreoffice.subprocess.run", _fake_run)

    session = LibreOfficeSession(
        LibreOfficeSessionConfig(
            soffice_path=soffice_path,
            startup_timeout_sec=1.0,
            exec_timeout_sec=1.0,
            profile_root=None,
        )
    )

    with pytest.raises(LibreOfficeUnavailableError):
        session.__enter__()

    assert removed_paths == [created_dir]
    assert session._temp_profile_dir is None
