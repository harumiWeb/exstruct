"""Tests for the bundled LibreOffice bridge helper."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
from types import ModuleType, SimpleNamespace
from typing import Any, cast

from _pytest.monkeypatch import MonkeyPatch
import pytest


def _load_bridge_module(monkeypatch: MonkeyPatch) -> ModuleType:
    """Import the bridge module with a stub UNO runtime."""

    class _PropertyValue:
        def __init__(self) -> None:
            self.Name = ""
            self.Value = None

    uno_module = ModuleType("uno")
    uno_module.getComponentContext = lambda: SimpleNamespace(ServiceManager=None)  # type: ignore[attr-defined]
    uno_module.systemPathToFileUrl = (  # type: ignore[attr-defined]
        lambda path: "file:///" + path.replace("\\", "/")
    )

    beans_module = ModuleType("com.sun.star.beans")
    beans_module.PropertyValue = _PropertyValue  # type: ignore[attr-defined]

    for name in ("uno", "com", "com.sun", "com.sun.star", "com.sun.star.beans"):
        sys.modules.pop(name, None)
    monkeypatch.setitem(sys.modules, "uno", uno_module)
    monkeypatch.setitem(sys.modules, "com", ModuleType("com"))
    monkeypatch.setitem(sys.modules, "com.sun", ModuleType("com.sun"))
    monkeypatch.setitem(sys.modules, "com.sun.star", ModuleType("com.sun.star"))
    monkeypatch.setitem(sys.modules, "com.sun.star.beans", beans_module)
    module_path = (
        Path(__file__).resolve().parents[2]
        / "src"
        / "exstruct"
        / "core"
        / "_libreoffice_bridge.py"
    )
    module_name = "exstruct.core._libreoffice_bridge"
    sys.modules.pop(module_name, None)
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


class _FakeResolver:
    def __init__(self, results: list[object]) -> None:
        self._results = list(results)
        self.calls: list[str] = []

    def resolve(self, connection: str) -> object:
        self.calls.append(connection)
        result = self._results.pop(0)
        if isinstance(result, Exception):
            raise result
        return result


class _FakeServiceManager:
    def __init__(self, resolver: _FakeResolver) -> None:
        self._resolver = resolver

    def createInstanceWithContext(self, service_name: str, context: object) -> object:  # noqa: N802
        _ = context
        assert service_name == "com.sun.star.bridge.UnoUrlResolver"
        return self._resolver


class _FakeDrawPage:
    def __init__(self, items: list[object]) -> None:
        self._items = items

    def getCount(self) -> int:  # noqa: N802
        return len(self._items)

    def getByIndex(self, index: int) -> object:  # noqa: N802
        return self._items[index]


class _FakeNamedContainer:
    def __init__(self, names: list[str]) -> None:
        self._names = names

    def getElementNames(self) -> list[str]:  # noqa: N802
        return list(self._names)


class _FakeSheet:
    def __init__(self, *, chart_names: list[str], draw_items: list[object]) -> None:
        self._chart_names = chart_names
        self._draw_items = draw_items

    def getCharts(self) -> _FakeNamedContainer:  # noqa: N802
        return _FakeNamedContainer(self._chart_names)

    def getDrawPage(self) -> _FakeDrawPage:  # noqa: N802
        return _FakeDrawPage(self._draw_items)


class _FakeSheets:
    def __init__(self, mapping: dict[str, _FakeSheet]) -> None:
        self._mapping = mapping

    def getElementNames(self) -> list[str]:  # noqa: N802
        return list(self._mapping)

    def getByName(self, name: str) -> object:  # noqa: N802
        return self._mapping[name]


class _FakeDocument:
    def __init__(self, mapping: dict[str, _FakeSheet]) -> None:
        self._mapping = mapping
        self.closed = False
        self.disposed = False

    def getSheets(self) -> _FakeSheets:  # noqa: N802
        return _FakeSheets(self._mapping)

    def close(self, deliver_ownership: bool) -> None:
        _ = deliver_ownership
        self.closed = True

    def dispose(self) -> None:
        self.disposed = True


def test_bridge_parse_args_accepts_probe_and_handshake(
    monkeypatch: MonkeyPatch,
) -> None:
    module = _load_bridge_module(monkeypatch)

    monkeypatch.setattr(
        sys,
        "argv",
        ["bridge", "--probe"],
    )
    probe_args = module._parse_args()
    assert probe_args.probe is True

    monkeypatch.setattr(
        sys,
        "argv",
        [
            "bridge",
            "--handshake",
            "--host",
            "127.0.0.1",
            "--port",
            "2002",
            "--connect-timeout",
            "1.5",
        ],
    )
    handshake_args = module._parse_args()
    assert handshake_args.handshake is True
    assert handshake_args.file is None
    assert handshake_args.connect_timeout == 1.5


def test_bridge_main_returns_after_handshake(monkeypatch: MonkeyPatch) -> None:
    module = _load_bridge_module(monkeypatch)
    calls: dict[str, object] = {}

    def _fake_resolve_context(host: str, port: int, *, timeout_sec: float) -> object:
        calls["host"] = host
        calls["port"] = port
        calls["timeout_sec"] = timeout_sec
        return SimpleNamespace(ServiceManager=None)

    monkeypatch.setattr(module, "_resolve_context", _fake_resolve_context)
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "bridge",
            "--handshake",
            "--host",
            "127.0.0.1",
            "--port",
            "2002",
            "--connect-timeout",
            "2.5",
        ],
    )

    assert module.main() == 0
    assert calls == {"host": "127.0.0.1", "port": 2002, "timeout_sec": 2.5}


def test_bridge_resolve_context_retries_until_success(
    monkeypatch: MonkeyPatch,
) -> None:
    module = _load_bridge_module(monkeypatch)
    resolved_context = SimpleNamespace(ServiceManager="remote")
    resolver = _FakeResolver([RuntimeError("not ready"), resolved_context])
    local_context = SimpleNamespace(ServiceManager=_FakeServiceManager(resolver))

    monkeypatch.setattr(module.uno, "getComponentContext", lambda: local_context)
    monkeypatch.setattr(module.time, "sleep", lambda _seconds: None)

    context = module._resolve_context("127.0.0.1", 2002, timeout_sec=0.5)

    assert context is resolved_context
    assert resolver.calls == [
        "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext",
        "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext",
    ]


def test_bridge_load_document_builds_hidden_readonly_props(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    module = _load_bridge_module(monkeypatch)
    workbook_path = tmp_path / "book.xlsx"
    workbook_path.write_text("", encoding="utf-8")
    captured: dict[str, object] = {}

    class _Desktop:
        def loadComponentFromURL(  # noqa: N802
            self,
            url: str,
            target: str,
            search_flags: int,
            properties: tuple[object, ...],
        ) -> object:
            captured["url"] = url
            captured["target"] = target
            captured["search_flags"] = search_flags
            captured["properties"] = properties
            return SimpleNamespace(name="doc")

    document = module._load_document(_Desktop(), workbook_path)

    assert document.name == "doc"
    assert captured["url"] == f"file:///{workbook_path.resolve().as_posix()}"
    assert captured["target"] == "_blank"
    assert captured["search_flags"] == 0
    properties = list(cast(tuple[Any, ...], captured["properties"]))
    assert [(prop.Name, prop.Value) for prop in properties] == [
        ("Hidden", True),
        ("ReadOnly", True),
    ]


def test_bridge_extracts_chart_and_draw_page_payloads(
    monkeypatch: MonkeyPatch,
) -> None:
    module = _load_bridge_module(monkeypatch)
    chart_shape = SimpleNamespace(
        Name="Chart 1",
        PersistName="Object 1",
        Position=SimpleNamespace(X=1270, Y=2540),
        Size=SimpleNamespace(Width=2540, Height=5080),
        getShapeType=lambda: "com.sun.star.drawing.OLE2Shape",
    )
    ignored_chart_shape = SimpleNamespace(
        Name="Other",
        PersistName="Other 1",
        Position=SimpleNamespace(X=0, Y=0),
        Size=SimpleNamespace(Width=0, Height=0),
        getShapeType=lambda: "com.sun.star.drawing.OLE2Shape",
    )
    connector_shape = SimpleNamespace(
        Name="",
        String="flow",
        Position=SimpleNamespace(X=2540, Y=3810),
        Size=SimpleNamespace(Width=1270, Height=1270),
        RotateAngle=9000,
        StartShape=SimpleNamespace(Name="Start"),
        EndShape=SimpleNamespace(Name="End"),
        getShapeType=lambda: "com.sun.star.drawing.ConnectorShape",
    )
    broken_shape = SimpleNamespace(
        Position=SimpleNamespace(X=None, Y=None),
        Size=SimpleNamespace(Width=None, Height=None),
        RotateAngle=None,
        StartShape=None,
        EndShape=None,
        getShapeType=lambda: (_ for _ in ()).throw(RuntimeError("boom")),
    )
    document = _FakeDocument(
        {
            "Sheet1": _FakeSheet(
                chart_names=["Object 1"],
                draw_items=[
                    chart_shape,
                    ignored_chart_shape,
                    connector_shape,
                    broken_shape,
                ],
            )
        }
    )

    chart_payload = module._extract_chart_payload(document)
    draw_payload = module._extract_draw_page_payload(document)

    assert chart_payload == {
        "sheets": {
            "Sheet1": [
                {
                    "name": "Chart 1",
                    "persist_name": "Object 1",
                    "left": 36,
                    "top": 72,
                    "width": 72,
                    "height": 144,
                }
            ]
        }
    }
    assert draw_payload == {
        "sheets": {
            "Sheet1": [
                {
                    "name": "Shape 3",
                    "shape_type": "com.sun.star.drawing.ConnectorShape",
                    "text": "flow",
                    "left": 72,
                    "top": 108,
                    "width": 36,
                    "height": 36,
                    "rotation": 90.0,
                    "is_connector": True,
                    "start_shape_name": "Start",
                    "end_shape_name": "End",
                },
                {
                    "name": "Shape 4",
                    "shape_type": None,
                    "text": "",
                    "left": None,
                    "top": None,
                    "width": None,
                    "height": None,
                    "rotation": None,
                    "is_connector": False,
                    "start_shape_name": None,
                    "end_shape_name": None,
                },
            ]
        }
    }


def test_bridge_close_document_falls_back_to_dispose(
    monkeypatch: MonkeyPatch,
) -> None:
    module = _load_bridge_module(monkeypatch)

    class _Document:
        def __init__(self) -> None:
            self.dispose_called = False

        def close(self, deliver_ownership: bool) -> None:
            _ = deliver_ownership
            raise RuntimeError("close failed")

        def dispose(self) -> None:
            self.dispose_called = True

    document = _Document()
    module._close_document(document)

    assert document.dispose_called is True


def test_bridge_safe_shape_name_and_numeric_helpers(
    monkeypatch: MonkeyPatch,
) -> None:
    module = _load_bridge_module(monkeypatch)

    assert (
        module._safe_shape_name(
            SimpleNamespace(StartShape=SimpleNamespace(Name="A")), "StartShape"
        )
        == "A"
    )
    assert (
        module._safe_shape_name(SimpleNamespace(StartShape=None), "StartShape") is None
    )
    assert module._safe_shape_name(SimpleNamespace(), "StartShape") is None
    assert module._hmm_to_points(2540) == 72
    assert module._hmm_to_points("nope") is None
    assert module._rotation_to_degrees(4500) == 45.0
    assert module._rotation_to_degrees(None) is None


def test_bridge_parse_args_rejects_missing_host_port(monkeypatch: MonkeyPatch) -> None:
    module = _load_bridge_module(monkeypatch)
    monkeypatch.setattr(sys, "argv", ["bridge", "--handshake"])

    with pytest.raises(SystemExit):
        module._parse_args()
