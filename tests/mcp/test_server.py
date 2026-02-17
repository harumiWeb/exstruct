from __future__ import annotations

from collections.abc import Awaitable, Callable
import importlib
import logging
from pathlib import Path
from typing import Any, cast

import pytest

from exstruct.mcp import server
from exstruct.mcp.extract_runner import OnConflictPolicy
from exstruct.mcp.io import PathPolicy
from exstruct.mcp.tools import (
    ExtractToolInput,
    ExtractToolOutput,
    PatchToolInput,
    PatchToolOutput,
    ReadCellsToolInput,
    ReadCellsToolOutput,
    ReadFormulasToolInput,
    ReadFormulasToolOutput,
    ReadJsonChunkToolInput,
    ReadJsonChunkToolOutput,
    ReadRangeToolInput,
    ReadRangeToolOutput,
    ValidateInputToolInput,
    ValidateInputToolOutput,
)

anyio: Any = pytest.importorskip("anyio")

ToolFunc = Callable[..., object] | Callable[..., Awaitable[object]]


class DummyApp:
    def __init__(self) -> None:
        self.tools: dict[str, ToolFunc] = {}

    def tool(self, *, name: str) -> Callable[[ToolFunc], ToolFunc]:
        def decorator(func: ToolFunc) -> ToolFunc:
            self.tools[name] = func
            return func

        return decorator


async def _call_async(
    func: Callable[..., Awaitable[object]],
    kwargs: dict[str, object],
) -> object:
    return await func(**kwargs)


def test_parse_args_defaults(tmp_path: Path) -> None:
    config = server._parse_args(["--root", str(tmp_path)])
    assert config.root == tmp_path
    assert config.deny_globs == []
    assert config.log_level == "INFO"
    assert config.log_file is None
    assert config.on_conflict == "overwrite"
    assert config.warmup is False


def test_parse_args_with_options(tmp_path: Path) -> None:
    log_file = tmp_path / "log.txt"
    config = server._parse_args(
        [
            "--root",
            str(tmp_path),
            "--deny-glob",
            "**/*.tmp",
            "--deny-glob",
            "**/*.secret",
            "--log-level",
            "DEBUG",
            "--log-file",
            str(log_file),
            "--on-conflict",
            "rename",
            "--warmup",
        ]
    )
    assert config.deny_globs == ["**/*.tmp", "**/*.secret"]
    assert config.log_level == "DEBUG"
    assert config.log_file == log_file
    assert config.on_conflict == "rename"
    assert config.warmup is True


def test_import_mcp_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    def _raise(_: str) -> None:
        raise ModuleNotFoundError("mcp")

    monkeypatch.setattr(importlib, "import_module", _raise)
    with pytest.raises(RuntimeError):
        server._import_mcp()


def test_coerce_filter() -> None:
    assert server._coerce_filter(None) is None
    assert server._coerce_filter({"a": 1}) == {"a": 1}


def test_coerce_patch_ops_accepts_object_and_json_string() -> None:
    result = server._coerce_patch_ops(
        [
            {"op": "add_sheet", "sheet": "New"},
            '{"op":"set_value","sheet":"New","cell":"A1","value":"x"}',
        ]
    )
    assert result == [
        {"op": "add_sheet", "sheet": "New"},
        {"op": "set_value", "sheet": "New", "cell": "A1", "value": "x"},
    ]


def test_coerce_patch_ops_rejects_invalid_json_string() -> None:
    with pytest.raises(
        ValueError, match=r"Invalid patch operation at ops\[0\]: invalid JSON"
    ):
        server._coerce_patch_ops(["{invalid json}"])


def test_coerce_patch_ops_rejects_non_object_json_value() -> None:
    with pytest.raises(
        ValueError,
        match=r"Invalid patch operation at ops\[0\]: JSON value must be an object",
    ):
        server._coerce_patch_ops(['["not","object"]'])


def test_register_tools_uses_default_on_conflict(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    app = DummyApp()
    policy = PathPolicy(root=tmp_path)
    calls: dict[str, tuple[object, ...]] = {}

    def fake_run_extract_tool(
        payload: ExtractToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> ExtractToolOutput:
        calls["extract"] = (payload, policy, on_conflict)
        return ExtractToolOutput(out_path="out.json")

    def fake_run_read_json_chunk_tool(
        payload: ReadJsonChunkToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadJsonChunkToolOutput:
        calls["chunk"] = (payload, policy)
        return ReadJsonChunkToolOutput(chunk="{}")

    def fake_run_validate_input_tool(
        payload: ValidateInputToolInput,
        *,
        policy: PathPolicy,
    ) -> ValidateInputToolOutput:
        calls["validate"] = (payload, policy)
        return ValidateInputToolOutput(is_readable=True)

    def fake_run_patch_tool(
        payload: PatchToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> PatchToolOutput:
        calls["patch"] = (payload, policy, on_conflict)
        return PatchToolOutput(out_path="out.xlsx", patch_diff=[])

    async def fake_run_sync(func: Callable[[], object]) -> object:
        return func()

    monkeypatch.setattr(server, "run_extract_tool", fake_run_extract_tool)
    monkeypatch.setattr(
        server, "run_read_json_chunk_tool", fake_run_read_json_chunk_tool
    )
    monkeypatch.setattr(server, "run_validate_input_tool", fake_run_validate_input_tool)
    monkeypatch.setattr(server, "run_patch_tool", fake_run_patch_tool)
    monkeypatch.setattr(anyio.to_thread, "run_sync", fake_run_sync)

    server._register_tools(app, policy, default_on_conflict="rename")

    extract_tool = cast(Callable[..., Awaitable[object]], app.tools["exstruct_extract"])
    anyio.run(_call_async, extract_tool, {"xlsx_path": "in.xlsx"})
    read_chunk_tool = cast(
        Callable[..., Awaitable[object]], app.tools["exstruct_read_json_chunk"]
    )
    anyio.run(
        _call_async,
        read_chunk_tool,
        {"out_path": "out.json", "filter": {"rows": [1, 2]}},
    )
    validate_tool = cast(
        Callable[..., Awaitable[object]], app.tools["exstruct_validate_input"]
    )
    anyio.run(_call_async, validate_tool, {"xlsx_path": "in.xlsx"})
    patch_tool = cast(Callable[..., Awaitable[object]], app.tools["exstruct_patch"])
    anyio.run(
        _call_async,
        patch_tool,
        {"xlsx_path": "in.xlsx", "ops": [{"op": "add_sheet", "sheet": "New"}]},
    )

    assert calls["extract"][2] == "rename"
    chunk_call = cast(tuple[ReadJsonChunkToolInput, PathPolicy], calls["chunk"])
    assert chunk_call[0].filter is not None
    assert calls["patch"][2] == "rename"
    patch_call = cast(
        tuple[PatchToolInput, PathPolicy, OnConflictPolicy], calls["patch"]
    )
    assert patch_call[0].dry_run is False
    assert patch_call[0].return_inverse_ops is False
    assert patch_call[0].preflight_formula_check is False


def test_register_tools_passes_read_tool_arguments(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    app = DummyApp()
    policy = PathPolicy(root=tmp_path)
    calls: dict[str, tuple[object, ...]] = {}

    def fake_run_extract_tool(
        payload: ExtractToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> ExtractToolOutput:
        return ExtractToolOutput(out_path="out.json")

    def fake_run_read_json_chunk_tool(
        payload: ReadJsonChunkToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadJsonChunkToolOutput:
        return ReadJsonChunkToolOutput(chunk="{}")

    def fake_run_read_range_tool(
        payload: ReadRangeToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadRangeToolOutput:
        calls["range"] = (payload, policy)
        return ReadRangeToolOutput(
            book_name="book",
            sheet_name="Sheet1",
            range="A1:B2",
            cells=[],
        )

    def fake_run_read_cells_tool(
        payload: ReadCellsToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadCellsToolOutput:
        calls["cells"] = (payload, policy)
        return ReadCellsToolOutput(book_name="book", sheet_name="Sheet1", cells=[])

    def fake_run_read_formulas_tool(
        payload: ReadFormulasToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadFormulasToolOutput:
        calls["formulas"] = (payload, policy)
        return ReadFormulasToolOutput(
            book_name="book", sheet_name="Sheet1", formulas=[]
        )

    def fake_run_validate_input_tool(
        payload: ValidateInputToolInput,
        *,
        policy: PathPolicy,
    ) -> ValidateInputToolOutput:
        return ValidateInputToolOutput(is_readable=True)

    def fake_run_patch_tool(
        payload: PatchToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> PatchToolOutput:
        return PatchToolOutput(out_path="out.xlsx", patch_diff=[])

    async def fake_run_sync(func: Callable[[], object]) -> object:
        return func()

    monkeypatch.setattr(server, "run_extract_tool", fake_run_extract_tool)
    monkeypatch.setattr(
        server, "run_read_json_chunk_tool", fake_run_read_json_chunk_tool
    )
    monkeypatch.setattr(server, "run_read_range_tool", fake_run_read_range_tool)
    monkeypatch.setattr(server, "run_read_cells_tool", fake_run_read_cells_tool)
    monkeypatch.setattr(server, "run_read_formulas_tool", fake_run_read_formulas_tool)
    monkeypatch.setattr(server, "run_validate_input_tool", fake_run_validate_input_tool)
    monkeypatch.setattr(server, "run_patch_tool", fake_run_patch_tool)
    monkeypatch.setattr(anyio.to_thread, "run_sync", fake_run_sync)

    server._register_tools(app, policy, default_on_conflict="overwrite")

    read_range_tool = cast(
        Callable[..., Awaitable[object]], app.tools["exstruct_read_range"]
    )
    anyio.run(
        _call_async,
        read_range_tool,
        {"out_path": "out.json", "range": "A1:B2", "include_formulas": True},
    )
    read_cells_tool = cast(
        Callable[..., Awaitable[object]], app.tools["exstruct_read_cells"]
    )
    anyio.run(
        _call_async,
        read_cells_tool,
        {"out_path": "out.json", "addresses": ["J98", "J124"]},
    )
    read_formulas_tool = cast(
        Callable[..., Awaitable[object]], app.tools["exstruct_read_formulas"]
    )
    anyio.run(
        _call_async,
        read_formulas_tool,
        {"out_path": "out.json", "range": "J2:J201", "include_values": True},
    )

    range_call = cast(tuple[ReadRangeToolInput, PathPolicy], calls["range"])
    assert range_call[0].range == "A1:B2"
    assert range_call[0].include_formulas is True
    cells_call = cast(tuple[ReadCellsToolInput, PathPolicy], calls["cells"])
    assert cells_call[0].addresses == ["J98", "J124"]
    formulas_call = cast(tuple[ReadFormulasToolInput, PathPolicy], calls["formulas"])
    assert formulas_call[0].range == "J2:J201"
    assert formulas_call[0].include_values is True


def test_register_tools_accepts_patch_ops_json_strings(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    app = DummyApp()
    policy = PathPolicy(root=tmp_path)
    calls: dict[str, tuple[object, ...]] = {}

    def fake_run_extract_tool(
        payload: ExtractToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> ExtractToolOutput:
        return ExtractToolOutput(out_path="out.json")

    def fake_run_read_json_chunk_tool(
        payload: ReadJsonChunkToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadJsonChunkToolOutput:
        return ReadJsonChunkToolOutput(chunk="{}")

    def fake_run_validate_input_tool(
        payload: ValidateInputToolInput,
        *,
        policy: PathPolicy,
    ) -> ValidateInputToolOutput:
        return ValidateInputToolOutput(is_readable=True)

    def fake_run_patch_tool(
        payload: PatchToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> PatchToolOutput:
        calls["patch"] = (payload, policy, on_conflict)
        return PatchToolOutput(out_path="out.xlsx", patch_diff=[])

    async def fake_run_sync(func: Callable[[], object]) -> object:
        return func()

    monkeypatch.setattr(server, "run_extract_tool", fake_run_extract_tool)
    monkeypatch.setattr(
        server, "run_read_json_chunk_tool", fake_run_read_json_chunk_tool
    )
    monkeypatch.setattr(server, "run_validate_input_tool", fake_run_validate_input_tool)
    monkeypatch.setattr(server, "run_patch_tool", fake_run_patch_tool)
    monkeypatch.setattr(anyio.to_thread, "run_sync", fake_run_sync)

    server._register_tools(app, policy, default_on_conflict="overwrite")
    patch_tool = cast(Callable[..., Awaitable[object]], app.tools["exstruct_patch"])
    anyio.run(
        _call_async,
        patch_tool,
        {
            "xlsx_path": "in.xlsx",
            "ops": ['{"op":"add_sheet","sheet":"New"}'],
        },
    )
    patch_call = cast(
        tuple[PatchToolInput, PathPolicy, OnConflictPolicy], calls["patch"]
    )
    assert patch_call[0].ops[0].op == "add_sheet"
    assert patch_call[0].ops[0].sheet == "New"


def test_register_tools_rejects_invalid_patch_ops_json_strings(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    app = DummyApp()
    policy = PathPolicy(root=tmp_path)

    def fake_run_extract_tool(
        payload: ExtractToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> ExtractToolOutput:
        return ExtractToolOutput(out_path="out.json")

    def fake_run_read_json_chunk_tool(
        payload: ReadJsonChunkToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadJsonChunkToolOutput:
        return ReadJsonChunkToolOutput(chunk="{}")

    def fake_run_validate_input_tool(
        payload: ValidateInputToolInput,
        *,
        policy: PathPolicy,
    ) -> ValidateInputToolOutput:
        return ValidateInputToolOutput(is_readable=True)

    def fake_run_patch_tool(
        payload: PatchToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> PatchToolOutput:
        return PatchToolOutput(out_path="out.xlsx", patch_diff=[])

    async def fake_run_sync(func: Callable[[], object]) -> object:
        return func()

    monkeypatch.setattr(server, "run_extract_tool", fake_run_extract_tool)
    monkeypatch.setattr(
        server, "run_read_json_chunk_tool", fake_run_read_json_chunk_tool
    )
    monkeypatch.setattr(server, "run_validate_input_tool", fake_run_validate_input_tool)
    monkeypatch.setattr(server, "run_patch_tool", fake_run_patch_tool)
    monkeypatch.setattr(anyio.to_thread, "run_sync", fake_run_sync)

    server._register_tools(app, policy, default_on_conflict="overwrite")
    patch_tool = cast(Callable[..., Awaitable[object]], app.tools["exstruct_patch"])

    with pytest.raises(
        ValueError, match=r"Invalid patch operation at ops\[0\]: invalid JSON"
    ):
        anyio.run(
            _call_async,
            patch_tool,
            {"xlsx_path": "in.xlsx", "ops": ['{"op":"set_value"']},
        )


def test_register_tools_passes_patch_default_on_conflict(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    app = DummyApp()
    policy = PathPolicy(root=tmp_path)
    calls: dict[str, tuple[object, ...]] = {}

    def fake_run_extract_tool(
        payload: ExtractToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> ExtractToolOutput:
        return ExtractToolOutput(out_path="out.json")

    def fake_run_read_json_chunk_tool(
        payload: ReadJsonChunkToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadJsonChunkToolOutput:
        return ReadJsonChunkToolOutput(chunk="{}")

    def fake_run_validate_input_tool(
        payload: ValidateInputToolInput,
        *,
        policy: PathPolicy,
    ) -> ValidateInputToolOutput:
        return ValidateInputToolOutput(is_readable=True)

    def fake_run_patch_tool(
        payload: PatchToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> PatchToolOutput:
        calls["patch"] = (payload, policy, on_conflict)
        return PatchToolOutput(out_path="out.xlsx", patch_diff=[])

    async def fake_run_sync(func: Callable[[], object]) -> object:
        return func()

    monkeypatch.setattr(server, "run_extract_tool", fake_run_extract_tool)
    monkeypatch.setattr(
        server, "run_read_json_chunk_tool", fake_run_read_json_chunk_tool
    )
    monkeypatch.setattr(server, "run_validate_input_tool", fake_run_validate_input_tool)
    monkeypatch.setattr(server, "run_patch_tool", fake_run_patch_tool)
    monkeypatch.setattr(anyio.to_thread, "run_sync", fake_run_sync)

    server._register_tools(app, policy, default_on_conflict="overwrite")
    patch_tool = cast(Callable[..., Awaitable[object]], app.tools["exstruct_patch"])
    anyio.run(
        _call_async,
        patch_tool,
        {"xlsx_path": "in.xlsx", "ops": [{"op": "add_sheet", "sheet": "New"}]},
    )

    assert calls["patch"][2] == "overwrite"


def test_register_tools_passes_patch_extended_flags(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    app = DummyApp()
    policy = PathPolicy(root=tmp_path)
    calls: dict[str, tuple[object, ...]] = {}

    def fake_run_extract_tool(
        payload: ExtractToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> ExtractToolOutput:
        return ExtractToolOutput(out_path="out.json")

    def fake_run_read_json_chunk_tool(
        payload: ReadJsonChunkToolInput,
        *,
        policy: PathPolicy,
    ) -> ReadJsonChunkToolOutput:
        return ReadJsonChunkToolOutput(chunk="{}")

    def fake_run_validate_input_tool(
        payload: ValidateInputToolInput,
        *,
        policy: PathPolicy,
    ) -> ValidateInputToolOutput:
        return ValidateInputToolOutput(is_readable=True)

    def fake_run_patch_tool(
        payload: PatchToolInput,
        *,
        policy: PathPolicy,
        on_conflict: OnConflictPolicy,
    ) -> PatchToolOutput:
        calls["patch"] = (payload, policy, on_conflict)
        return PatchToolOutput(out_path="out.xlsx", patch_diff=[])

    async def fake_run_sync(func: Callable[[], object]) -> object:
        return func()

    monkeypatch.setattr(server, "run_extract_tool", fake_run_extract_tool)
    monkeypatch.setattr(
        server, "run_read_json_chunk_tool", fake_run_read_json_chunk_tool
    )
    monkeypatch.setattr(server, "run_validate_input_tool", fake_run_validate_input_tool)
    monkeypatch.setattr(server, "run_patch_tool", fake_run_patch_tool)
    monkeypatch.setattr(anyio.to_thread, "run_sync", fake_run_sync)

    server._register_tools(app, policy, default_on_conflict="overwrite")
    patch_tool = cast(Callable[..., Awaitable[object]], app.tools["exstruct_patch"])
    anyio.run(
        _call_async,
        patch_tool,
        {
            "xlsx_path": "in.xlsx",
            "ops": [{"op": "add_sheet", "sheet": "New"}],
            "dry_run": True,
            "return_inverse_ops": True,
            "preflight_formula_check": True,
        },
    )
    patch_call = cast(
        tuple[PatchToolInput, PathPolicy, OnConflictPolicy], calls["patch"]
    )
    assert patch_call[0].dry_run is True
    assert patch_call[0].return_inverse_ops is True
    assert patch_call[0].preflight_formula_check is True


def test_run_server_sets_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    created: dict[str, object] = {}

    def fake_import() -> None:
        created["imported"] = True

    class _App:
        def run(self) -> None:
            created["ran"] = True

    def fake_create_app(policy: PathPolicy, *, on_conflict: OnConflictPolicy) -> _App:
        created["policy"] = policy
        created["on_conflict"] = on_conflict
        return _App()

    monkeypatch.setattr(server, "_import_mcp", fake_import)
    monkeypatch.setattr(server, "_create_app", fake_create_app)
    config = server.ServerConfig(root=tmp_path)
    server.run_server(config)
    assert created["imported"] is True
    assert created["ran"] is True
    assert created["on_conflict"] == "overwrite"


def test_configure_logging_with_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    log_file = tmp_path / "server.log"
    config = server.ServerConfig(root=tmp_path, log_file=log_file)
    captured: dict[str, object] = {}

    def _basic_config(**kwargs: object) -> None:
        captured.update(kwargs)

    monkeypatch.setattr(logging, "basicConfig", _basic_config)
    server._configure_logging(config)
    handlers = cast(list[logging.Handler], captured["handlers"])
    assert any(isinstance(handler, logging.FileHandler) for handler in handlers)


def test_warmup_exstruct_imports(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[str] = []

    def _record(name: str) -> object:
        calls.append(name)
        return object()

    monkeypatch.setattr(importlib, "import_module", _record)
    server._warmup_exstruct()
    assert "exstruct.core.cells" in calls
    assert "exstruct.core.integrate" in calls
