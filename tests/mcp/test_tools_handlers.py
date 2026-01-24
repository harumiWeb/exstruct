from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.mcp import tools
from exstruct.mcp.chunk_reader import (
    ReadJsonChunkFilter,
    ReadJsonChunkRequest,
    ReadJsonChunkResult,
)
from exstruct.mcp.extract_runner import ExtractRequest, ExtractResult
from exstruct.mcp.validate_input import ValidateInputRequest, ValidateInputResult


def test_run_extract_tool_prefers_payload_on_conflict(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_run_extract(
        request: ExtractRequest, *, policy: object | None = None
    ) -> ExtractResult:
        captured["request"] = request
        return ExtractResult(out_path="out.json")

    monkeypatch.setattr(tools, "run_extract", _fake_run_extract)
    payload = tools.ExtractToolInput(xlsx_path="input.xlsx", on_conflict="skip")
    tools.run_extract_tool(payload, on_conflict="rename")
    request = captured["request"]
    assert isinstance(request, ExtractRequest)
    assert request.on_conflict == "skip"


def test_run_extract_tool_uses_default_on_conflict(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_run_extract(
        request: ExtractRequest, *, policy: object | None = None
    ) -> ExtractResult:
        captured["request"] = request
        return ExtractResult(out_path="out.json")

    monkeypatch.setattr(tools, "run_extract", _fake_run_extract)
    payload = tools.ExtractToolInput(xlsx_path="input.xlsx", on_conflict=None)
    tools.run_extract_tool(payload, on_conflict="rename")
    request = captured["request"]
    assert isinstance(request, ExtractRequest)
    assert request.on_conflict == "rename"


def test_run_read_json_chunk_tool_builds_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_read_json_chunk(
        request: ReadJsonChunkRequest, *, policy: object | None = None
    ) -> ReadJsonChunkResult:
        captured["request"] = request
        return ReadJsonChunkResult(chunk="{}", next_cursor=None, warnings=[])

    monkeypatch.setattr(tools, "read_json_chunk", _fake_read_json_chunk)
    payload = tools.ReadJsonChunkToolInput(
        out_path="out.json", filter=ReadJsonChunkFilter(rows=(1, 2))
    )
    tools.run_read_json_chunk_tool(payload)
    request = captured["request"]
    assert isinstance(request, ReadJsonChunkRequest)
    assert request.out_path == Path("out.json")
    assert request.filter is not None


def test_run_validate_input_tool_builds_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_validate_input(
        request: ValidateInputRequest, *, policy: object | None = None
    ) -> ValidateInputResult:
        captured["request"] = request
        return ValidateInputResult(is_readable=True)

    monkeypatch.setattr(tools, "validate_input", _fake_validate_input)
    payload = tools.ValidateInputToolInput(xlsx_path="input.xlsx")
    tools.run_validate_input_tool(payload)
    request = captured["request"]
    assert isinstance(request, ValidateInputRequest)
    assert request.xlsx_path == Path("input.xlsx")
