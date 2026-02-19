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
from exstruct.mcp.patch_runner import MakeRequest, PatchRequest, PatchResult
from exstruct.mcp.sheet_reader import (
    ReadCellsRequest,
    ReadCellsResult,
    ReadFormulasRequest,
    ReadFormulasResult,
    ReadRangeRequest,
    ReadRangeResult,
)
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


def test_run_read_range_tool_builds_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_read_range(
        request: ReadRangeRequest, *, policy: object | None = None
    ) -> ReadRangeResult:
        captured["request"] = request
        return ReadRangeResult(
            book_name="book",
            sheet_name="Sheet1",
            range="A1:B2",
            cells=[],
        )

    monkeypatch.setattr(tools, "read_range", _fake_read_range)
    payload = tools.ReadRangeToolInput(out_path="out.json", range="A1:B2")
    tools.run_read_range_tool(payload)
    request = captured["request"]
    assert isinstance(request, ReadRangeRequest)
    assert request.out_path == Path("out.json")
    assert request.range == "A1:B2"


def test_run_read_cells_tool_builds_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_read_cells(
        request: ReadCellsRequest, *, policy: object | None = None
    ) -> ReadCellsResult:
        captured["request"] = request
        return ReadCellsResult(book_name="book", sheet_name="Sheet1", cells=[])

    monkeypatch.setattr(tools, "read_cells", _fake_read_cells)
    payload = tools.ReadCellsToolInput(out_path="out.json", addresses=["A1", "B2"])
    tools.run_read_cells_tool(payload)
    request = captured["request"]
    assert isinstance(request, ReadCellsRequest)
    assert request.out_path == Path("out.json")
    assert request.addresses == ["A1", "B2"]


def test_run_read_formulas_tool_builds_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_read_formulas(
        request: ReadFormulasRequest, *, policy: object | None = None
    ) -> ReadFormulasResult:
        captured["request"] = request
        return ReadFormulasResult(book_name="book", sheet_name="Sheet1", formulas=[])

    monkeypatch.setattr(tools, "read_formulas", _fake_read_formulas)
    payload = tools.ReadFormulasToolInput(out_path="out.json", range="J2:J20")
    tools.run_read_formulas_tool(payload)
    request = captured["request"]
    assert isinstance(request, ReadFormulasRequest)
    assert request.out_path == Path("out.json")
    assert request.range == "J2:J20"


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


def test_run_patch_tool_builds_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_run_patch(
        request: PatchRequest, *, policy: object | None = None
    ) -> PatchResult:
        captured["request"] = request
        return PatchResult(out_path="out.xlsx", patch_diff=[], engine="openpyxl")

    monkeypatch.setattr(tools, "run_patch", _fake_run_patch)
    payload = tools.PatchToolInput(
        xlsx_path="input.xlsx",
        ops=[{"op": "add_sheet", "sheet": "New"}],
        dry_run=True,
        return_inverse_ops=True,
        preflight_formula_check=True,
    )
    tools.run_patch_tool(payload, on_conflict="rename")
    request = captured["request"]
    assert isinstance(request, PatchRequest)
    assert request.xlsx_path == Path("input.xlsx")
    assert request.on_conflict == "rename"
    assert request.dry_run is True
    assert request.return_inverse_ops is True
    assert request.preflight_formula_check is True
    assert request.backend == "auto"


def test_run_make_tool_builds_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def _fake_run_make(
        request: MakeRequest, *, policy: object | None = None
    ) -> PatchResult:
        captured["request"] = request
        return PatchResult(out_path="out.xlsx", patch_diff=[], engine="openpyxl")

    monkeypatch.setattr(tools, "run_make", _fake_run_make)
    payload = tools.MakeToolInput(
        out_path="output.xlsx",
        ops=[{"op": "add_sheet", "sheet": "New"}],
        dry_run=True,
        return_inverse_ops=True,
        preflight_formula_check=True,
    )
    tools.run_make_tool(payload, on_conflict="rename")
    request = captured["request"]
    assert isinstance(request, MakeRequest)
    assert request.out_path == Path("output.xlsx")
    assert request.on_conflict == "rename"
    assert request.dry_run is True
    assert request.return_inverse_ops is True
    assert request.preflight_formula_check is True
    assert request.backend == "auto"
