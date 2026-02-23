from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.mcp.patch import service
from exstruct.mcp.patch_runner import MakeRequest, PatchOp, PatchRequest, PatchResult


def test_patch_runner_run_patch_delegates_to_service(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import exstruct.mcp.patch_runner as patch_runner

    expected = PatchResult(out_path="out.xlsx", patch_diff=[], engine="openpyxl")

    def _fake_run_patch(
        request: PatchRequest, *, policy: object | None = None
    ) -> PatchResult:
        return expected

    monkeypatch.setattr(service, "run_patch", _fake_run_patch)
    request = PatchRequest(
        xlsx_path=Path("input.xlsx"),
        ops=[PatchOp(op="add_sheet", sheet="Data")],
    )
    result = patch_runner.run_patch(request)
    assert result is expected


def test_patch_runner_run_make_delegates_to_service(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import exstruct.mcp.patch_runner as patch_runner

    expected = PatchResult(out_path="out.xlsx", patch_diff=[], engine="openpyxl")

    def _fake_run_make(
        request: MakeRequest, *, policy: object | None = None
    ) -> PatchResult:
        return expected

    monkeypatch.setattr(service, "run_make", _fake_run_make)
    request = MakeRequest(out_path=Path("output.xlsx"), ops=[])
    result = patch_runner.run_make(request)
    assert result is expected
