from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.mcp.patch.models import OpenpyxlEngineResult, PatchOp, PatchRequest
from exstruct.mcp.patch.ops.openpyxl_ops import apply_openpyxl_ops
from exstruct.mcp.patch.ops.xlwings_ops import apply_xlwings_ops


def test_apply_openpyxl_ops_delegates_to_legacy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import exstruct.mcp.patch.internal as legacy_runner

    expected = (("diff",), ("inverse",), ("issues",), ("warn",))

    def _fake_apply_ops_openpyxl(
        request: PatchRequest,
        input_path: Path,
        output_path: Path,
    ) -> tuple[tuple[str, ...], tuple[str, ...], tuple[str, ...], tuple[str, ...]]:
        return expected

    monkeypatch.setattr(legacy_runner, "_apply_ops_openpyxl", _fake_apply_ops_openpyxl)
    result = apply_openpyxl_ops(
        PatchRequest(
            xlsx_path=Path("input.xlsx"),
            ops=[PatchOp(op="add_sheet", sheet="Data")],
        ),
        Path("input.xlsx"),
        Path("output.xlsx"),
    )
    assert result == OpenpyxlEngineResult(
        patch_diff=["diff"],
        inverse_ops=["inverse"],
        formula_issues=["issues"],
        op_warnings=["warn"],
    )


def test_apply_xlwings_ops_delegates_to_legacy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import exstruct.mcp.patch.internal as legacy_runner

    expected = ("diff",)

    def _fake_apply_ops_xlwings(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> tuple[str, ...]:
        return expected

    monkeypatch.setattr(legacy_runner, "_apply_ops_xlwings", _fake_apply_ops_xlwings)
    result = apply_xlwings_ops(
        Path("input.xlsx"),
        Path("output.xlsx"),
        [PatchOp(op="add_sheet", sheet="Data")],
        auto_formula=False,
    )
    assert result == ["diff"]
