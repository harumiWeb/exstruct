from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.mcp.patch.engine.xlwings_engine import apply_xlwings_engine
from exstruct.mcp.patch.models import (
    FormulaIssue,
    OpenpyxlEngineResult,
    PatchDiffItem,
    PatchOp,
    PatchRequest,
    PatchValue,
)
from exstruct.mcp.patch.ops.openpyxl_ops import apply_openpyxl_ops
from exstruct.mcp.patch.ops.xlwings_ops import apply_xlwings_ops


def test_coerce_model_list_accepts_valid_items_and_skips_invalid() -> None:
    from exstruct.mcp.patch.ops import openpyxl_ops

    items: list[object] = [
        PatchOp(op="add_sheet", sheet="Data"),
        {"op": "add_sheet", "sheet": "Data2"},
        PatchValue(kind="value", value="x"),
        "invalid",
    ]

    coerced = openpyxl_ops._coerce_model_list(items, PatchOp)
    assert coerced == [
        PatchOp(op="add_sheet", sheet="Data"),
        PatchOp(op="add_sheet", sheet="Data2"),
    ]


def test_apply_openpyxl_ops_delegates_to_legacy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import exstruct.mcp.patch.internal as legacy_runner

    expected: tuple[
        tuple[dict[str, object], ...],
        tuple[dict[str, object], ...],
        tuple[dict[str, object], ...],
        tuple[str, ...],
    ] = (
        (
            {
                "op_index": 0,
                "op": "add_sheet",
                "sheet": "Data",
                "cell": "A1",
                "before": None,
                "after": {"kind": "sheet", "value": "created"},
                "status": "applied",
            },
        ),
        (
            {
                "op": "add_sheet",
                "sheet": "UndoData",
            },
        ),
        (
            {
                "level": "warning",
                "code": "div0_error",
                "sheet": "Data",
                "cell": "A1",
                "message": "formula warning",
            },
        ),
        ("warn",),
    )

    def _fake_apply_ops_openpyxl(
        request: PatchRequest,
        input_path: Path,
        output_path: Path,
    ) -> tuple[
        tuple[dict[str, object], ...],
        tuple[dict[str, object], ...],
        tuple[dict[str, object], ...],
        tuple[str, ...],
    ]:
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
        patch_diff=[
            PatchDiffItem(
                op_index=0,
                op="add_sheet",
                sheet="Data",
                cell="A1",
                before=None,
                after=PatchValue(kind="sheet", value="created"),
                status="applied",
            )
        ],
        inverse_ops=[PatchOp(op="add_sheet", sheet="UndoData")],
        formula_issues=[
            FormulaIssue(
                level="warning",
                code="div0_error",
                sheet="Data",
                cell="A1",
                message="formula warning",
            )
        ],
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


def test_apply_xlwings_engine_delegates_to_ops(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import exstruct.mcp.patch.engine.xlwings_engine as engine_module

    expected: list[object] = ["ok"]

    def _fake_apply_xlwings_ops(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        assert input_path == Path("input.xlsx")
        assert output_path == Path("output.xlsx")
        assert ops == [PatchOp(op="add_sheet", sheet="Data")]
        assert auto_formula is True
        return expected

    monkeypatch.setattr(engine_module, "apply_xlwings_ops", _fake_apply_xlwings_ops)

    result = apply_xlwings_engine(
        Path("input.xlsx"),
        Path("output.xlsx"),
        [PatchOp(op="add_sheet", sheet="Data")],
        True,
    )
    assert result == expected
