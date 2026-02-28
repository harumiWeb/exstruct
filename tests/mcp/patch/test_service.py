from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
import pytest

from exstruct.cli.availability import ComAvailability
from exstruct.mcp.patch import (
    internal as patch_internal,
    runtime as patch_runtime,
    service,
)
from exstruct.mcp.patch.models import OpenpyxlEngineResult
from exstruct.mcp.patch_runner import MakeRequest, PatchOp, PatchRequest, PatchResult


def _create_workbook(path: Path) -> None:
    """Create a minimal workbook fixture for patch tests.

    Args:
        path: Target workbook path.
    """
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet["A1"] = "old"
    workbook.save(path)
    workbook.close()


def test_patch_runner_run_patch_delegates_to_service(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Verify patch_runner.run_patch delegates to patch.service.run_patch.

    Args:
        monkeypatch: Pytest monkeypatch fixture.
    """
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
    """Verify patch_runner.run_make delegates to patch.service.run_make.

    Args:
        monkeypatch: Pytest monkeypatch fixture.
    """
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


def test_service_run_patch_backend_auto_prefers_com(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify backend=auto uses COM when available.

    Args:
        tmp_path: Temporary directory fixture.
        monkeypatch: Pytest monkeypatch fixture.
    """
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    calls: dict[str, bool] = {}

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    def _fake_apply_xlwings_engine(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        calls["com"] = True
        return []

    monkeypatch.setattr(service, "apply_xlwings_engine", _fake_apply_xlwings_engine)
    result = service.run_patch(
        PatchRequest(
            xlsx_path=input_path,
            ops=[PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="new")],
            on_conflict="rename",
            backend="auto",
        )
    )
    assert result.error is None
    assert result.engine == "com"
    assert calls["com"] is True


def test_service_run_patch_backend_auto_fallbacks_to_openpyxl_on_com_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify backend=auto falls back to openpyxl when COM apply fails.

    Args:
        tmp_path: Temporary directory fixture.
        monkeypatch: Pytest monkeypatch fixture.
    """
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    def _raise_com_error(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        raise RuntimeError("boom")

    def _fake_apply_openpyxl_engine(
        request: PatchRequest,
        input_path: Path,
        output_path: Path,
    ) -> OpenpyxlEngineResult:
        return OpenpyxlEngineResult()

    monkeypatch.setattr(service, "apply_xlwings_engine", _raise_com_error)
    monkeypatch.setattr(service, "apply_openpyxl_engine", _fake_apply_openpyxl_engine)
    result = service.run_patch(
        PatchRequest(
            xlsx_path=input_path,
            ops=[PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="new")],
            on_conflict="rename",
            backend="auto",
        )
    )
    assert result.error is None
    assert result.engine == "openpyxl"
    assert any("falling back to openpyxl" in warning for warning in result.warnings)


def test_service_run_patch_backend_auto_fallbacks_to_openpyxl_on_com_patch_op_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify backend=auto falls back when COM path raises PatchOpError."""
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    def _raise_com_patch_error(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        detail = patch_internal.PatchErrorDetail(
            op_index=0,
            op="set_value",
            sheet="Sheet1",
            cell="A1",
            message="COM call failed.",
            error_code="com_runtime_error",
            raw_com_message="(-2147352567, 'Exception occurred.')",
        )
        raise patch_runtime.PatchOpError(detail)

    def _fake_apply_openpyxl_engine(
        request: PatchRequest,
        input_path: Path,
        output_path: Path,
    ) -> OpenpyxlEngineResult:
        return OpenpyxlEngineResult()

    monkeypatch.setattr(service, "apply_xlwings_engine", _raise_com_patch_error)
    monkeypatch.setattr(service, "apply_openpyxl_engine", _fake_apply_openpyxl_engine)
    result = service.run_patch(
        PatchRequest(
            xlsx_path=input_path,
            ops=[PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="new")],
            on_conflict="rename",
            backend="auto",
        )
    )
    assert result.error is None
    assert result.engine == "openpyxl"
    assert any("falling back to openpyxl" in warning for warning in result.warnings)


def test_service_run_patch_backend_com_does_not_fallback_on_com_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify backend=com propagates COM errors without fallback.

    Args:
        tmp_path: Temporary directory fixture.
        monkeypatch: Pytest monkeypatch fixture.
    """
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    def _raise_com_error(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        raise RuntimeError("boom")

    monkeypatch.setattr(service, "apply_xlwings_engine", _raise_com_error)
    with pytest.raises(RuntimeError, match=r"COM patch failed"):
        service.run_patch(
            PatchRequest(
                xlsx_path=input_path,
                ops=[PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="new")],
                on_conflict="rename",
                backend="com",
            )
        )


def test_service_run_patch_backend_com_uses_com_for_apply_table_style(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify backend=com executes apply_table_style on COM backend.

    Args:
        tmp_path: Temporary directory fixture.
        monkeypatch: Pytest monkeypatch fixture.
    """
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    calls: dict[str, bool] = {}

    def _fake_apply_xlwings_engine(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        calls["com"] = True
        return []

    def _fake_apply_openpyxl_engine(
        request: PatchRequest,
        input_path: Path,
        output_path: Path,
    ) -> OpenpyxlEngineResult:
        raise AssertionError(
            "openpyxl backend should not be called for apply_table_style"
        )

    monkeypatch.setattr(service, "apply_xlwings_engine", _fake_apply_xlwings_engine)
    monkeypatch.setattr(service, "apply_openpyxl_engine", _fake_apply_openpyxl_engine)
    result = service.run_patch(
        PatchRequest(
            xlsx_path=input_path,
            ops=[
                PatchOp(
                    op="apply_table_style",
                    sheet="Sheet1",
                    range="A1:B3",
                    style="TableStyleMedium2",
                    table_name="SalesTable",
                )
            ],
            on_conflict="rename",
            backend="com",
        )
    )
    assert result.error is None
    assert result.engine == "com"
    assert calls["com"] is True


def test_service_run_patch_backend_auto_prefers_com_for_apply_table_style(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify backend=auto prefers COM for apply_table_style when available.

    Args:
        tmp_path: Temporary directory fixture.
        monkeypatch: Pytest monkeypatch fixture.
    """
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    calls: dict[str, bool] = {}

    def _fake_apply_xlwings_engine(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        calls["com"] = True
        return []

    def _fake_apply_openpyxl_engine(
        request: PatchRequest,
        input_path: Path,
        output_path: Path,
    ) -> OpenpyxlEngineResult:
        raise AssertionError(
            "openpyxl backend should not be called for apply_table_style"
        )

    monkeypatch.setattr(service, "apply_xlwings_engine", _fake_apply_xlwings_engine)
    monkeypatch.setattr(service, "apply_openpyxl_engine", _fake_apply_openpyxl_engine)
    result = service.run_patch(
        PatchRequest(
            xlsx_path=input_path,
            ops=[
                PatchOp(
                    op="apply_table_style",
                    sheet="Sheet1",
                    range="A1:B3",
                    style="TableStyleMedium2",
                    table_name="SalesTable",
                )
            ],
            on_conflict="rename",
            backend="auto",
        )
    )
    assert result.error is None
    assert result.engine == "com"
    assert calls["com"] is True


def test_service_run_patch_allows_create_chart_with_apply_table_style_on_com(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify mixed chart/table ops run on COM backend."""
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    calls: dict[str, object] = {}

    def _fake_apply_xlwings_engine(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        calls["engine"] = "com"
        calls["ops"] = [op.op for op in ops]
        return []

    def _fake_apply_openpyxl_engine(
        request: PatchRequest,
        input_path: Path,
        output_path: Path,
    ) -> OpenpyxlEngineResult:
        raise AssertionError(
            "openpyxl backend should not be called for mixed create_chart/apply_table_style"
        )

    monkeypatch.setattr(service, "apply_xlwings_engine", _fake_apply_xlwings_engine)
    monkeypatch.setattr(service, "apply_openpyxl_engine", _fake_apply_openpyxl_engine)
    result = service.run_patch(
        PatchRequest(
            xlsx_path=input_path,
            ops=[
                PatchOp(
                    op="create_chart",
                    sheet="Sheet1",
                    chart_type="line",
                    data_range="A1:B2",
                    anchor_cell="D2",
                ),
                PatchOp(
                    op="apply_table_style",
                    sheet="Sheet1",
                    range="A1:B2",
                    style="TableStyleMedium2",
                ),
            ],
            on_conflict="rename",
            backend="auto",
        )
    )

    assert result.error is None
    assert result.engine == "com"
    assert calls["engine"] == "com"
    assert calls["ops"] == ["create_chart", "apply_table_style"]


def test_service_run_patch_allows_mixed_request_on_backend_com(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify backend=com runs mixed chart/table ops on COM."""
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    calls: dict[str, object] = {}

    def _fake_apply_xlwings_engine(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        calls["engine"] = "com"
        calls["ops"] = [op.op for op in ops]
        return []

    monkeypatch.setattr(service, "apply_xlwings_engine", _fake_apply_xlwings_engine)
    result = service.run_patch(
        PatchRequest(
            xlsx_path=input_path,
            ops=[
                PatchOp(
                    op="create_chart",
                    sheet="Sheet1",
                    chart_type="line",
                    data_range="A1:B2",
                    anchor_cell="D2",
                ),
                PatchOp(
                    op="apply_table_style",
                    sheet="Sheet1",
                    range="A1:B2",
                    style="TableStyleMedium2",
                ),
            ],
            on_conflict="rename",
            backend="com",
        )
    )

    assert result.error is None
    assert result.engine == "com"
    assert calls["engine"] == "com"
    assert calls["ops"] == ["create_chart", "apply_table_style"]


def test_service_run_patch_mixed_request_requires_com_when_auto_unavailable(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Verify mixed chart/table ops fail clearly when COM is unavailable."""
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        patch_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=False, reason="not available"),
    )

    with pytest.raises(
        ValueError,
        match=(
            r"create_chart \+ apply_table_style requests require "
            r"Windows Excel COM availability"
        ),
    ):
        service.run_patch(
            PatchRequest(
                xlsx_path=input_path,
                ops=[
                    PatchOp(
                        op="create_chart",
                        sheet="Sheet1",
                        chart_type="line",
                        data_range="A1:B2",
                        anchor_cell="D2",
                    ),
                    PatchOp(
                        op="apply_table_style",
                        sheet="Sheet1",
                        range="A1:B2",
                        style="TableStyleMedium2",
                    ),
                ],
                on_conflict="rename",
                backend="auto",
            )
        )
