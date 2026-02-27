from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import TypeVar

from pydantic import BaseModel, ValidationError

from exstruct.mcp.io import PathPolicy

from . import runtime
from .engine.openpyxl_engine import apply_openpyxl_engine
from .engine.xlwings_engine import apply_xlwings_engine
from .models import (
    FormulaIssue,
    MakeRequest,
    PatchDiffItem,
    PatchErrorDetail,
    PatchOp,
    PatchRequest,
    PatchResult,
)
from .types import PatchOpType

TModel = TypeVar("TModel", bound=BaseModel)


def run_make(request: MakeRequest, *, policy: PathPolicy | None = None) -> PatchResult:
    """Create a new workbook and apply patch operations in one call."""
    resolved_output = runtime.resolve_make_output_path(request.out_path, policy=policy)
    runtime.ensure_supported_extension(resolved_output)
    runtime.validate_make_request_constraints(request, resolved_output)
    seed_path = runtime.build_make_seed_path(resolved_output)
    initial_sheet_name = runtime.resolve_make_initial_sheet_name(request)
    try:
        runtime.create_seed_workbook(
            seed_path,
            resolved_output.suffix.lower(),
            initial_sheet_name=initial_sheet_name,
        )
        patch_request = PatchRequest(
            xlsx_path=seed_path,
            ops=request.ops,
            sheet=request.sheet,
            out_dir=resolved_output.parent,
            out_name=resolved_output.name,
            on_conflict=request.on_conflict,
            auto_formula=request.auto_formula,
            dry_run=request.dry_run,
            return_inverse_ops=request.return_inverse_ops,
            preflight_formula_check=request.preflight_formula_check,
            backend=request.backend,
        )
        return run_patch(patch_request, policy=policy)
    finally:
        if seed_path.exists():
            seed_path.unlink()


def run_patch(
    request: PatchRequest, *, policy: PathPolicy | None = None
) -> PatchResult:
    """Run a patch operation and write the updated workbook."""
    resolved_input = runtime.resolve_input_path(request.xlsx_path, policy=policy)
    runtime.ensure_supported_extension(resolved_input)
    output_path = runtime.resolve_output_path(
        resolved_input,
        out_dir=request.out_dir,
        out_name=request.out_name,
        policy=policy,
    )
    warnings: list[str] = []
    runtime.append_large_ops_warning(warnings, request.ops)
    effective_request = _resolve_effective_request(request)
    if resolved_input.suffix.lower() == ".xls" and runtime.contains_design_ops(
        effective_request.ops
    ):
        raise ValueError(
            "Design operations are not supported for .xls files. Convert to .xlsx/.xlsm first."
        )
    com = runtime.get_com_availability()
    selected_engine = runtime.select_patch_engine(
        request=effective_request,
        input_path=resolved_input,
        com_available=com.available,
    )
    output_path, warning, skipped = runtime.apply_conflict_policy(
        output_path, effective_request.on_conflict
    )
    if warning:
        warnings.append(warning)
    if skipped and not effective_request.dry_run:
        return PatchResult(
            out_path=str(output_path),
            patch_diff=[],
            inverse_ops=[],
            formula_issues=[],
            warnings=warnings,
            engine=selected_engine,
        )
    if skipped and effective_request.dry_run:
        warnings.append(
            "Dry-run mode ignores on_conflict=skip and simulates patch without writing."
        )
    if (
        selected_engine == "openpyxl"
        and com.reason
        and effective_request.backend == "auto"
    ):
        warnings.append(f"COM unavailable: {com.reason}")
    if selected_engine == "openpyxl" and runtime.requires_openpyxl_backend(
        effective_request
    ):
        warnings.append("Using openpyxl backend due to patch request constraints.")

    runtime.ensure_output_dir(output_path)
    if selected_engine == "com":
        try:
            diff = apply_xlwings_engine(
                resolved_input,
                output_path,
                effective_request.ops,
                effective_request.auto_formula,
            )
            return PatchResult(
                out_path=str(output_path),
                patch_diff=_coerce_patch_diff_items(diff),
                inverse_ops=[],
                formula_issues=[],
                warnings=warnings,
                engine="com",
            )
        except runtime.PatchOpError as exc:
            if _should_fallback_on_com_patch_error(
                exc,
                request=effective_request,
                input_path=resolved_input,
            ):
                warnings.append(
                    f"COM patch failed; falling back to openpyxl. ({exc!r})"
                )
                return _apply_with_openpyxl(
                    effective_request,
                    resolved_input,
                    output_path,
                    warnings,
                )
            error = _coerce_patch_error_detail(exc.detail)
            return PatchResult(
                out_path=str(output_path),
                patch_diff=[],
                inverse_ops=[],
                formula_issues=[],
                warnings=warnings,
                error=error,
                engine="com",
            )
        except Exception as exc:
            if runtime.allow_auto_openpyxl_fallback(effective_request, resolved_input):
                warnings.append(
                    f"COM patch failed; falling back to openpyxl. ({exc!r})"
                )
                return _apply_with_openpyxl(
                    effective_request,
                    resolved_input,
                    output_path,
                    warnings,
                )
            raise RuntimeError(f"COM patch failed: {exc}") from exc

    return _apply_with_openpyxl(
        effective_request,
        resolved_input,
        output_path,
        warnings,
    )


def _resolve_effective_request(
    request: PatchRequest,
) -> PatchRequest:
    """Resolve request-level backend adjustments and mixed-op guards."""
    if runtime.contains_create_chart_op(request.ops):
        if not runtime.contains_apply_table_style_op(request.ops):
            return request
        raise ValueError(
            "create_chart and apply_table_style cannot be combined in one patch request. "
            "create_chart is COM-only, and a patch request runs on a single backend engine. "
            "Run them in separate requests to isolate chart and table operations."
        )
    return request


def _should_fallback_on_com_patch_error(
    exc: runtime.PatchOpError, *, request: PatchRequest, input_path: Path
) -> bool:
    """Return whether PatchOpError from COM path should trigger openpyxl fallback."""
    if not runtime.allow_auto_openpyxl_fallback(request, input_path):
        return False
    detail = exc.detail
    if detail.error_code == "com_runtime_error":
        return True
    return detail.raw_com_message is not None


def _apply_with_openpyxl(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
    warnings: list[str],
) -> PatchResult:
    """Apply patch operations using openpyxl."""
    try:
        engine_result = apply_openpyxl_engine(
            request,
            input_path,
            output_path,
        )
    except runtime.PatchOpError as exc:
        error = _coerce_patch_error_detail(exc.detail)
        return PatchResult(
            out_path=str(output_path),
            patch_diff=[],
            inverse_ops=[],
            formula_issues=[],
            warnings=warnings,
            error=error,
            engine="openpyxl",
        )
    except ValueError:
        raise
    except FileNotFoundError:
        raise
    except OSError:
        raise
    except Exception as exc:
        raise RuntimeError(f"openpyxl patch failed: {exc}") from exc

    patch_diff = _coerce_patch_diff_items(engine_result.patch_diff)
    typed_inverse_ops = _coerce_inverse_ops(engine_result.inverse_ops)
    typed_formula_issues = _coerce_formula_issues(engine_result.formula_issues)
    warnings.extend(engine_result.op_warnings)
    if not request.dry_run:
        warnings.append(
            "openpyxl editing may drop shapes/charts or unsupported elements."
        )
    _append_skip_warnings(warnings, patch_diff)
    if (
        not request.dry_run
        and request.preflight_formula_check
        and any(issue.level == "error" for issue in typed_formula_issues)
    ):
        issue = typed_formula_issues[0]
        op_index, op_name = _find_preflight_issue_origin(issue, request.ops)
        error = PatchErrorDetail(
            op_index=op_index,
            op=op_name,
            sheet=issue.sheet,
            cell=issue.cell,
            message=f"Formula health check failed: {issue.message}",
            hint=None,
            expected_fields=[],
            example_op=None,
        )
        return PatchResult(
            out_path=str(output_path),
            patch_diff=[],
            inverse_ops=[],
            formula_issues=typed_formula_issues,
            warnings=warnings,
            error=error,
            engine="openpyxl",
        )
    return PatchResult(
        out_path=str(output_path),
        patch_diff=patch_diff,
        inverse_ops=typed_inverse_ops,
        formula_issues=typed_formula_issues,
        warnings=warnings,
        engine="openpyxl",
    )


def _append_skip_warnings(warnings: list[str], diff: list[PatchDiffItem]) -> None:
    """Append warning messages for skipped conditional operations."""
    for item in diff:
        if item.status != "skipped":
            continue
        warnings.append(
            f"Skipped op[{item.op_index}] {item.op} at {item.sheet}!{item.cell} due to condition mismatch."
        )


def _find_preflight_issue_origin(
    issue: FormulaIssue, ops: list[PatchOp]
) -> tuple[int, PatchOpType]:
    """Find the most likely op index/op name for a preflight formula issue."""
    for index, op in enumerate(ops):
        if _op_targets_issue_cell(op, issue.sheet, issue.cell):
            return index, op.op
    return -1, "set_value"


def _op_targets_issue_cell(op: PatchOp, sheet: str, cell: str) -> bool:
    """Return True when an op can affect the specified sheet/cell."""
    if op.sheet != sheet:
        return False
    if op.cell is not None:
        return op.cell == cell
    if op.range is None:
        return False
    for row in runtime.expand_range_coordinates(op.range):
        if cell in row:
            return True
    return False


def _coerce_patch_diff_items(items: Sequence[object]) -> list[PatchDiffItem]:
    """Coerce backend diff items into canonical PatchDiffItem models."""
    return _coerce_model_list(items, PatchDiffItem)


def _coerce_inverse_ops(items: Sequence[object]) -> list[PatchOp]:
    """Coerce backend inverse ops into canonical PatchOp models."""
    return _coerce_model_list(items, PatchOp)


def _coerce_formula_issues(items: Sequence[object]) -> list[FormulaIssue]:
    """Coerce backend formula findings into canonical FormulaIssue models."""
    return _coerce_model_list(items, FormulaIssue)


def _coerce_patch_error_detail(detail: object) -> PatchErrorDetail | None:
    """Coerce backend error detail into canonical PatchErrorDetail model."""
    coerced = _coerce_model_list([detail], PatchErrorDetail)
    if not coerced:
        return None
    return coerced[0]


def _coerce_model_list(
    items: Sequence[object], model_cls: type[TModel]
) -> list[TModel]:
    """Convert model-like items to target Pydantic models and skip invalid entries."""
    coerced: list[TModel] = []
    for item in items:
        try:
            if isinstance(item, model_cls):
                coerced.append(item)
                continue
            source: object
            if isinstance(item, BaseModel):
                source = item.model_dump(mode="python")
            else:
                source = item
            coerced.append(model_cls.model_validate(source))
        except ValidationError:
            continue
    return coerced


__all__ = ["run_make", "run_patch"]
