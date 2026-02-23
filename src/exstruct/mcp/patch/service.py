from __future__ import annotations

from pathlib import Path

from exstruct.mcp.io import PathPolicy
import exstruct.mcp.patch.legacy_runner as runner
from exstruct.mcp.patch.legacy_runner import (
    FormulaIssue,
    MakeRequest,
    PatchDiffItem,
    PatchErrorDetail,
    PatchOp,
    PatchRequest,
    PatchResult,
)

from .engine.openpyxl_engine import apply_openpyxl_engine
from .engine.xlwings_engine import apply_xlwings_engine
from .types import PatchOpType


def run_make(request: MakeRequest, *, policy: PathPolicy | None = None) -> PatchResult:
    """Create a new workbook and apply patch operations in one call."""
    resolved_output = runner._resolve_make_output_path(request.out_path, policy=policy)
    runner._ensure_supported_extension(resolved_output)
    runner._validate_make_request_constraints(request, resolved_output)
    seed_path = runner._build_make_seed_path(resolved_output)
    initial_sheet_name = runner._resolve_make_initial_sheet_name(request)
    try:
        runner._create_seed_workbook(
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
    resolved_input = runner._resolve_input_path(request.xlsx_path, policy=policy)
    runner._ensure_supported_extension(resolved_input)
    output_path = runner._resolve_output_path(
        resolved_input,
        out_dir=request.out_dir,
        out_name=request.out_name,
        policy=policy,
    )
    warnings: list[str] = []
    runner._append_large_ops_warning(warnings, request.ops)
    effective_request = request
    if request.backend == "com" and runner._contains_apply_table_style_op(request.ops):
        warnings.append(
            "backend='com' does not support apply_table_style; falling back to openpyxl."
        )
        effective_request = request.model_copy(update={"backend": "openpyxl"})
    if resolved_input.suffix.lower() == ".xls" and runner._contains_design_ops(
        effective_request.ops
    ):
        raise ValueError(
            "Design operations are not supported for .xls files. Convert to .xlsx/.xlsm first."
        )
    com = runner.get_com_availability()
    selected_engine = runner._select_patch_engine(
        request=effective_request,
        input_path=resolved_input,
        com_available=com.available,
    )
    output_path, warning, skipped = runner._apply_conflict_policy(
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
    if selected_engine == "openpyxl" and runner._requires_openpyxl_backend(
        effective_request
    ):
        warnings.append("Using openpyxl backend due to patch request constraints.")

    runner._ensure_output_dir(output_path)
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
                patch_diff=[item for item in diff if isinstance(item, PatchDiffItem)],
                inverse_ops=[],
                formula_issues=[],
                warnings=warnings,
                engine="com",
            )
        except runner.PatchOpError as exc:
            return PatchResult(
                out_path=str(output_path),
                patch_diff=[],
                inverse_ops=[],
                formula_issues=[],
                warnings=warnings,
                error=exc.detail,
                engine="com",
            )
        except Exception as exc:
            if runner._allow_auto_openpyxl_fallback(effective_request, resolved_input):
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


def _apply_with_openpyxl(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
    warnings: list[str],
) -> PatchResult:
    """Apply patch operations using openpyxl."""
    try:
        diff, inverse_ops, formula_issues, op_warnings = apply_openpyxl_engine(
            request,
            input_path,
            output_path,
        )
    except runner.PatchOpError as exc:
        return PatchResult(
            out_path=str(output_path),
            patch_diff=[],
            inverse_ops=[],
            formula_issues=[],
            warnings=warnings,
            error=exc.detail,
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

    patch_diff = [item for item in diff if isinstance(item, PatchDiffItem)]
    typed_inverse_ops = [item for item in inverse_ops if isinstance(item, PatchOp)]
    typed_formula_issues = [
        item for item in formula_issues if isinstance(item, FormulaIssue)
    ]
    warnings.extend(op_warnings)
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
    for row in runner._expand_range_coordinates(op.range):
        if cell in row:
            return True
    return False


__all__ = ["run_make", "run_patch"]
