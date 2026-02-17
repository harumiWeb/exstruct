from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
import re
from typing import Literal, Protocol, runtime_checkable

from pydantic import BaseModel, Field, field_validator, model_validator
import xlwings as xw

from exstruct.cli.availability import get_com_availability

from .extract_runner import OnConflictPolicy
from .io import PathPolicy

PatchOpType = Literal[
    "set_value",
    "set_formula",
    "add_sheet",
    "set_range_values",
    "fill_formula",
    "set_value_if",
    "set_formula_if",
]
PatchStatus = Literal["applied", "skipped"]
PatchValueKind = Literal["value", "formula", "sheet"]
FormulaIssueLevel = Literal["warning", "error"]
FormulaIssueCode = Literal[
    "invalid_token",
    "ref_error",
    "name_error",
    "div0_error",
    "value_error",
    "na_error",
    "circular_ref_suspected",
]

_ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
_A1_PATTERN = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*$")
_A1_RANGE_PATTERN = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*:[A-Za-z]{1,3}[1-9][0-9]*$")


@runtime_checkable
class OpenpyxlCellProtocol(Protocol):
    """Protocol for openpyxl cell access used by patch runner."""

    value: str | int | float | None
    data_type: str | None


@runtime_checkable
class OpenpyxlWorksheetProtocol(Protocol):
    """Protocol for openpyxl worksheet access used by patch runner."""

    def __getitem__(self, key: str) -> OpenpyxlCellProtocol: ...


@runtime_checkable
class OpenpyxlWorkbookProtocol(Protocol):
    """Protocol for openpyxl workbook access used by patch runner."""

    sheetnames: list[str]

    def __getitem__(self, key: str) -> OpenpyxlWorksheetProtocol: ...

    def create_sheet(self, title: str) -> OpenpyxlWorksheetProtocol: ...

    def save(self, filename: str | Path) -> None: ...

    def close(self) -> None: ...


@runtime_checkable
class XlwingsRangeProtocol(Protocol):
    """Protocol for xlwings range access used by patch runner."""

    value: str | int | float | None
    formula: str | None


@runtime_checkable
class XlwingsSheetProtocol(Protocol):
    """Protocol for xlwings sheet access used by patch runner."""

    name: str

    def range(self, cell: str) -> XlwingsRangeProtocol: ...


@runtime_checkable
class XlwingsSheetsProtocol(Protocol):
    """Protocol for xlwings sheets collection."""

    def __iter__(self) -> Iterator[XlwingsSheetProtocol]: ...

    def __len__(self) -> int: ...

    def __getitem__(self, index: int) -> XlwingsSheetProtocol: ...

    def add(
        self, name: str, after: XlwingsSheetProtocol | None = None
    ) -> XlwingsSheetProtocol: ...


@runtime_checkable
class XlwingsWorkbookProtocol(Protocol):
    """Protocol for xlwings workbook access used by patch runner."""

    sheets: XlwingsSheetsProtocol

    def save(self, filename: str) -> None: ...

    def close(self) -> None: ...


class PatchOp(BaseModel):
    """Single patch operation for an Excel workbook.

    Operation types and their required fields:

    - ``set_value``: Set a cell value. Requires ``sheet``, ``cell``, ``value``.
    - ``set_formula``: Set a cell formula. Requires ``sheet``, ``cell``, ``formula`` (must start with ``=``).
    - ``add_sheet``: Add a new worksheet. Requires ``sheet`` (new sheet name). No ``cell``/``value``/``formula``.
    - ``set_range_values``: Set values for a rectangular range. Requires ``sheet``, ``range`` (e.g. ``A1:C3``), ``values`` (2D list matching range shape).
    - ``fill_formula``: Fill a formula across a single row or column. Requires ``sheet``, ``range``, ``base_cell``, ``formula``.
    - ``set_value_if``: Conditionally set value. Requires ``sheet``, ``cell``, ``value``. ``expected`` is optional; ``null`` matches an empty cell. Skips if current value != expected.
    - ``set_formula_if``: Conditionally set formula. Requires ``sheet``, ``cell``, ``formula``. ``expected`` is optional; ``null`` matches an empty cell. Skips if current value != expected.
    """

    op: PatchOpType = Field(
        description=(
            "Operation type: 'set_value', 'set_formula', 'add_sheet', "
            "'set_range_values', 'fill_formula', 'set_value_if', or 'set_formula_if'."
        )
    )
    sheet: str = Field(
        description="Target sheet name. For add_sheet, this is the new sheet name."
    )
    cell: str | None = Field(
        default=None,
        description="Cell reference in A1 notation (e.g. 'B2'). Required for set_value, set_formula, set_value_if, set_formula_if.",
    )
    range: str | None = Field(
        default=None,
        description="Range reference in A1 notation (e.g. 'A1:C3'). Required for set_range_values and fill_formula.",
    )
    base_cell: str | None = Field(
        default=None,
        description="Base cell for formula translation in fill_formula (e.g. 'C2').",
    )
    expected: str | int | float | None = Field(
        default=None,
        description="Expected current value for conditional ops (set_value_if, set_formula_if). Operation is skipped if mismatch.",
    )
    value: str | int | float | None = Field(
        default=None,
        description="Value to set. Use null to clear a cell. For set_value and set_value_if.",
    )
    values: list[list[str | int | float | None]] | None = Field(
        default=None,
        description="2D list of values for set_range_values. Shape must match the range dimensions.",
    )
    formula: str | None = Field(
        default=None,
        description="Formula string starting with '=' (e.g. '=SUM(A1:A10)'). For set_formula, set_formula_if, fill_formula.",
    )

    @field_validator("sheet")
    @classmethod
    def _validate_sheet(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("sheet must not be empty.")
        return value

    @field_validator("cell")
    @classmethod
    def _validate_cell(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not _A1_PATTERN.match(candidate):
            raise ValueError(f"Invalid cell reference: {value}")
        return candidate.upper()

    @field_validator("base_cell")
    @classmethod
    def _validate_base_cell(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not _A1_PATTERN.match(candidate):
            raise ValueError(f"Invalid base_cell reference: {value}")
        return candidate.upper()

    @field_validator("range")
    @classmethod
    def _validate_range(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not _A1_RANGE_PATTERN.match(candidate):
            raise ValueError(f"Invalid range reference: {value}")
        start, end = candidate.split(":", maxsplit=1)
        return f"{start.upper()}:{end.upper()}"

    @model_validator(mode="after")
    def _validate_op(self) -> PatchOp:
        if self.op == "add_sheet":
            _validate_add_sheet(self)
            return self
        if self.op == "set_value":
            _validate_cell_required(self)
            _validate_set_value(self)
            return self
        if self.op == "set_formula":
            _validate_cell_required(self)
            _validate_set_formula(self)
            return self
        if self.op == "set_range_values":
            _validate_set_range_values(self)
            return self
        if self.op == "fill_formula":
            _validate_fill_formula(self)
            return self
        if self.op == "set_value_if":
            _validate_cell_required(self)
            _validate_set_value_if(self)
            return self
        if self.op == "set_formula_if":
            _validate_cell_required(self)
            _validate_set_formula_if(self)
            return self
        return self


def _validate_add_sheet(op: PatchOp) -> None:
    """Validate add_sheet operation."""
    if op.cell is not None:
        raise ValueError("add_sheet does not accept cell.")
    if op.range is not None:
        raise ValueError("add_sheet does not accept range.")
    if op.base_cell is not None:
        raise ValueError("add_sheet does not accept base_cell.")
    if op.expected is not None:
        raise ValueError("add_sheet does not accept expected.")
    if op.value is not None:
        raise ValueError("add_sheet does not accept value.")
    if op.values is not None:
        raise ValueError("add_sheet does not accept values.")
    if op.formula is not None:
        raise ValueError("add_sheet does not accept formula.")


def _validate_cell_required(op: PatchOp) -> None:
    """Validate that the operation has a cell value."""
    if op.cell is None:
        raise ValueError(f"{op.op} requires cell.")


def _validate_set_value(op: PatchOp) -> None:
    """Validate set_value operation."""
    if op.range is not None:
        raise ValueError("set_value does not accept range.")
    if op.base_cell is not None:
        raise ValueError("set_value does not accept base_cell.")
    if op.expected is not None:
        raise ValueError("set_value does not accept expected.")
    if op.values is not None:
        raise ValueError("set_value does not accept values.")
    if op.formula is not None:
        raise ValueError("set_value does not accept formula.")


def _validate_set_formula(op: PatchOp) -> None:
    """Validate set_formula operation."""
    if op.range is not None:
        raise ValueError("set_formula does not accept range.")
    if op.base_cell is not None:
        raise ValueError("set_formula does not accept base_cell.")
    if op.expected is not None:
        raise ValueError("set_formula does not accept expected.")
    if op.values is not None:
        raise ValueError("set_formula does not accept values.")
    if op.value is not None:
        raise ValueError("set_formula does not accept value.")
    if op.formula is None:
        raise ValueError("set_formula requires formula.")
    if not op.formula.startswith("="):
        raise ValueError("set_formula requires formula starting with '='.")


def _validate_set_range_values(op: PatchOp) -> None:
    """Validate set_range_values operation."""
    if op.cell is not None:
        raise ValueError("set_range_values does not accept cell.")
    if op.base_cell is not None:
        raise ValueError("set_range_values does not accept base_cell.")
    if op.expected is not None:
        raise ValueError("set_range_values does not accept expected.")
    if op.formula is not None:
        raise ValueError("set_range_values does not accept formula.")
    if op.range is None:
        raise ValueError("set_range_values requires range.")
    if op.values is None:
        raise ValueError("set_range_values requires values.")
    if not op.values:
        raise ValueError("set_range_values requires non-empty values.")
    if not all(op.values):
        raise ValueError("set_range_values values rows must not be empty.")
    expected_width = len(op.values[0])
    if any(len(row) != expected_width for row in op.values):
        raise ValueError("set_range_values requires rectangular values.")


def _validate_fill_formula(op: PatchOp) -> None:
    """Validate fill_formula operation."""
    if op.cell is not None:
        raise ValueError("fill_formula does not accept cell.")
    if op.expected is not None:
        raise ValueError("fill_formula does not accept expected.")
    if op.value is not None:
        raise ValueError("fill_formula does not accept value.")
    if op.values is not None:
        raise ValueError("fill_formula does not accept values.")
    if op.range is None:
        raise ValueError("fill_formula requires range.")
    if op.base_cell is None:
        raise ValueError("fill_formula requires base_cell.")
    if op.formula is None:
        raise ValueError("fill_formula requires formula.")
    if not op.formula.startswith("="):
        raise ValueError("fill_formula requires formula starting with '='.")


def _validate_set_value_if(op: PatchOp) -> None:
    """Validate set_value_if operation."""
    if op.formula is not None:
        raise ValueError("set_value_if does not accept formula.")
    if op.range is not None:
        raise ValueError("set_value_if does not accept range.")
    if op.values is not None:
        raise ValueError("set_value_if does not accept values.")
    if op.base_cell is not None:
        raise ValueError("set_value_if does not accept base_cell.")


def _validate_set_formula_if(op: PatchOp) -> None:
    """Validate set_formula_if operation."""
    if op.value is not None:
        raise ValueError("set_formula_if does not accept value.")
    if op.range is not None:
        raise ValueError("set_formula_if does not accept range.")
    if op.values is not None:
        raise ValueError("set_formula_if does not accept values.")
    if op.base_cell is not None:
        raise ValueError("set_formula_if does not accept base_cell.")
    if op.formula is None:
        raise ValueError("set_formula_if requires formula.")
    if not op.formula.startswith("="):
        raise ValueError("set_formula_if requires formula starting with '='.")


class PatchValue(BaseModel):
    """Normalized before/after value in patch diff."""

    kind: PatchValueKind
    value: str | int | float | None


class PatchDiffItem(BaseModel):
    """Applied change record for patch operations."""

    op_index: int
    op: PatchOpType
    sheet: str
    cell: str | None = None
    before: PatchValue | None = None
    after: PatchValue | None = None
    status: PatchStatus = "applied"


class PatchErrorDetail(BaseModel):
    """Structured error details for patch failures."""

    op_index: int
    op: PatchOpType
    sheet: str
    cell: str | None
    message: str


class FormulaIssue(BaseModel):
    """Formula health-check finding."""

    sheet: str
    cell: str
    level: FormulaIssueLevel
    code: FormulaIssueCode
    message: str


class PatchRequest(BaseModel):
    """Input model for ExStruct MCP patch."""

    xlsx_path: Path
    ops: list[PatchOp]
    out_dir: Path | None = None
    out_name: str | None = None
    on_conflict: OnConflictPolicy = "overwrite"
    auto_formula: bool = False
    dry_run: bool = False
    return_inverse_ops: bool = False
    preflight_formula_check: bool = False


class PatchResult(BaseModel):
    """Output model for ExStruct MCP patch."""

    out_path: str
    patch_diff: list[PatchDiffItem] = Field(default_factory=list)
    inverse_ops: list[PatchOp] = Field(default_factory=list)
    formula_issues: list[FormulaIssue] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    error: PatchErrorDetail | None = None


def run_patch(
    request: PatchRequest, *, policy: PathPolicy | None = None
) -> PatchResult:
    """Run a patch operation and write the updated workbook.

    Args:
        request: Patch request payload.
        policy: Optional path policy for access control.

    Returns:
        Patch result with output path and diff.

    Raises:
        FileNotFoundError: If the input file does not exist.
        ValueError: If validation fails or the path violates policy.
        RuntimeError: If a backend operation fails.
    """
    resolved_input = _resolve_input_path(request.xlsx_path, policy=policy)
    _ensure_supported_extension(resolved_input)
    output_path = _resolve_output_path(
        resolved_input,
        out_dir=request.out_dir,
        out_name=request.out_name,
        policy=policy,
    )
    output_path, warning, skipped = _apply_conflict_policy(
        output_path, request.on_conflict
    )
    warnings: list[str] = []
    if warning:
        warnings.append(warning)
    if skipped and not request.dry_run:
        return PatchResult(
            out_path=str(output_path),
            patch_diff=[],
            inverse_ops=[],
            formula_issues=[],
            warnings=warnings,
        )
    if skipped and request.dry_run:
        warnings.append(
            "Dry-run mode ignores on_conflict=skip and simulates patch without writing."
        )

    com = get_com_availability()
    if resolved_input.suffix.lower() == ".xls" and not com.available:
        raise ValueError(
            ".xls editing requires Windows Excel COM (xlwings) in this environment."
        )

    use_openpyxl = _requires_openpyxl_backend(request)
    if use_openpyxl and com.available:
        warnings.append("Using openpyxl backend for extended patch features.")

    _ensure_output_dir(output_path)
    if com.available and not use_openpyxl:
        try:
            diff = _apply_ops_xlwings(
                resolved_input,
                output_path,
                request.ops,
                request.auto_formula,
            )
            return PatchResult(
                out_path=str(output_path),
                patch_diff=diff,
                inverse_ops=[],
                formula_issues=[],
                warnings=warnings,
            )
        except PatchOpError as exc:
            return PatchResult(
                out_path=str(output_path),
                patch_diff=[],
                inverse_ops=[],
                formula_issues=[],
                warnings=warnings,
                error=exc.detail,
            )
        except Exception as exc:
            fallback = _maybe_fallback_openpyxl(
                request,
                resolved_input,
                output_path,
                warnings,
                reason=f"COM patch failed; falling back to openpyxl. ({exc!r})",
            )
            if fallback is not None:
                return fallback
            raise RuntimeError(f"COM patch failed: {exc}") from exc

    if com.reason:
        warnings.append(f"COM unavailable: {com.reason}")
    return _apply_with_openpyxl(
        request,
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
        diff, inverse_ops, formula_issues = _apply_ops_openpyxl(
            request,
            input_path,
            output_path,
        )
    except PatchOpError as exc:
        return PatchResult(
            out_path=str(output_path),
            patch_diff=[],
            inverse_ops=[],
            formula_issues=[],
            warnings=warnings,
            error=exc.detail,
        )
    except ValueError:
        raise
    except FileNotFoundError:
        raise
    except OSError:
        raise
    except Exception as exc:
        raise RuntimeError(f"openpyxl patch failed: {exc}") from exc

    if not request.dry_run:
        warnings.append(
            "openpyxl editing may drop shapes/charts or unsupported elements."
        )
    _append_skip_warnings(warnings, diff)
    if (
        not request.dry_run
        and request.preflight_formula_check
        and any(issue.level == "error" for issue in formula_issues)
    ):
        issue = formula_issues[0]
        op_index, op_name = _find_preflight_issue_origin(issue, request.ops)
        error = PatchErrorDetail(
            op_index=op_index,
            op=op_name,
            sheet=issue.sheet,
            cell=issue.cell,
            message=f"Formula health check failed: {issue.message}",
        )
        return PatchResult(
            out_path=str(output_path),
            patch_diff=[],
            inverse_ops=[],
            formula_issues=formula_issues,
            warnings=warnings,
            error=error,
        )
    return PatchResult(
        out_path=str(output_path),
        patch_diff=diff,
        inverse_ops=inverse_ops,
        formula_issues=formula_issues,
        warnings=warnings,
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
    for row in _expand_range_coordinates(op.range):
        if cell in row:
            return True
    return False


def _maybe_fallback_openpyxl(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
    warnings: list[str],
    *,
    reason: str,
) -> PatchResult | None:
    """Attempt openpyxl fallback after COM failure."""
    if input_path.suffix.lower() == ".xls":
        warnings.append(reason)
        return None
    warnings.append(reason)
    return _apply_with_openpyxl(
        request,
        input_path,
        output_path,
        warnings,
    )


def _requires_openpyxl_backend(request: PatchRequest) -> bool:
    """Return True if request requires openpyxl backend for extended features."""
    if request.dry_run or request.return_inverse_ops or request.preflight_formula_check:
        return True
    return any(
        op.op in {"set_range_values", "fill_formula", "set_value_if", "set_formula_if"}
        for op in request.ops
    )


def _resolve_input_path(path: Path, *, policy: PathPolicy | None) -> Path:
    """Resolve and validate the input path."""
    resolved = policy.ensure_allowed(path) if policy else path.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"Input file not found: {resolved}")
    if not resolved.is_file():
        raise ValueError(f"Input path is not a file: {resolved}")
    return resolved


def _ensure_supported_extension(path: Path) -> None:
    """Validate that the input file extension is supported."""
    if path.suffix.lower() not in _ALLOWED_EXTENSIONS:
        raise ValueError(f"Unsupported file extension: {path.suffix}")


def _resolve_output_path(
    input_path: Path,
    *,
    out_dir: Path | None,
    out_name: str | None,
    policy: PathPolicy | None,
) -> Path:
    """Build and validate the output path."""
    target_dir = out_dir or input_path.parent
    target_dir = policy.ensure_allowed(target_dir) if policy else target_dir.resolve()
    name = _normalize_output_name(input_path, out_name)
    output_path = (target_dir / name).resolve()
    if policy is not None:
        output_path = policy.ensure_allowed(output_path)
    return output_path


def _normalize_output_name(input_path: Path, out_name: str | None) -> str:
    """Normalize output filename with a safe suffix."""
    if out_name:
        candidate = Path(out_name)
        return (
            candidate.name
            if candidate.suffix
            else f"{candidate.name}{input_path.suffix}"
        )
    return f"{input_path.stem}_patched{input_path.suffix}"


def _ensure_output_dir(path: Path) -> None:
    """Ensure the output directory exists before writing."""
    path.parent.mkdir(parents=True, exist_ok=True)


def _apply_conflict_policy(
    output_path: Path, on_conflict: OnConflictPolicy
) -> tuple[Path, str | None, bool]:
    """Apply output conflict policy to a resolved output path."""
    if not output_path.exists():
        return output_path, None, False
    if on_conflict == "skip":
        return (
            output_path,
            f"Output exists; skipping write: {output_path.name}",
            True,
        )
    if on_conflict == "rename":
        renamed = _next_available_path(output_path)
        return (
            renamed,
            f"Output exists; renamed to: {renamed.name}",
            False,
        )
    return output_path, None, False


def _next_available_path(path: Path) -> Path:
    """Return the next available path by appending a numeric suffix."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for idx in range(1, 10_000):
        candidate = path.with_name(f"{stem}_{idx}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Failed to resolve unique path for {path}")


def _apply_ops_openpyxl(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
) -> tuple[list[PatchDiffItem], list[PatchOp], list[FormulaIssue]]:
    """Apply operations using openpyxl."""
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError(f"openpyxl is not available: {exc}") from exc

    if input_path.suffix.lower() == ".xls":
        raise ValueError("openpyxl cannot edit .xls files.")

    if input_path.suffix.lower() == ".xlsm":
        workbook = load_workbook(input_path, keep_vba=True)
    else:
        workbook = load_workbook(input_path)
    try:
        diff, inverse_ops = _apply_ops_to_openpyxl_workbook(
            workbook,
            request.ops,
            request.auto_formula,
            return_inverse_ops=request.return_inverse_ops,
        )
        formula_issues = (
            _collect_formula_issues_openpyxl(workbook)
            if request.preflight_formula_check
            else []
        )
        if not request.dry_run and not (
            request.preflight_formula_check
            and any(issue.level == "error" for issue in formula_issues)
        ):
            workbook.save(output_path)
    finally:
        workbook.close()
    return diff, inverse_ops, formula_issues


def _apply_ops_to_openpyxl_workbook(
    workbook: OpenpyxlWorkbookProtocol,
    ops: list[PatchOp],
    auto_formula: bool,
    *,
    return_inverse_ops: bool,
) -> tuple[list[PatchDiffItem], list[PatchOp]]:
    """Apply ops to an openpyxl workbook instance."""
    sheets = _openpyxl_sheet_map(workbook)
    diff: list[PatchDiffItem] = []
    inverse_ops: list[PatchOp] = []
    for index, op in enumerate(ops):
        try:
            item, inverse = _apply_openpyxl_op(
                workbook, sheets, op, index, auto_formula
            )
            diff.append(item)
            if return_inverse_ops and item.status == "applied" and inverse is not None:
                inverse_ops.append(inverse)
        except ValueError as exc:
            raise PatchOpError.from_op(index, op, exc) from exc
    if return_inverse_ops:
        inverse_ops.reverse()
    return diff, inverse_ops


def _openpyxl_sheet_map(
    workbook: OpenpyxlWorkbookProtocol,
) -> dict[str, OpenpyxlWorksheetProtocol]:
    """Build a sheet map for openpyxl workbooks."""
    sheet_names = getattr(workbook, "sheetnames", None)
    if not isinstance(sheet_names, list):
        raise ValueError("Invalid workbook: sheetnames missing.")
    return {name: workbook[name] for name in sheet_names}


def _apply_openpyxl_op(
    workbook: OpenpyxlWorkbookProtocol,
    sheets: dict[str, OpenpyxlWorksheetProtocol],
    op: PatchOp,
    index: int,
    auto_formula: bool,
) -> tuple[PatchDiffItem, PatchOp | None]:
    """Apply a single op to openpyxl workbook."""
    if op.op == "add_sheet":
        return _apply_openpyxl_add_sheet(workbook, sheets, op, index)

    existing_sheet = sheets.get(op.sheet)
    if existing_sheet is None:
        raise ValueError(f"Sheet not found: {op.sheet}")

    if op.op == "set_range_values":
        return _apply_openpyxl_set_range_values(existing_sheet, op, index)

    if op.op == "fill_formula":
        return _apply_openpyxl_fill_formula(existing_sheet, op, index)

    if op.op in {"set_value", "set_formula", "set_value_if", "set_formula_if"}:
        return _apply_openpyxl_cell_op(existing_sheet, op, index, auto_formula)
    raise ValueError(f"Unsupported op: {op.op}")


def _apply_openpyxl_add_sheet(
    workbook: OpenpyxlWorkbookProtocol,
    sheets: dict[str, OpenpyxlWorksheetProtocol],
    op: PatchOp,
    index: int,
) -> tuple[PatchDiffItem, PatchOp | None]:
    """Apply add_sheet op."""
    if op.sheet in sheets:
        raise ValueError(f"Sheet already exists: {op.sheet}")
    sheet = workbook.create_sheet(title=op.sheet)
    sheets[op.sheet] = sheet
    return (
        PatchDiffItem(
            op_index=index,
            op=op.op,
            sheet=op.sheet,
            cell=None,
            before=None,
            after=PatchValue(kind="sheet", value=op.sheet),
        ),
        None,
    )


def _apply_openpyxl_set_range_values(
    sheet: OpenpyxlWorksheetProtocol,
    op: PatchOp,
    index: int,
) -> tuple[PatchDiffItem, PatchOp | None]:
    """Apply set_range_values op."""
    if op.range is None or op.values is None:
        raise ValueError("set_range_values requires range and values.")
    coordinates = _expand_range_coordinates(op.range)
    rows, cols = _shape_of_coordinates(coordinates)
    if len(op.values) != rows:
        raise ValueError("set_range_values values height does not match range.")
    if any(len(row) != cols for row in op.values):
        raise ValueError("set_range_values values width does not match range.")
    for r_idx, row in enumerate(coordinates):
        for c_idx, coord in enumerate(row):
            sheet[coord].value = op.values[r_idx][c_idx]
    return (
        PatchDiffItem(
            op_index=index,
            op=op.op,
            sheet=op.sheet,
            cell=op.range,
            before=None,
            after=PatchValue(kind="value", value=f"{rows}x{cols}"),
        ),
        None,
    )


def _apply_openpyxl_fill_formula(
    sheet: OpenpyxlWorksheetProtocol,
    op: PatchOp,
    index: int,
) -> tuple[PatchDiffItem, PatchOp | None]:
    """Apply fill_formula op."""
    if op.range is None or op.formula is None or op.base_cell is None:
        raise ValueError("fill_formula requires range, base_cell and formula.")
    coordinates = _expand_range_coordinates(op.range)
    rows, cols = _shape_of_coordinates(coordinates)
    if rows != 1 and cols != 1:
        raise ValueError("fill_formula range must be a single row or a single column.")
    for row in coordinates:
        for coord in row:
            sheet[coord].value = _translate_formula(op.formula, op.base_cell, coord)
    return (
        PatchDiffItem(
            op_index=index,
            op=op.op,
            sheet=op.sheet,
            cell=op.range,
            before=None,
            after=PatchValue(kind="formula", value=op.formula),
        ),
        None,
    )


def _apply_openpyxl_cell_op(
    sheet: OpenpyxlWorksheetProtocol,
    op: PatchOp,
    index: int,
    auto_formula: bool,
) -> tuple[PatchDiffItem, PatchOp | None]:
    """Apply single-cell operations."""
    cell_ref = op.cell
    if cell_ref is None:
        raise ValueError(f"{op.op} requires cell.")
    cell = sheet[cell_ref]
    before = _openpyxl_cell_value(cell)

    if op.op == "set_value":
        after = _set_cell_value(cell, op.value, auto_formula, op_name="set_value")
        return _build_cell_result(
            op, index, cell_ref, before, after
        ), _build_inverse_cell_op(op, cell_ref, before)
    if op.op == "set_formula":
        formula = _require_formula(op.formula, "set_formula")
        cell.value = formula
        after = PatchValue(kind="formula", value=formula)
        return _build_cell_result(
            op, index, cell_ref, before, after
        ), _build_inverse_cell_op(op, cell_ref, before)
    if op.op == "set_value_if":
        if not _values_equal_for_condition(
            _patch_value_to_primitive(before), op.expected
        ):
            return _build_skipped_result(op, index, cell_ref, before), None
        after = _set_cell_value(cell, op.value, auto_formula, op_name="set_value_if")
        return _build_cell_result(
            op, index, cell_ref, before, after
        ), _build_inverse_cell_op(op, cell_ref, before)
    formula_if = _require_formula(op.formula, "set_formula_if")
    if not _values_equal_for_condition(_patch_value_to_primitive(before), op.expected):
        return _build_skipped_result(op, index, cell_ref, before), None
    cell.value = formula_if
    after = PatchValue(kind="formula", value=formula_if)
    return _build_cell_result(
        op, index, cell_ref, before, after
    ), _build_inverse_cell_op(op, cell_ref, before)


def _set_cell_value(
    cell: OpenpyxlCellProtocol,
    value: str | int | float | None,
    auto_formula: bool,
    *,
    op_name: str,
) -> PatchValue:
    """Set cell value with auto_formula handling."""
    if isinstance(value, str) and value.startswith("="):
        if not auto_formula:
            raise ValueError(f"{op_name} rejects values starting with '='.")
        cell.value = value
        return PatchValue(kind="formula", value=value)
    cell.value = value
    return PatchValue(kind="value", value=value)


def _build_cell_result(
    op: PatchOp,
    index: int,
    cell_ref: str,
    before: PatchValue | None,
    after: PatchValue | None,
) -> PatchDiffItem:
    """Build applied diff item for single-cell op."""
    return PatchDiffItem(
        op_index=index,
        op=op.op,
        sheet=op.sheet,
        cell=cell_ref,
        before=before,
        after=after,
    )


def _build_skipped_result(
    op: PatchOp,
    index: int,
    cell_ref: str,
    before: PatchValue | None,
) -> PatchDiffItem:
    """Build skipped diff item."""
    return PatchDiffItem(
        op_index=index,
        op=op.op,
        sheet=op.sheet,
        cell=cell_ref,
        before=before,
        after=before,
        status="skipped",
    )


def _require_formula(formula: str | None, op_name: str) -> str:
    """Require a non-null formula string."""
    if formula is None:
        raise ValueError(f"{op_name} requires formula.")
    return formula


def _openpyxl_cell_value(cell: OpenpyxlCellProtocol) -> PatchValue | None:
    """Normalize an openpyxl cell value into PatchValue."""
    value = getattr(cell, "value", None)
    if value is None:
        return None
    data_type = getattr(cell, "data_type", None)
    if data_type == "f":
        text = _normalize_formula(value)
        return PatchValue(kind="formula", value=text)
    return PatchValue(kind="value", value=value)


def _normalize_formula(value: object) -> str:
    """Ensure formula string starts with '='."""
    text = str(value)
    return text if text.startswith("=") else f"={text}"


def _expand_range_coordinates(range_ref: str) -> list[list[str]]:
    """Expand A1 range string into a 2D list of coordinates."""
    try:
        from openpyxl.utils.cell import get_column_letter, range_boundaries
    except ImportError as exc:
        raise RuntimeError(f"openpyxl is not available: {exc}") from exc
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    if min_col > max_col or min_row > max_row:
        raise ValueError(f"Invalid range reference: {range_ref}")
    rows: list[list[str]] = []
    for row_idx in range(min_row, max_row + 1):
        row: list[str] = []
        for col_idx in range(min_col, max_col + 1):
            row.append(f"{get_column_letter(col_idx)}{row_idx}")
        rows.append(row)
    return rows


def _shape_of_coordinates(coordinates: list[list[str]]) -> tuple[int, int]:
    """Return rows/cols for expanded coordinates."""
    if not coordinates or not coordinates[0]:
        raise ValueError("Range expansion resulted in an empty coordinate set.")
    return len(coordinates), len(coordinates[0])


def _translate_formula(formula: str, origin: str, target: str) -> str:
    """Translate formula with relative references from origin to target."""
    try:
        from openpyxl.formula.translate import Translator
    except ImportError as exc:
        raise RuntimeError(f"openpyxl is not available: {exc}") from exc
    translated = Translator(formula, origin=origin).translate_formula(target)
    return str(translated)


def _patch_value_to_primitive(value: PatchValue | None) -> str | int | float | None:
    """Convert PatchValue into primitive value for condition checks."""
    if value is None:
        return None
    return value.value


def _values_equal_for_condition(
    current: str | int | float | None,
    expected: str | int | float | None,
) -> bool:
    """Compare values for conditional update checks."""
    return current == expected


def _build_inverse_cell_op(
    op: PatchOp,
    cell_ref: str,
    before: PatchValue | None,
) -> PatchOp | None:
    """Build inverse operation for single-cell updates."""
    if op.op not in {"set_value", "set_formula", "set_value_if", "set_formula_if"}:
        return None
    if before is None:
        return PatchOp(op="set_value", sheet=op.sheet, cell=cell_ref, value=None)
    if before.kind == "formula":
        return PatchOp(
            op="set_formula",
            sheet=op.sheet,
            cell=cell_ref,
            formula=str(before.value),
        )
    return PatchOp(op="set_value", sheet=op.sheet, cell=cell_ref, value=before.value)


def _collect_formula_issues_openpyxl(
    workbook: OpenpyxlWorkbookProtocol,
) -> list[FormulaIssue]:
    """Collect simple formula issues by scanning formula text."""
    token_map: dict[str, tuple[FormulaIssueCode, FormulaIssueLevel]] = {
        "#REF!": ("ref_error", "error"),
        "#NAME?": ("name_error", "error"),
        "#DIV/0!": ("div0_error", "error"),
        "#VALUE!": ("value_error", "error"),
        "#N/A": ("na_error", "warning"),
    }
    issues: list[FormulaIssue] = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        iter_rows = getattr(sheet, "iter_rows", None)
        if iter_rows is None:
            continue
        for row in iter_rows():
            for cell in row:
                raw = getattr(cell, "value", None)
                if not isinstance(raw, str) or not raw.startswith("="):
                    continue
                normalized = raw.upper()
                if "==" in normalized:
                    issues.append(
                        FormulaIssue(
                            sheet=sheet_name,
                            cell=str(getattr(cell, "coordinate", "")),
                            level="warning",
                            code="invalid_token",
                            message="Formula contains duplicated '=' token.",
                        )
                    )
                for token, (code, level) in token_map.items():
                    if token in normalized:
                        issues.append(
                            FormulaIssue(
                                sheet=sheet_name,
                                cell=str(getattr(cell, "coordinate", "")),
                                level=level,
                                code=code,
                                message=f"Formula contains error token {token}.",
                            )
                        )
    return issues


def _apply_ops_xlwings(
    input_path: Path,
    output_path: Path,
    ops: list[PatchOp],
    auto_formula: bool,
) -> list[PatchDiffItem]:
    """Apply operations using Excel COM via xlwings."""
    diff: list[PatchDiffItem] = []
    try:
        with _xlwings_workbook(input_path) as workbook:
            sheets = {sheet.name: sheet for sheet in workbook.sheets}
            for index, op in enumerate(ops):
                try:
                    diff.append(
                        _apply_xlwings_op(workbook, sheets, op, index, auto_formula)
                    )
                except ValueError as exc:
                    raise PatchOpError.from_op(index, op, exc) from exc
            workbook.save(str(output_path))
    except ValueError:
        raise
    except Exception as exc:
        raise RuntimeError(f"COM patch failed: {exc}") from exc
    return diff


def _apply_xlwings_op(
    workbook: XlwingsWorkbookProtocol,
    sheets: dict[str, XlwingsSheetProtocol],
    op: PatchOp,
    index: int,
    auto_formula: bool,
) -> PatchDiffItem:
    """Apply a single op to an xlwings workbook."""
    # Extended ops are routed to openpyxl by _requires_openpyxl_backend.
    # Keep this explicit guard to prevent accidental path regressions.
    if op.op not in {"add_sheet", "set_value", "set_formula"}:
        raise ValueError(f"Unsupported op: {op.op}")
    if op.op == "add_sheet":
        if op.sheet in sheets:
            raise ValueError(f"Sheet already exists: {op.sheet}")
        last = workbook.sheets[-1] if workbook.sheets else None
        sheet = workbook.sheets.add(name=op.sheet, after=last)
        sheets[op.sheet] = sheet
        return PatchDiffItem(
            op_index=index,
            op=op.op,
            sheet=op.sheet,
            cell=None,
            before=None,
            after=PatchValue(kind="sheet", value=op.sheet),
        )

    existing_sheet = sheets.get(op.sheet)
    if existing_sheet is None:
        raise ValueError(f"Sheet not found: {op.sheet}")
    cell_ref = op.cell
    if cell_ref is None:
        raise ValueError(f"{op.op} requires cell.")
    rng = existing_sheet.range(cell_ref)
    before = _xlwings_cell_value(rng)
    if op.op == "set_value":
        if isinstance(op.value, str) and op.value.startswith("="):
            if not auto_formula:
                raise ValueError("set_value rejects values starting with '='.")
            rng.formula = op.value
            after = PatchValue(kind="formula", value=op.value)
        else:
            rng.value = op.value
            after = PatchValue(kind="value", value=op.value)
        return PatchDiffItem(
            op_index=index,
            op=op.op,
            sheet=op.sheet,
            cell=cell_ref,
            before=before,
            after=after,
        )
    if op.op == "set_formula":
        formula = op.formula
        if formula is None:
            raise ValueError("set_formula requires formula.")
        rng.formula = formula
        after = PatchValue(kind="formula", value=formula)
        return PatchDiffItem(
            op_index=index,
            op=op.op,
            sheet=op.sheet,
            cell=cell_ref,
            before=before,
            after=after,
        )
    raise ValueError(f"Unsupported op: {op.op}")


def _xlwings_cell_value(cell: XlwingsRangeProtocol) -> PatchValue | None:
    """Normalize an xlwings cell value into PatchValue."""
    formula = getattr(cell, "formula", None)
    if isinstance(formula, str) and formula.startswith("="):
        return PatchValue(kind="formula", value=formula)
    value = getattr(cell, "value", None)
    if value is None:
        return None
    return PatchValue(kind="value", value=value)


@contextmanager
def _xlwings_workbook(file_path: Path) -> Iterator[XlwingsWorkbookProtocol]:
    """Open an Excel workbook with a dedicated COM app."""
    app = xw.App(add_book=False, visible=False)
    app.display_alerts = False
    app.screen_updating = False
    workbook = app.books.open(str(file_path))
    try:
        yield workbook
    finally:
        try:
            workbook.close()
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            try:
                app.kill()
            except Exception:
                pass


class PatchOpError(ValueError):
    """Patch operation error with structured detail."""

    def __init__(self, detail: PatchErrorDetail) -> None:
        super().__init__(detail.message)
        self.detail = detail

    @classmethod
    def from_op(cls, index: int, op: PatchOp, exc: Exception) -> PatchOpError:
        """Build a PatchOpError from an op and exception."""
        detail = PatchErrorDetail(
            op_index=index,
            op=op.op,
            sheet=op.sheet,
            cell=op.cell,
            message=str(exc),
        )
        return cls(detail)
