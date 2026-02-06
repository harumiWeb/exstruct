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

PatchOpType = Literal["set_value", "set_formula", "add_sheet"]
PatchStatus = Literal["applied", "skipped"]
PatchValueKind = Literal["value", "formula", "sheet"]

_ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
_A1_PATTERN = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*$")


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
    """Single patch operation for an Excel workbook."""

    op: PatchOpType
    sheet: str
    cell: str | None = None
    value: str | int | float | None = None
    formula: str | None = None

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

    @model_validator(mode="after")
    def _validate_op(self) -> PatchOp:
        if self.op == "add_sheet":
            _validate_add_sheet(self)
            return self
        _validate_cell_required(self)
        if self.op == "set_value":
            _validate_set_value(self)
            return self
        if self.op == "set_formula":
            _validate_set_formula(self)
            return self
        return self


def _validate_add_sheet(op: PatchOp) -> None:
    """Validate add_sheet operation."""
    if op.cell is not None:
        raise ValueError("add_sheet does not accept cell.")
    if op.value is not None:
        raise ValueError("add_sheet does not accept value.")
    if op.formula is not None:
        raise ValueError("add_sheet does not accept formula.")


def _validate_cell_required(op: PatchOp) -> None:
    """Validate that the operation has a cell value."""
    if op.cell is None:
        raise ValueError(f"{op.op} requires cell.")


def _validate_set_value(op: PatchOp) -> None:
    """Validate set_value operation."""
    if op.formula is not None:
        raise ValueError("set_value does not accept formula.")


def _validate_set_formula(op: PatchOp) -> None:
    """Validate set_formula operation."""
    if op.value is not None:
        raise ValueError("set_formula does not accept value.")
    if op.formula is None:
        raise ValueError("set_formula requires formula.")
    if not op.formula.startswith("="):
        raise ValueError("set_formula requires formula starting with '='.")


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


class PatchRequest(BaseModel):
    """Input model for ExStruct MCP patch."""

    xlsx_path: Path
    ops: list[PatchOp]
    out_dir: Path | None = None
    out_name: str | None = None
    on_conflict: OnConflictPolicy = "rename"
    auto_formula: bool = False


class PatchResult(BaseModel):
    """Output model for ExStruct MCP patch."""

    out_path: str
    patch_diff: list[PatchDiffItem] = Field(default_factory=list)
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
    if skipped:
        return PatchResult(out_path=str(output_path), patch_diff=[], warnings=warnings)

    com = get_com_availability()
    if resolved_input.suffix.lower() == ".xls" and not com.available:
        raise ValueError(
            ".xls editing requires Windows Excel COM (xlwings) in this environment."
        )

    _ensure_output_dir(output_path)
    if com.available:
        try:
            diff = _apply_ops_xlwings(
                resolved_input,
                output_path,
                request.ops,
                request.auto_formula,
            )
            return PatchResult(
                out_path=str(output_path), patch_diff=diff, warnings=warnings
            )
        except PatchOpError as exc:
            return PatchResult(
                out_path=str(output_path),
                patch_diff=[],
                warnings=warnings,
                error=exc.detail,
            )
        except Exception as exc:
            fallback = _maybe_fallback_openpyxl(
                resolved_input,
                output_path,
                request.ops,
                warnings,
                reason=f"COM patch failed; falling back to openpyxl. ({exc!r})",
                auto_formula=request.auto_formula,
            )
            if fallback is not None:
                return fallback
            raise RuntimeError(f"COM patch failed: {exc}") from exc

    if com.reason:
        warnings.append(f"COM unavailable: {com.reason}")
    return _apply_with_openpyxl(
        resolved_input,
        output_path,
        request.ops,
        warnings,
        auto_formula=request.auto_formula,
    )


def _apply_with_openpyxl(
    input_path: Path,
    output_path: Path,
    ops: list[PatchOp],
    warnings: list[str],
    *,
    auto_formula: bool,
) -> PatchResult:
    """Apply patch operations using openpyxl."""
    try:
        diff = _apply_ops_openpyxl(input_path, output_path, ops, auto_formula)
    except PatchOpError as exc:
        return PatchResult(
            out_path=str(output_path),
            patch_diff=[],
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

    warnings.append("openpyxl editing may drop shapes/charts or unsupported elements.")
    return PatchResult(out_path=str(output_path), patch_diff=diff, warnings=warnings)


def _maybe_fallback_openpyxl(
    input_path: Path,
    output_path: Path,
    ops: list[PatchOp],
    warnings: list[str],
    *,
    reason: str,
    auto_formula: bool,
) -> PatchResult | None:
    """Attempt openpyxl fallback after COM failure."""
    if input_path.suffix.lower() == ".xls":
        warnings.append(reason)
        return None
    warnings.append(reason)
    return _apply_with_openpyxl(
        input_path,
        output_path,
        ops,
        warnings,
        auto_formula=auto_formula,
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
    input_path: Path,
    output_path: Path,
    ops: list[PatchOp],
    auto_formula: bool,
) -> list[PatchDiffItem]:
    """Apply operations using openpyxl."""
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError(f"openpyxl is not available: {exc}") from exc

    if input_path.suffix.lower() == ".xls":
        raise ValueError("openpyxl cannot edit .xls files.")

    workbook = load_workbook(input_path)
    try:
        diff = _apply_ops_to_openpyxl_workbook(workbook, ops, auto_formula)
        workbook.save(output_path)
    finally:
        workbook.close()
    return diff


def _apply_ops_to_openpyxl_workbook(
    workbook: OpenpyxlWorkbookProtocol, ops: list[PatchOp], auto_formula: bool
) -> list[PatchDiffItem]:
    """Apply ops to an openpyxl workbook instance."""
    sheets = _openpyxl_sheet_map(workbook)
    diff: list[PatchDiffItem] = []
    for index, op in enumerate(ops):
        try:
            diff.append(_apply_openpyxl_op(workbook, sheets, op, index, auto_formula))
        except ValueError as exc:
            raise PatchOpError.from_op(index, op, exc) from exc
    return diff


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
) -> PatchDiffItem:
    """Apply a single op to openpyxl workbook."""
    if op.op == "add_sheet":
        if op.sheet in sheets:
            raise ValueError(f"Sheet already exists: {op.sheet}")
        sheet = workbook.create_sheet(title=op.sheet)
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
    cell = existing_sheet[cell_ref]
    before = _openpyxl_cell_value(cell)
    if op.op == "set_value":
        if isinstance(op.value, str) and op.value.startswith("="):
            if not auto_formula:
                raise ValueError("set_value rejects values starting with '='.")
            cell.value = op.value
            after = PatchValue(kind="formula", value=op.value)
        else:
            cell.value = op.value
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
        cell.value = formula
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
