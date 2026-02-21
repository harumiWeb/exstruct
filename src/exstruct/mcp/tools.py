from __future__ import annotations

from pathlib import Path
from typing import Literal

from pydantic import BaseModel, Field

from exstruct import ExtractionMode

from .chunk_reader import (
    ReadJsonChunkFilter,
    ReadJsonChunkRequest,
    ReadJsonChunkResult,
    read_json_chunk,
)
from .extract_runner import (
    ExtractOptions,
    ExtractRequest,
    ExtractResult,
    OnConflictPolicy,
    WorkbookMeta,
    run_extract,
)
from .io import PathPolicy
from .patch_runner import (
    FormulaIssue,
    MakeRequest,
    PatchDiffItem,
    PatchErrorDetail,
    PatchOp,
    PatchRequest,
    PatchResult,
    run_make,
    run_patch,
)
from .sheet_reader import (
    CellReadItem,
    FormulaReadItem,
    ReadCellsRequest,
    ReadCellsResult,
    ReadFormulasRequest,
    ReadFormulasResult,
    ReadRangeRequest,
    ReadRangeResult,
    read_cells,
    read_formulas,
    read_range,
)
from .validate_input import (
    ValidateInputRequest,
    ValidateInputResult,
    validate_input,
)


class ExtractToolInput(BaseModel):
    """MCP tool input for ExStruct extraction."""

    xlsx_path: str
    mode: ExtractionMode = "standard"
    format: Literal["json", "yaml", "yml", "toon"] = "json"  # noqa: A003
    out_dir: str | None = None
    out_name: str | None = None
    on_conflict: OnConflictPolicy | None = None
    options: ExtractOptions = Field(default_factory=ExtractOptions)


class ExtractToolOutput(BaseModel):
    """MCP tool output for ExStruct extraction."""

    out_path: str
    workbook_meta: WorkbookMeta | None = None
    warnings: list[str] = Field(default_factory=list)
    engine: Literal["internal_api", "cli_subprocess"] = "internal_api"


class ReadJsonChunkToolInput(BaseModel):
    """MCP tool input for JSON chunk reading."""

    out_path: str
    sheet: str | None = None
    max_bytes: int = Field(default=50_000, ge=1)
    filter: ReadJsonChunkFilter | None = Field(default=None)  # noqa: A003
    cursor: str | None = None


class ReadJsonChunkToolOutput(BaseModel):
    """MCP tool output for JSON chunk reading."""

    chunk: str
    next_cursor: str | None = None
    warnings: list[str] = Field(default_factory=list)


class ReadRangeToolInput(BaseModel):
    """MCP tool input for range reading."""

    out_path: str
    sheet: str | None = None
    range: str = Field(...)  # noqa: A003
    include_formulas: bool = False
    include_empty: bool = True
    max_cells: int = Field(default=10_000, ge=1)


class ReadRangeToolOutput(BaseModel):
    """MCP tool output for range reading."""

    book_name: str | None = None
    sheet_name: str
    range: str  # noqa: A003
    cells: list[CellReadItem] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ReadCellsToolInput(BaseModel):
    """MCP tool input for cell list reading."""

    out_path: str
    sheet: str | None = None
    addresses: list[str] = Field(min_length=1)
    include_formulas: bool = True


class ReadCellsToolOutput(BaseModel):
    """MCP tool output for cell list reading."""

    book_name: str | None = None
    sheet_name: str
    cells: list[CellReadItem] = Field(default_factory=list)
    missing_cells: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ReadFormulasToolInput(BaseModel):
    """MCP tool input for formula reading."""

    out_path: str
    sheet: str | None = None
    range: str | None = None  # noqa: A003
    include_values: bool = False


class ReadFormulasToolOutput(BaseModel):
    """MCP tool output for formula reading."""

    book_name: str | None = None
    sheet_name: str
    range: str | None = None  # noqa: A003
    formulas: list[FormulaReadItem] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class ValidateInputToolInput(BaseModel):
    """MCP tool input for validating Excel files."""

    xlsx_path: str


class ValidateInputToolOutput(BaseModel):
    """MCP tool output for validating Excel files."""

    is_readable: bool
    warnings: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)


class RuntimePathExamples(BaseModel):
    """Path examples for MCP runtime diagnostics."""

    relative: str
    absolute: str


class RuntimeInfoToolOutput(BaseModel):
    """MCP tool output for runtime environment information."""

    root: str
    cwd: str
    platform: str
    path_examples: RuntimePathExamples


class PatchToolInput(BaseModel):
    """MCP tool input for patching Excel files."""

    xlsx_path: str
    ops: list[PatchOp]
    out_dir: str | None = None
    out_name: str | None = None
    on_conflict: OnConflictPolicy | None = None
    auto_formula: bool = False
    dry_run: bool = False
    return_inverse_ops: bool = False
    preflight_formula_check: bool = False
    backend: Literal["auto", "com", "openpyxl"] = "auto"


class MakeToolInput(BaseModel):
    """MCP tool input for creating and patching new Excel files."""

    out_path: str
    ops: list[PatchOp] = Field(default_factory=list)
    on_conflict: OnConflictPolicy | None = None
    auto_formula: bool = False
    dry_run: bool = False
    return_inverse_ops: bool = False
    preflight_formula_check: bool = False
    backend: Literal["auto", "com", "openpyxl"] = "auto"


class PatchToolOutput(BaseModel):
    """MCP tool output for patching Excel files."""

    out_path: str
    patch_diff: list[PatchDiffItem] = Field(default_factory=list)
    inverse_ops: list[PatchOp] = Field(default_factory=list)
    formula_issues: list[FormulaIssue] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    error: PatchErrorDetail | None = None
    engine: Literal["com", "openpyxl"]


class MakeToolOutput(BaseModel):
    """MCP tool output for workbook creation and patching."""

    out_path: str
    patch_diff: list[PatchDiffItem] = Field(default_factory=list)
    inverse_ops: list[PatchOp] = Field(default_factory=list)
    formula_issues: list[FormulaIssue] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    error: PatchErrorDetail | None = None
    engine: Literal["com", "openpyxl"]


def run_extract_tool(
    payload: ExtractToolInput,
    *,
    policy: PathPolicy | None = None,
    on_conflict: OnConflictPolicy | None = None,
) -> ExtractToolOutput:
    """Run the extraction tool handler.

    Args:
        payload: Tool input payload.
        policy: Optional path policy for access control.

    Returns:
        Tool output payload.
    """
    request = ExtractRequest(
        xlsx_path=Path(payload.xlsx_path),
        mode=payload.mode,
        format=payload.format,
        out_dir=Path(payload.out_dir) if payload.out_dir else None,
        out_name=payload.out_name,
        on_conflict=payload.on_conflict or on_conflict or "overwrite",
        options=payload.options,
    )
    result = run_extract(request, policy=policy)
    return _to_tool_output(result)


def run_read_json_chunk_tool(
    payload: ReadJsonChunkToolInput, *, policy: PathPolicy | None = None
) -> ReadJsonChunkToolOutput:
    """Run the JSON chunk tool handler.

    Args:
        payload: Tool input payload.
        policy: Optional path policy for access control.

    Returns:
        Tool output payload.
    """
    request = ReadJsonChunkRequest(
        out_path=Path(payload.out_path),
        sheet=payload.sheet,
        max_bytes=payload.max_bytes,
        filter=payload.filter,
        cursor=payload.cursor,
    )
    result = read_json_chunk(request, policy=policy)
    return _to_read_json_chunk_output(result)


def run_read_range_tool(
    payload: ReadRangeToolInput, *, policy: PathPolicy | None = None
) -> ReadRangeToolOutput:
    """Run the range read tool handler.

    Args:
        payload: Tool input payload.
        policy: Optional path policy for access control.

    Returns:
        Tool output payload.
    """
    request = ReadRangeRequest(
        out_path=Path(payload.out_path),
        sheet=payload.sheet,
        range=payload.range,
        include_formulas=payload.include_formulas,
        include_empty=payload.include_empty,
        max_cells=payload.max_cells,
    )
    result = read_range(request, policy=policy)
    return _to_read_range_tool_output(result)


def run_read_cells_tool(
    payload: ReadCellsToolInput, *, policy: PathPolicy | None = None
) -> ReadCellsToolOutput:
    """Run the cell list read tool handler.

    Args:
        payload: Tool input payload.
        policy: Optional path policy for access control.

    Returns:
        Tool output payload.
    """
    request = ReadCellsRequest(
        out_path=Path(payload.out_path),
        sheet=payload.sheet,
        addresses=payload.addresses,
        include_formulas=payload.include_formulas,
    )
    result = read_cells(request, policy=policy)
    return _to_read_cells_tool_output(result)


def run_read_formulas_tool(
    payload: ReadFormulasToolInput, *, policy: PathPolicy | None = None
) -> ReadFormulasToolOutput:
    """Run the formulas read tool handler.

    Args:
        payload: Tool input payload.
        policy: Optional path policy for access control.

    Returns:
        Tool output payload.
    """
    request = ReadFormulasRequest(
        out_path=Path(payload.out_path),
        sheet=payload.sheet,
        range=payload.range,
        include_values=payload.include_values,
    )
    result = read_formulas(request, policy=policy)
    return _to_read_formulas_tool_output(result)


def run_validate_input_tool(
    payload: ValidateInputToolInput, *, policy: PathPolicy | None = None
) -> ValidateInputToolOutput:
    """Run the validate input tool handler.

    Args:
        payload: Tool input payload.
        policy: Optional path policy for access control.

    Returns:
        Tool output payload.
    """
    request = ValidateInputRequest(xlsx_path=Path(payload.xlsx_path))
    result = validate_input(request, policy=policy)
    return _to_validate_input_output(result)


def run_patch_tool(
    payload: PatchToolInput,
    *,
    policy: PathPolicy | None = None,
    on_conflict: OnConflictPolicy | None = None,
) -> PatchToolOutput:
    """Run the patch tool handler.

    Args:
        payload: Tool input payload.
        policy: Optional path policy for access control.
        on_conflict: Optional conflict policy override.

    Returns:
        Tool output payload.
    """
    request = PatchRequest(
        xlsx_path=Path(payload.xlsx_path),
        ops=payload.ops,
        out_dir=Path(payload.out_dir) if payload.out_dir else None,
        out_name=payload.out_name,
        on_conflict=payload.on_conflict or on_conflict or "overwrite",
        auto_formula=payload.auto_formula,
        dry_run=payload.dry_run,
        return_inverse_ops=payload.return_inverse_ops,
        preflight_formula_check=payload.preflight_formula_check,
        backend=payload.backend,
    )
    result = run_patch(request, policy=policy)
    return _to_patch_tool_output(result)


def run_make_tool(
    payload: MakeToolInput,
    *,
    policy: PathPolicy | None = None,
    on_conflict: OnConflictPolicy | None = None,
) -> MakeToolOutput:
    """Run the make tool handler.

    Args:
        payload: Tool input payload.
        policy: Optional path policy for access control.
        on_conflict: Optional conflict policy override.

    Returns:
        Tool output payload.
    """
    request = MakeRequest(
        out_path=Path(payload.out_path),
        ops=payload.ops,
        on_conflict=payload.on_conflict or on_conflict or "overwrite",
        auto_formula=payload.auto_formula,
        dry_run=payload.dry_run,
        return_inverse_ops=payload.return_inverse_ops,
        preflight_formula_check=payload.preflight_formula_check,
        backend=payload.backend,
    )
    result = run_make(request, policy=policy)
    return _to_make_tool_output(result)


def _to_tool_output(result: ExtractResult) -> ExtractToolOutput:
    """Convert internal result to tool output model.

    Args:
        result: Internal extraction result.

    Returns:
        Tool output payload.
    """
    return ExtractToolOutput(
        out_path=result.out_path,
        workbook_meta=result.workbook_meta,
        warnings=result.warnings,
        engine=result.engine,
    )


def _to_read_json_chunk_output(
    result: ReadJsonChunkResult,
) -> ReadJsonChunkToolOutput:
    """Convert internal result to JSON chunk tool output.

    Args:
        result: Internal chunk result.

    Returns:
        Tool output payload.
    """
    return ReadJsonChunkToolOutput(
        chunk=result.chunk,
        next_cursor=result.next_cursor,
        warnings=result.warnings,
    )


def _to_read_range_tool_output(result: ReadRangeResult) -> ReadRangeToolOutput:
    """Convert internal result to range tool output.

    Args:
        result: Internal range read result.

    Returns:
        Tool output payload.
    """
    return ReadRangeToolOutput(
        book_name=result.book_name,
        sheet_name=result.sheet_name,
        range=result.range,
        cells=result.cells,
        warnings=result.warnings,
    )


def _to_read_cells_tool_output(result: ReadCellsResult) -> ReadCellsToolOutput:
    """Convert internal result to cell list tool output.

    Args:
        result: Internal cell list read result.

    Returns:
        Tool output payload.
    """
    return ReadCellsToolOutput(
        book_name=result.book_name,
        sheet_name=result.sheet_name,
        cells=result.cells,
        missing_cells=result.missing_cells,
        warnings=result.warnings,
    )


def _to_read_formulas_tool_output(
    result: ReadFormulasResult,
) -> ReadFormulasToolOutput:
    """Convert internal result to formulas tool output.

    Args:
        result: Internal formula read result.

    Returns:
        Tool output payload.
    """
    return ReadFormulasToolOutput(
        book_name=result.book_name,
        sheet_name=result.sheet_name,
        range=result.range,
        formulas=result.formulas,
        warnings=result.warnings,
    )


def _to_validate_input_output(
    result: ValidateInputResult,
) -> ValidateInputToolOutput:
    """Convert internal result to validate input tool output.

    Args:
        result: Internal validation result.

    Returns:
        Tool output payload.
    """
    return ValidateInputToolOutput(
        is_readable=result.is_readable,
        warnings=result.warnings,
        errors=result.errors,
    )


def _to_patch_tool_output(result: PatchResult) -> PatchToolOutput:
    """Convert internal result to patch tool output.

    Args:
        result: Internal patch result.

    Returns:
        Tool output payload.
    """
    return PatchToolOutput(
        out_path=result.out_path,
        patch_diff=result.patch_diff,
        inverse_ops=result.inverse_ops,
        formula_issues=result.formula_issues,
        warnings=result.warnings,
        error=result.error,
        engine=result.engine,
    )


def _to_make_tool_output(result: PatchResult) -> MakeToolOutput:
    """Convert internal result to make tool output.

    Args:
        result: Internal make result.

    Returns:
        Tool output payload.
    """
    return MakeToolOutput(
        out_path=result.out_path,
        patch_diff=result.patch_diff,
        inverse_ops=result.inverse_ops,
        formula_issues=result.formula_issues,
        warnings=result.warnings,
        error=result.error,
        engine=result.engine,
    )
