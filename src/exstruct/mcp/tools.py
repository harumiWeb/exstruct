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


class ValidateInputToolInput(BaseModel):
    """MCP tool input for validating Excel files."""

    xlsx_path: str


class ValidateInputToolOutput(BaseModel):
    """MCP tool output for validating Excel files."""

    is_readable: bool
    warnings: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)


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
