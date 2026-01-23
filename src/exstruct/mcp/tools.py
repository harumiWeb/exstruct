from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, Field

from exstruct import ExtractionMode

from .extract_runner import ExtractRequest, ExtractResult, WorkbookMeta, run_extract
from .io import PathPolicy


class ExtractToolInput(BaseModel):
    """MCP tool input for ExStruct extraction."""

    xlsx_path: str
    mode: ExtractionMode = "standard"
    format: Literal["json", "yaml", "yml", "toon"] = "json"  # noqa: A003
    out_dir: str | None = None
    out_name: str | None = None
    options: dict[str, Any] = Field(default_factory=dict)


class ExtractToolOutput(BaseModel):
    """MCP tool output for ExStruct extraction."""

    out_path: str
    workbook_meta: WorkbookMeta | None = None
    warnings: list[str] = Field(default_factory=list)
    engine: Literal["internal_api", "cli_subprocess"] = "internal_api"


def run_extract_tool(
    payload: ExtractToolInput, *, policy: PathPolicy | None = None
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
        options=payload.options,
    )
    result = run_extract(request, policy=policy)
    return _to_tool_output(result)


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
