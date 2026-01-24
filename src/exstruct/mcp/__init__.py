"""MCP server integration for ExStruct."""

from __future__ import annotations

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
    WorkbookMeta,
    run_extract,
)
from .io import PathPolicy
from .tools import (
    ExtractToolInput,
    ExtractToolOutput,
    ReadJsonChunkToolInput,
    ReadJsonChunkToolOutput,
    ValidateInputToolInput,
    ValidateInputToolOutput,
    run_extract_tool,
    run_read_json_chunk_tool,
    run_validate_input_tool,
)
from .validate_input import (
    ValidateInputRequest,
    ValidateInputResult,
    validate_input,
)

__all__ = [
    "ExtractRequest",
    "ExtractResult",
    "ExtractOptions",
    "ExtractToolInput",
    "ExtractToolOutput",
    "PathPolicy",
    "ReadJsonChunkFilter",
    "ReadJsonChunkRequest",
    "ReadJsonChunkResult",
    "ReadJsonChunkToolInput",
    "ReadJsonChunkToolOutput",
    "ValidateInputRequest",
    "ValidateInputResult",
    "ValidateInputToolInput",
    "ValidateInputToolOutput",
    "WorkbookMeta",
    "read_json_chunk",
    "validate_input",
    "run_extract",
    "run_extract_tool",
    "run_read_json_chunk_tool",
    "run_validate_input_tool",
]
