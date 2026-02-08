from __future__ import annotations

import argparse
import functools
import importlib
import logging
import os
from pathlib import Path
from types import ModuleType
from typing import TYPE_CHECKING, Any, Literal, cast

import anyio
from pydantic import BaseModel, Field

from exstruct import ExtractionMode

from .extract_runner import OnConflictPolicy
from .io import PathPolicy
from .patch_runner import PatchOp
from .tools import (
    ExtractToolInput,
    ExtractToolOutput,
    PatchToolInput,
    PatchToolOutput,
    ReadJsonChunkToolInput,
    ReadJsonChunkToolOutput,
    ValidateInputToolInput,
    ValidateInputToolOutput,
    run_extract_tool,
    run_patch_tool,
    run_read_json_chunk_tool,
    run_validate_input_tool,
)

if TYPE_CHECKING:  # pragma: no cover - typing only
    from mcp.server.fastmcp import FastMCP

logger = logging.getLogger(__name__)


class ServerConfig(BaseModel):
    """Configuration for the MCP server process."""

    root: Path = Field(..., description="Root directory for file access.")
    deny_globs: list[str] = Field(default_factory=list, description="Denied glob list.")
    log_level: str = Field(default="INFO", description="Logging level.")
    log_file: Path | None = Field(default=None, description="Optional log file path.")
    on_conflict: OnConflictPolicy = Field(
        default="overwrite", description="Output conflict policy."
    )
    warmup: bool = Field(default=False, description="Warm up heavy imports on start.")


def main(argv: list[str] | None = None) -> int:
    """Run the MCP server entrypoint.

    Args:
        argv: Optional CLI arguments for testing.

    Returns:
        Exit code (0 for success, 1 for failure).
    """
    config = _parse_args(argv)
    _configure_logging(config)
    try:
        run_server(config)
    except Exception as exc:  # pragma: no cover - surface runtime errors
        logger.error("MCP server failed: %s", exc)
        return 1
    return 0


def run_server(config: ServerConfig) -> None:
    """Start the MCP server.

    Args:
        config: Server configuration.
    """
    os.environ.setdefault("EXSTRUCT_BORDER_CLUSTER_BACKEND", "python")
    logger.info(
        "Border cluster backend set to %s for MCP.",
        os.getenv("EXSTRUCT_BORDER_CLUSTER_BACKEND"),
    )
    _import_mcp()
    policy = PathPolicy(root=config.root, deny_globs=config.deny_globs)
    logger.info("MCP root: %s", policy.normalize_root())
    if config.warmup:
        _warmup_exstruct()
    app = _create_app(policy, on_conflict=config.on_conflict)
    app.run()


def _parse_args(argv: list[str] | None) -> ServerConfig:
    """Parse CLI arguments into server config.

    Args:
        argv: Optional CLI argument list.

    Returns:
        Parsed server configuration.
    """
    parser = argparse.ArgumentParser(description="ExStruct MCP server (stdio).")
    parser.add_argument("--root", type=Path, required=True, help="Workspace root.")
    parser.add_argument(
        "--deny-glob",
        action="append",
        default=[],
        help="Glob pattern to deny (can be specified multiple times).",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="Logging level (DEBUG, INFO, WARNING, ERROR).",
    )
    parser.add_argument("--log-file", type=Path, help="Optional log file path.")
    parser.add_argument(
        "--on-conflict",
        choices=["overwrite", "skip", "rename"],
        default="overwrite",
        help="Output conflict policy (overwrite/skip/rename).",
    )
    parser.add_argument(
        "--warmup",
        action="store_true",
        help="Warm up heavy imports on startup to reduce tool latency.",
    )
    args = parser.parse_args(argv)
    return ServerConfig(
        root=args.root,
        deny_globs=list(args.deny_glob),
        log_level=args.log_level,
        log_file=args.log_file,
        on_conflict=args.on_conflict,
        warmup=bool(args.warmup),
    )


def _configure_logging(config: ServerConfig) -> None:
    """Configure logging for the server process.

    Args:
        config: Server configuration.
    """
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    if config.log_file is not None:
        handlers.append(logging.FileHandler(config.log_file))
    logging.basicConfig(
        level=config.log_level.upper(),
        handlers=handlers,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


def _import_mcp() -> ModuleType:
    """Import the MCP SDK module or raise a helpful error.

    Returns:
        Imported MCP module.
    """
    try:
        return importlib.import_module("mcp")
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "MCP SDK is not installed. Install with `pip install exstruct[mcp]`."
        ) from exc


def _warmup_exstruct() -> None:
    """Warm up heavy imports to reduce first-call latency."""
    logger.info("Warming up ExStruct imports...")
    importlib.import_module("exstruct.core.cells")
    importlib.import_module("exstruct.core.integrate")
    logger.info("Warmup completed.")


def _create_app(policy: PathPolicy, *, on_conflict: OnConflictPolicy) -> FastMCP:
    """Create the MCP FastMCP application.

    Args:
        policy: Path policy for filesystem access.

    Returns:
        FastMCP application instance.
    """
    from mcp.server.fastmcp import FastMCP

    app = FastMCP("ExStruct MCP", json_response=True)
    _register_tools(app, policy, default_on_conflict=on_conflict)
    return app


def _register_tools(
    app: FastMCP, policy: PathPolicy, *, default_on_conflict: OnConflictPolicy
) -> None:
    """Register MCP tools for the server.

    Args:
        app: FastMCP application instance.
        policy: Path policy for filesystem access.
    """

    async def _extract_tool(  # pylint: disable=redefined-builtin
        xlsx_path: str,
        mode: ExtractionMode = "standard",
        format: Literal["json", "yaml", "yml", "toon"] = "json",  # noqa: A002
        out_dir: str | None = None,
        out_name: str | None = None,
        on_conflict: OnConflictPolicy | None = None,
        options: dict[str, Any] | None = None,
    ) -> ExtractToolOutput:
        """Handle the ExStruct extraction tool call.

        Args:
            xlsx_path: Path to the Excel workbook.
            mode: Extraction mode.
            format: Output format.
            out_dir: Optional output directory.
            out_name: Optional output filename.
            options: Additional options (reserved for future use).

        Returns:
            Extraction result payload.
        """
        payload = ExtractToolInput(
            xlsx_path=xlsx_path,
            mode=mode,
            format=format,
            out_dir=out_dir,
            out_name=out_name,
            on_conflict=on_conflict,
            options=options or {},
        )
        effective_on_conflict = on_conflict or default_on_conflict
        work = functools.partial(
            run_extract_tool,
            payload,
            policy=policy,
            on_conflict=effective_on_conflict,
        )
        result = cast(ExtractToolOutput, await anyio.to_thread.run_sync(work))
        return result

    tool = app.tool(name="exstruct_extract")
    tool(_extract_tool)

    async def _read_json_chunk_tool(  # pylint: disable=redefined-builtin
        out_path: str,
        sheet: str | None = None,
        max_bytes: int = 50_000,
        filter: dict[str, Any] | None = None,  # noqa: A002
        cursor: str | None = None,
    ) -> ReadJsonChunkToolOutput:
        """Handle JSON chunk tool call.

        Args:
            out_path: Path to the JSON output file.
            sheet: Optional sheet name.
            max_bytes: Maximum chunk size in bytes.
            filter: Optional filter payload.
            cursor: Optional cursor for pagination.

        Returns:
            JSON chunk result payload.
        """
        payload = ReadJsonChunkToolInput(
            out_path=out_path,
            sheet=sheet,
            max_bytes=max_bytes,
            filter=_coerce_filter(filter),
            cursor=cursor,
        )
        work = functools.partial(
            run_read_json_chunk_tool,
            payload,
            policy=policy,
        )
        result = cast(ReadJsonChunkToolOutput, await anyio.to_thread.run_sync(work))
        return result

    chunk_tool = app.tool(name="exstruct_read_json_chunk")
    chunk_tool(_read_json_chunk_tool)

    async def _validate_input_tool(xlsx_path: str) -> ValidateInputToolOutput:
        """Handle input validation tool call.

        Args:
            xlsx_path: Path to the Excel workbook.

        Returns:
            Validation result payload.
        """
        payload = ValidateInputToolInput(xlsx_path=xlsx_path)
        work = functools.partial(
            run_validate_input_tool,
            payload,
            policy=policy,
        )
        result = cast(ValidateInputToolOutput, await anyio.to_thread.run_sync(work))
        return result

    validate_tool = app.tool(name="exstruct_validate_input")
    validate_tool(_validate_input_tool)

    async def _patch_tool(
        xlsx_path: str,
        ops: list[PatchOp],
        out_dir: str | None = None,
        out_name: str | None = None,
        on_conflict: OnConflictPolicy | None = None,
        auto_formula: bool = False,
        dry_run: bool = False,
        return_inverse_ops: bool = False,
        preflight_formula_check: bool = False,
    ) -> PatchToolOutput:
        """Edit an Excel workbook by applying patch operations.

        Supports cell value updates, formula updates, and adding new sheets.
        Operations are applied atomically: all succeed or none are saved.

        Args:
            xlsx_path: Path to the Excel workbook to edit.
            ops: Patch operations to apply in order. Each op has an 'op' field
                specifying the type: 'set_value' (set cell value), 'set_formula'
                (set cell formula starting with '='), 'add_sheet' (create new sheet),
                'set_range_values' (bulk set rectangular range), 'fill_formula'
                (fill formula across a row/column), 'set_value_if' (conditional
                value update), 'set_formula_if' (conditional formula update).
            out_dir: Output directory. Defaults to same directory as input.
            out_name: Output filename. Defaults to '{stem}_patched{ext}'.
            on_conflict: Conflict policy when output file exists:
                'overwrite' (replace), 'skip' (do nothing), 'rename' (auto-rename).
                Defaults to server --on-conflict setting.
            auto_formula: When true, values starting with '=' in set_value ops
                are treated as formulas instead of being rejected.
            dry_run: When true, compute diff without saving changes.
            return_inverse_ops: When true, return inverse (undo) operations.
            preflight_formula_check: When true, scan formulas for errors
                like #REF!, #NAME?, #DIV/0! before saving.

        Returns:
            Patch result with output path, applied diffs, and any warnings.
        """
        payload = PatchToolInput(
            xlsx_path=xlsx_path,
            ops=ops,
            out_dir=out_dir,
            out_name=out_name,
            on_conflict=on_conflict,
            auto_formula=auto_formula,
            dry_run=dry_run,
            return_inverse_ops=return_inverse_ops,
            preflight_formula_check=preflight_formula_check,
        )
        effective_on_conflict = on_conflict or default_on_conflict
        work = functools.partial(
            run_patch_tool,
            payload,
            policy=policy,
            on_conflict=effective_on_conflict,
        )
        result = cast(PatchToolOutput, await anyio.to_thread.run_sync(work))
        return result

    patch_tool = app.tool(name="exstruct_patch")
    patch_tool(_patch_tool)


def _coerce_filter(filter_data: dict[str, Any] | None) -> dict[str, Any] | None:
    """Normalize filter input for chunk reading.

    Args:
        filter_data: Filter payload from MCP tool call.

    Returns:
        Normalized filter dict or None.
    """
    if not filter_data:
        return None
    return dict(filter_data)
