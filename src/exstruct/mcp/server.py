from __future__ import annotations

import argparse
import functools
import importlib
import logging
import os
from pathlib import Path
import sys
from types import ModuleType
from typing import TYPE_CHECKING, Any, Literal, cast

import anyio
from pydantic import BaseModel, Field

from exstruct import ExtractionMode

from .extract_runner import OnConflictPolicy
from .io import PathPolicy
from .op_schema import build_patch_tool_mini_schema
from .patch.normalize import (
    build_patch_op_error_message as _normalize_build_patch_op_error_message,
    coerce_patch_ops as _normalize_coerce_patch_ops,
    parse_patch_op_json as _normalize_parse_patch_op_json,
)
from .tools import (
    DescribeOpToolInput,
    DescribeOpToolOutput,
    ExtractToolInput,
    ExtractToolOutput,
    ListOpsToolOutput,
    MakeToolInput,
    MakeToolOutput,
    PatchToolInput,
    PatchToolOutput,
    ReadCellsToolInput,
    ReadCellsToolOutput,
    ReadFormulasToolInput,
    ReadFormulasToolOutput,
    ReadJsonChunkToolInput,
    ReadJsonChunkToolOutput,
    ReadRangeToolInput,
    ReadRangeToolOutput,
    RuntimeInfoToolOutput,
    ValidateInputToolInput,
    ValidateInputToolOutput,
    run_describe_op_tool,
    run_extract_tool,
    run_list_ops_tool,
    run_make_tool,
    run_patch_tool,
    run_read_cells_tool,
    run_read_formulas_tool,
    run_read_json_chunk_tool,
    run_read_range_tool,
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
    artifact_bridge_dir: Path | None = Field(
        default=None,
        description="Optional bridge directory for mirrored artifacts.",
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
    app = _create_app(
        policy,
        on_conflict=config.on_conflict,
        artifact_bridge_dir=config.artifact_bridge_dir,
    )
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
        "--artifact-bridge-dir",
        type=Path,
        help="Optional directory to mirror generated artifacts for chat handoff.",
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
        artifact_bridge_dir=args.artifact_bridge_dir,
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


def _create_app(
    policy: PathPolicy,
    *,
    on_conflict: OnConflictPolicy,
    artifact_bridge_dir: Path | None = None,
) -> FastMCP:
    """Create the MCP FastMCP application.

    Args:
        policy: Path policy for filesystem access.

    Returns:
        FastMCP application instance.
    """
    from mcp.server.fastmcp import FastMCP

    app = FastMCP("ExStruct MCP", json_response=True)
    _register_tools(
        app,
        policy,
        default_on_conflict=on_conflict,
        artifact_bridge_dir=artifact_bridge_dir,
    )
    return app


def _register_tools(
    app: FastMCP,
    policy: PathPolicy,
    *,
    default_on_conflict: OnConflictPolicy,
    artifact_bridge_dir: Path | None = None,
) -> None:
    """Register MCP tools for the server.

    Args:
        app: FastMCP application instance.
        policy: Path policy for filesystem access.
        default_on_conflict: Default conflict policy used when tool input omits it.
        artifact_bridge_dir: Optional directory for artifact mirroring handoff.
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
            mode: Extraction detail level. Allowed values are:
                "light" (cells + table candidates + print areas),
                "standard" (recommended default),
                "verbose" (adds richer metadata such as links/maps).
            format: Output format.
            out_dir: Optional output directory.
            out_name: Optional output filename.
            options: Additional options. Supports: pretty (bool), indent (int),
                sheets_dir (str), print_areas_dir (str), auto_page_breaks_dir (str),
                alpha_col (bool - convert column keys to Excel-style ABC names).

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
            sheet: Optional sheet name. Required when multiple sheets exist and
                no cursor/filter can disambiguate the target.
            max_bytes: Maximum chunk size in bytes. Start around 50_000 and
                increase (for example 120_000) when chunks are too small.
            filter: Optional filter payload with
                rows=[start,end], cols=[start,end] (1-based inclusive).
            cursor: Optional cursor for pagination (non-negative integer string).

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

    async def _read_range_tool(  # pylint: disable=redefined-builtin
        out_path: str,
        sheet: str | None = None,
        range: str = "A1",  # noqa: A002
        include_formulas: bool = False,
        include_empty: bool = True,
        max_cells: int = 10_000,
    ) -> ReadRangeToolOutput:
        """Read a rectangular cell range from extracted JSON.

        Args:
            out_path: Path to the extracted JSON file.
            sheet: Optional sheet name. Required when workbook has multiple sheets.
            range: A1 range (for example, \"A1:G10\").
            include_formulas: Include formula text for each returned cell.
            include_empty: Include empty cells in the range.
            max_cells: Safety limit for expanded cells.

        Returns:
            Range read result payload.
        """
        payload = ReadRangeToolInput(
            out_path=out_path,
            sheet=sheet,
            range=range,
            include_formulas=include_formulas,
            include_empty=include_empty,
            max_cells=max_cells,
        )
        work = functools.partial(
            run_read_range_tool,
            payload,
            policy=policy,
        )
        result = cast(ReadRangeToolOutput, await anyio.to_thread.run_sync(work))
        return result

    read_range_tool = app.tool(name="exstruct_read_range")
    read_range_tool(_read_range_tool)

    async def _read_cells_tool(
        out_path: str,
        addresses: list[str],
        sheet: str | None = None,
        include_formulas: bool = True,
    ) -> ReadCellsToolOutput:
        """Read specific cells from extracted JSON.

        Args:
            out_path: Path to the extracted JSON file.
            addresses: Target A1 addresses (for example, [\"J98\", \"J124\"]).
            sheet: Optional sheet name. Required when workbook has multiple sheets.
            include_formulas: Include formula text for each requested cell.

        Returns:
            Cell read result payload.
        """
        payload = ReadCellsToolInput(
            out_path=out_path,
            sheet=sheet,
            addresses=addresses,
            include_formulas=include_formulas,
        )
        work = functools.partial(
            run_read_cells_tool,
            payload,
            policy=policy,
        )
        result = cast(ReadCellsToolOutput, await anyio.to_thread.run_sync(work))
        return result

    read_cells_tool = app.tool(name="exstruct_read_cells")
    read_cells_tool(_read_cells_tool)

    async def _read_formulas_tool(  # pylint: disable=redefined-builtin
        out_path: str,
        sheet: str | None = None,
        range: str | None = None,  # noqa: A002
        include_values: bool = False,
    ) -> ReadFormulasToolOutput:
        """Read formulas from extracted JSON.

        Args:
            out_path: Path to the extracted JSON file.
            sheet: Optional sheet name. Required when workbook has multiple sheets.
            range: Optional A1 range to limit formula results (for example, \"J2:J201\").
            include_values: Include stored cell values with formulas.

        Returns:
            Formula read result payload.
        """
        payload = ReadFormulasToolInput(
            out_path=out_path,
            sheet=sheet,
            range=range,
            include_values=include_values,
        )
        work = functools.partial(
            run_read_formulas_tool,
            payload,
            policy=policy,
        )
        result = cast(ReadFormulasToolOutput, await anyio.to_thread.run_sync(work))
        return result

    read_formulas_tool = app.tool(name="exstruct_read_formulas")
    read_formulas_tool(_read_formulas_tool)

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

    async def _runtime_info_tool() -> RuntimeInfoToolOutput:
        """Return runtime diagnostics for MCP path troubleshooting.

        Returns:
            Runtime root/cwd/platform and valid path examples.
        """
        root = policy.normalize_root()
        return RuntimeInfoToolOutput(
            root=str(root),
            cwd=str(Path.cwd().resolve()),
            platform=sys.platform,
            path_examples={
                "relative": "outputs/book.xlsx",
                "absolute": str(root / "outputs" / "book.xlsx"),
            },
        )

    runtime_info_tool = app.tool(name="exstruct_get_runtime_info")
    runtime_info_tool(_runtime_info_tool)

    _register_op_schema_tools(app)

    async def _patch_tool(
        xlsx_path: str,
        ops: list[dict[str, Any] | str],
        sheet: str | None = None,
        out_dir: str | None = None,
        out_name: str | None = None,
        on_conflict: OnConflictPolicy | None = None,
        auto_formula: bool = False,
        dry_run: bool = False,
        return_inverse_ops: bool = False,
        preflight_formula_check: bool = False,
        backend: Literal["auto", "com", "openpyxl"] = "auto",
        mirror_artifact: bool = False,
    ) -> PatchToolOutput:
        """Edit an Excel workbook by applying patch operations.

        Supports cell value updates, formula updates, and adding new sheets.
        Operations are applied atomically: all succeed or none are saved.

        Args:
            xlsx_path: Path to the Excel workbook to edit.
            ops: Patch operations to apply in order. Preferred format is an
                object list (one object per operation). For compatibility with
                clients that cannot send object arrays, JSON object strings are
                also accepted and normalized before validation. Each operation
                has an 'op' field specifying the type: 'set_value' (set cell
                value), 'set_formula' (set cell formula starting with '='),
                'add_sheet' (create new sheet), 'set_range_values' (bulk set
                rectangular range), 'fill_formula' (fill formula across a
                row/column), 'set_value_if' (conditional value update),
                'set_formula_if' (conditional formula update),
                'draw_grid_border' (draw thin black grid border),
                'set_bold' (apply bold style),
                'set_font_size' (apply font size; requires font_size > 0 and exactly one of cell/range),
                'set_font_color' (apply font color; requires color and exactly one of cell/range),
                'set_fill_color' (apply solid fill),
                'set_dimensions' (set row height/column width),
                'auto_fit_columns' (auto-fit column widths with optional bounds),
                'merge_cells' (merge a rectangular range),
                'unmerge_cells' (unmerge ranges intersecting target),
                'set_alignment' (set horizontal/vertical alignment and wrap_text), and
                'set_style' (apply multiple style attributes in one op), and
                'apply_table_style' (create table and apply Excel table style), and
                'create_chart' (create a new chart; COM backend only), and
                'restore_design_snapshot' (internal inverse restore op).
            sheet: Optional default sheet name. Used when op.sheet is omitted
                for non-add_sheet ops. If both are set, op.sheet wins.
            out_dir: Output directory. Defaults to same directory as input.
            out_name: Output filename. Defaults to '{stem}_patched{ext}'.
                If stem already ends with '_patched', the same name is reused.
            on_conflict: Conflict policy when output file exists:
                'overwrite' (replace), 'skip' (do nothing), 'rename' (auto-rename).
                Defaults to server --on-conflict setting.
            auto_formula: When true, values starting with '=' in set_value ops
                are treated as formulas instead of being rejected.
            dry_run: When true, compute diff without saving changes.
            return_inverse_ops: When true, return inverse (undo) operations.
            preflight_formula_check: When true, scan formulas for errors
                like #REF!, #NAME?, #DIV/0! before saving.
            backend: Patch execution backend.
                - "auto" (default): prefer COM when available; otherwise openpyxl.
                  Uses openpyxl when dry_run/return_inverse_ops/preflight_formula_check
                  is enabled.
                - "com": force COM path (requires Excel COM and disallows
                  dry_run/return_inverse_ops/preflight_formula_check).
                - "openpyxl": force openpyxl path (.xls is not supported).
            mirror_artifact: When true, mirror output workbook to
                --artifact-bridge-dir after successful patch.

        Returns:
            Patch result with output path, applied diffs, and any warnings.
        """
        normalized_ops = _coerce_patch_ops(ops)
        payload = PatchToolInput(
            xlsx_path=xlsx_path,
            ops=normalized_ops,
            out_dir=out_dir,
            out_name=out_name,
            sheet=sheet,
            on_conflict=on_conflict,
            auto_formula=auto_formula,
            dry_run=dry_run,
            return_inverse_ops=return_inverse_ops,
            preflight_formula_check=preflight_formula_check,
            backend=backend,
            mirror_artifact=mirror_artifact,
        )
        effective_on_conflict = on_conflict or default_on_conflict
        if artifact_bridge_dir is None:
            work = functools.partial(
                run_patch_tool,
                payload,
                policy=policy,
                on_conflict=effective_on_conflict,
            )
        else:
            work = functools.partial(
                run_patch_tool,
                payload,
                policy=policy,
                on_conflict=effective_on_conflict,
                artifact_bridge_dir=artifact_bridge_dir,
            )
        result = cast(PatchToolOutput, await anyio.to_thread.run_sync(work))
        return result

    _patch_tool.__doc__ = _build_patch_tool_description()

    patch_tool = app.tool(name="exstruct_patch")
    patch_tool(_patch_tool)

    async def _make_tool(
        out_path: str,
        ops: list[dict[str, Any] | str] | None = None,
        sheet: str | None = None,
        on_conflict: OnConflictPolicy | None = None,
        auto_formula: bool = False,
        dry_run: bool = False,
        return_inverse_ops: bool = False,
        preflight_formula_check: bool = False,
        backend: Literal["auto", "com", "openpyxl"] = "auto",
        mirror_artifact: bool = False,
    ) -> MakeToolOutput:
        """Create a new Excel workbook and apply patch operations.

        Args:
            out_path: Output workbook path (.xlsx/.xlsm/.xls).
            ops: Optional patch operations. Accepts object list or JSON object strings.
            sheet: Optional default sheet name. Used when op.sheet is omitted
                for non-add_sheet ops. If both are set, op.sheet wins.
            on_conflict: Conflict policy when output file exists:
                'overwrite' (replace), 'skip' (do nothing), 'rename' (auto-rename).
                Defaults to server --on-conflict setting.
            auto_formula: When true, values starting with '=' in set_value ops
                are treated as formulas instead of being rejected.
            dry_run: When true, compute diff without saving changes.
            return_inverse_ops: When true, return inverse (undo) operations.
            preflight_formula_check: When true, scan formulas for errors
                like #REF!, #NAME?, #DIV/0! before saving.
            backend: Patch execution backend.
                - "auto" (default): prefer COM when available; otherwise openpyxl.
                - "com": force COM path (requires Excel COM).
                - "openpyxl": force openpyxl path (.xls is not supported).
            mirror_artifact: When true, mirror output workbook to
                --artifact-bridge-dir after successful make/patch.

        Returns:
            Patch-compatible result with output path, diff, and warnings.
        """
        normalized_ops = _coerce_patch_ops(ops or [])
        payload = MakeToolInput(
            out_path=out_path,
            ops=normalized_ops,
            sheet=sheet,
            on_conflict=on_conflict,
            auto_formula=auto_formula,
            dry_run=dry_run,
            return_inverse_ops=return_inverse_ops,
            preflight_formula_check=preflight_formula_check,
            backend=backend,
            mirror_artifact=mirror_artifact,
        )
        effective_on_conflict = on_conflict or default_on_conflict
        if artifact_bridge_dir is None:
            work = functools.partial(
                run_make_tool,
                payload,
                policy=policy,
                on_conflict=effective_on_conflict,
            )
        else:
            work = functools.partial(
                run_make_tool,
                payload,
                policy=policy,
                on_conflict=effective_on_conflict,
                artifact_bridge_dir=artifact_bridge_dir,
            )
        result = cast(MakeToolOutput, await anyio.to_thread.run_sync(work))
        return result

    make_tool = app.tool(name="exstruct_make")
    make_tool(_make_tool)


def _build_patch_tool_description() -> str:
    """Build exstruct_patch tool description with op mini schema."""
    base_description = """
Edit an Excel workbook by applying patch operations.

Supports cell value updates, formula updates, and adding new sheets.
Operations are applied atomically: all succeed or none are saved.

Args:
    xlsx_path: Path to the Excel workbook to edit.
    ops: Patch operations to apply in order. Preferred format is an
        object list (one object per operation). For compatibility with
        clients that cannot send object arrays, JSON object strings are
        also accepted and normalized before validation.
    sheet: Optional default sheet name. Used when op.sheet is omitted
        for non-add_sheet ops. If both are set, op.sheet wins.
    out_dir: Output directory. Defaults to same directory as input.
    out_name: Output filename. Defaults to '{stem}_patched{ext}'.
        If stem already ends with '_patched', the same name is reused.
    on_conflict: Conflict policy when output file exists:
        'overwrite' (replace), 'skip' (do nothing), 'rename' (auto-rename).
        Defaults to server --on-conflict setting.
    auto_formula: When true, values starting with '=' in set_value ops
        are treated as formulas instead of being rejected.
    dry_run: When true, compute diff without saving changes.
    return_inverse_ops: When true, return inverse (undo) operations.
    preflight_formula_check: When true, scan formulas for errors
        like #REF!, #NAME?, #DIV/0! before saving.
    backend: Patch execution backend.
        - "auto" (default): prefer COM when available; otherwise openpyxl.
          Uses openpyxl when dry_run/return_inverse_ops/preflight_formula_check
          is enabled.
        - "com": force COM path (requires Excel COM and disallows
          dry_run/return_inverse_ops/preflight_formula_check).
        - "openpyxl": force openpyxl path (.xls is not supported).
    mirror_artifact: When true, mirror output workbook to
        --artifact-bridge-dir after successful patch.

Returns:
    Patch result with output path, applied diffs, and any warnings.
"""
    return f"{base_description.strip()}\n\n{build_patch_tool_mini_schema()}"


def _register_op_schema_tools(app: FastMCP) -> None:
    """Register schema discovery tools."""

    async def _list_ops_tool() -> ListOpsToolOutput:
        """List all patch op names with short descriptions."""
        return run_list_ops_tool()

    async def _describe_op_tool(op: str) -> DescribeOpToolOutput:
        """Describe one patch op.

        Returns required/optional fields, constraints, example, and aliases.
        """
        payload = DescribeOpToolInput(op=op)
        return run_describe_op_tool(payload)

    list_ops_tool = app.tool(name="exstruct_list_ops")
    list_ops_tool(_list_ops_tool)
    describe_op_tool = app.tool(name="exstruct_describe_op")
    describe_op_tool(_describe_op_tool)


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


def _coerce_patch_ops(ops_data: list[dict[str, Any] | str]) -> list[dict[str, Any]]:
    """Normalize patch operations payload for MCP clients.

    Args:
        ops_data: Raw operations from MCP tool input.

    Returns:
        Patch operations as object list.

    Raises:
        ValueError: If a string op is not valid JSON object.
    """
    return _normalize_coerce_patch_ops(ops_data)


def _parse_patch_op_json(raw_op: str, index: int) -> dict[str, Any]:
    """Parse a JSON string patch operation into object form.

    Args:
        raw_op: Raw JSON string for one patch operation.
        index: Source index in the ops list.

    Returns:
        Parsed patch operation object.

    Raises:
        ValueError: If the string is not valid JSON object.
    """
    return _normalize_parse_patch_op_json(raw_op, index=index)


def _build_patch_op_error_message(index: int, reason: str) -> str:
    """Build a consistent validation message for invalid patch ops.

    Args:
        index: Source index in the ops list.
        reason: Validation failure reason.

    Returns:
        Human-readable error message.
    """
    return _normalize_build_patch_op_error_message(index, reason)
