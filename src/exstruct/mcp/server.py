from __future__ import annotations

import argparse
import importlib
import logging
from pathlib import Path
from types import ModuleType
from typing import TYPE_CHECKING, Any, Literal

from pydantic import BaseModel, Field

from exstruct import ExtractionMode

from .io import PathPolicy
from .tools import ExtractToolInput, ExtractToolOutput, run_extract_tool

if TYPE_CHECKING:  # pragma: no cover - typing only
    from mcp.server.fastmcp import FastMCP

logger = logging.getLogger(__name__)


class ServerConfig(BaseModel):
    """Configuration for the MCP server process."""

    root: Path = Field(..., description="Root directory for file access.")
    deny_globs: list[str] = Field(default_factory=list, description="Denied glob list.")
    log_level: str = Field(default="INFO", description="Logging level.")
    log_file: Path | None = Field(default=None, description="Optional log file path.")


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
    _import_mcp()
    policy = PathPolicy(root=config.root, deny_globs=config.deny_globs)
    logger.info("MCP root: %s", policy.normalize_root())
    app = _create_app(policy)
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
    args = parser.parse_args(argv)
    return ServerConfig(
        root=args.root,
        deny_globs=list(args.deny_glob),
        log_level=args.log_level,
        log_file=args.log_file,
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


def _create_app(policy: PathPolicy) -> FastMCP:
    """Create the MCP FastMCP application.

    Args:
        policy: Path policy for filesystem access.

    Returns:
        FastMCP application instance.
    """
    from mcp.server.fastmcp import FastMCP

    app = FastMCP("ExStruct MCP", json_response=True)
    _register_tools(app, policy)
    return app


def _register_tools(app: FastMCP, policy: PathPolicy) -> None:
    """Register MCP tools for the server.

    Args:
        app: FastMCP application instance.
        policy: Path policy for filesystem access.
    """

    def _extract_tool(
        xlsx_path: str,
        mode: ExtractionMode = "standard",
        format: Literal["json", "yaml", "yml", "toon"] = "json",  # noqa: A002
        out_dir: str | None = None,
        out_name: str | None = None,
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
            options=options or {},
        )
        return run_extract_tool(payload, policy=policy)

    tool = app.tool(name="exstruct.extract")
    tool(_extract_tool)
