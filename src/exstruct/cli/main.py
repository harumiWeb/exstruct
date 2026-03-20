"""Command-line interface for ExStruct extraction and editing."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

from exstruct import process_excel
from exstruct.cli.availability import get_com_availability
from exstruct.cli.edit import is_edit_subcommand, run_edit_cli
from exstruct.constraints import validate_libreoffice_process_request


def _ensure_utf8_stdout() -> None:
    """Reconfigure stdout to UTF-8 when supported.

    Windows consoles default to cp932 and can raise encoding errors when piping
    non-ASCII characters. Reconfiguring prevents failures without affecting
    environments that already default to UTF-8.
    """

    stdout = sys.stdout
    if not hasattr(stdout, "reconfigure"):
        return
    reconfigure = stdout.reconfigure
    try:
        reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        return


def _add_auto_page_breaks_argument(parser: argparse.ArgumentParser) -> None:
    """Add the auto page-break export option to the extraction CLI."""
    parser.add_argument(
        "--auto-page-breaks-dir",
        type=Path,
        help=(
            "Optional directory to write one file per auto page-break area "
            "(Excel COM only; not supported in libreoffice mode)."
        ),
    )


def build_parser() -> argparse.ArgumentParser:
    """Build the CLI argument parser."""
    parser = argparse.ArgumentParser(
        description="CLI for ExStruct extraction.",
        epilog=(
            "Editing commands:\n"
            "  exstruct patch --input book.xlsx --ops ops.json\n"
            "  exstruct make --output new.xlsx --ops ops.json\n"
            "  exstruct ops list\n"
            "  exstruct ops describe create_chart\n"
            "  exstruct validate --input book.xlsx"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="Excel file (.xlsx/.xlsm/.xls)")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output path. If omitted, writes to stdout.",
    )
    parser.add_argument(
        "-f",
        "--format",
        default="json",
        choices=["json", "yaml", "yml", "toon"],
        help="Export format",
    )
    parser.add_argument(
        "--image",
        action="store_true",
        help=(
            "Render per-sheet PNGs alongside structured output "
            "(Excel COM only; not supported in libreoffice mode)."
        ),
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help=(
            "Render PDF alongside structured output "
            "(Excel COM only; not supported in libreoffice mode)."
        ),
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=144,
        help="DPI for image rendering used with --image.",
    )
    parser.add_argument(
        "-m",
        "--mode",
        default="standard",
        choices=["light", "libreoffice", "standard", "verbose"],
        help=(
            "Extraction detail level. libreoffice is a best-effort rich extraction "
            "mode for .xlsx/.xlsm only and cannot be combined with PDF/PNG rendering "
            "or auto page-break export."
        ),
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON output (indent=2). Default is compact JSON.",
    )
    parser.add_argument(
        "--sheets-dir",
        type=Path,
        help="Optional directory to write one file per sheet (format follows --format).",
    )
    parser.add_argument(
        "--print-areas-dir",
        type=Path,
        help="Optional directory to write one file per print area (format follows --format).",
    )
    _add_auto_page_breaks_argument(parser)
    parser.add_argument(
        "--alpha-col",
        action="store_true",
        help="Output column keys as Excel-style ABC names (A, B, ..., Z, AA, ...) instead of 0-based indices.",
    )
    parser.add_argument(
        "--include-backend-metadata",
        action="store_true",
        help=(
            "Include shape/chart backend metadata fields "
            "(provenance, approximation_level, confidence)."
        ),
    )
    return parser


def _validate_auto_page_breaks_request(args: argparse.Namespace) -> None:
    """Validate runtime requirements for auto page-break export."""
    auto_page_breaks_dir = getattr(args, "auto_page_breaks_dir", None)
    if auto_page_breaks_dir is None:
        return

    message = (
        "--auto-page-breaks-dir requires --mode standard or --mode verbose "
        "with Excel COM."
    )
    if args.mode == "libreoffice":
        validate_libreoffice_process_request(
            args.input,
            mode=args.mode,
            include_auto_page_breaks=True,
            pdf=args.pdf,
            image=args.image,
        )
        return
    if args.mode == "light":
        raise RuntimeError(message)

    availability = get_com_availability()
    if availability.available:
        return

    reason = f" Reason: {availability.reason}" if availability.reason else ""
    raise RuntimeError(f"{message}{reason}")


def main(argv: list[str] | None = None) -> int:
    """Run the CLI entrypoint.

    Args:
        argv: Optional argument list for testing.

    Returns:
        Exit code (0 for success, 1 for failure).
    """
    _ensure_utf8_stdout()
    resolved_argv = list(sys.argv[1:] if argv is None else argv)
    if is_edit_subcommand(resolved_argv):
        return run_edit_cli(resolved_argv)

    parser = build_parser()
    args = parser.parse_args(resolved_argv)

    input_path: Path = args.input
    if not input_path.exists():
        print(f"File not found: {input_path}", flush=True)
        return 0

    try:
        _validate_auto_page_breaks_request(args)
        process_excel(
            file_path=input_path,
            output_path=args.output,
            out_fmt=args.format,
            image=args.image,
            pdf=args.pdf,
            dpi=args.dpi,
            mode=args.mode,
            pretty=args.pretty,
            sheets_dir=args.sheets_dir,
            print_areas_dir=args.print_areas_dir,
            auto_page_breaks_dir=getattr(args, "auto_page_breaks_dir", None),
            alpha_col=args.alpha_col,
            include_backend_metadata=args.include_backend_metadata,
        )
        return 0
    except Exception as e:
        print(f"Error: {e}", flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
