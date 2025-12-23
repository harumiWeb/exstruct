from __future__ import annotations

import argparse
from pathlib import Path
import sys

from exstruct import process_excel
from exstruct.cli.availability import ComAvailability, get_com_availability


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


def _add_auto_page_breaks_argument(
    parser: argparse.ArgumentParser, availability: ComAvailability
) -> None:
    """Add auto page-break export option when COM is available."""
    if not availability.available:
        return
    parser.add_argument(
        "--auto-page-breaks-dir",
        type=Path,
        help="Optional directory to write one file per auto page-break area (COM only).",
    )


def build_parser(
    availability: ComAvailability | None = None,
) -> argparse.ArgumentParser:
    """Build the CLI argument parser.

    Args:
        availability: Optional COM availability for tests or overrides.

    Returns:
        Configured argument parser.
    """
    parser = argparse.ArgumentParser(
        description="Dev-only CLI stub for ExStruct extraction."
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
        help="(placeholder) Render PNG alongside JSON",
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="(placeholder) Render PDF alongside JSON",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=144,
        help="DPI for image rendering (placeholder)",
    )
    parser.add_argument(
        "-m",
        "--mode",
        default="standard",
        choices=["light", "standard", "verbose"],
        help="Extraction detail level",
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
    resolved_availability = (
        availability if availability is not None else get_com_availability()
    )
    _add_auto_page_breaks_argument(parser, resolved_availability)
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the CLI entrypoint.

    Args:
        argv: Optional argument list for testing.

    Returns:
        Exit code (0 for success, 1 for failure).
    """
    _ensure_utf8_stdout()
    parser = build_parser()
    args = parser.parse_args(argv)

    input_path: Path = args.input
    if not input_path.exists():
        print(f"File not found: {input_path}", flush=True)
        return 0

    try:
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
        )
        return 0
    except Exception as e:
        print(f"Error: {e}", flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
