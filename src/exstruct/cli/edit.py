"""CLI subcommands for workbook editing operations."""

from __future__ import annotations

import argparse
from collections.abc import Callable
from functools import lru_cache
import json
from pathlib import Path
import sys
from typing import Any, cast, get_args

_EDIT_SUBCOMMANDS = frozenset({"patch", "make", "ops", "validate"})
_EXPLICIT_EDIT_TOKENS: dict[str, frozenset[str]] = {
    "patch": frozenset(
        {
            "--help",
            "-h",
            "--input",
            "--ops",
            "--sheet",
            "--on-conflict",
            "--backend",
            "--auto-formula",
            "--dry-run",
            "--return-inverse-ops",
            "--preflight-formula-check",
        }
    ),
    "make": frozenset(
        {
            "--help",
            "-h",
            "--ops",
            "--sheet",
            "--on-conflict",
            "--backend",
            "--auto-formula",
            "--dry-run",
            "--return-inverse-ops",
            "--preflight-formula-check",
        }
    ),
    "ops": frozenset({"--help", "-h", "list", "describe"}),
    "validate": frozenset({"--help", "-h", "--input"}),
}


def _load_patch_workbook_impl() -> Callable[[object], object]:
    from exstruct.edit.api import patch_workbook as patch_workbook_impl

    return cast(Callable[[object], object], patch_workbook_impl)


def _load_make_workbook_impl() -> Callable[[object], object]:
    from exstruct.edit.api import make_workbook as make_workbook_impl

    return cast(Callable[[object], object], make_workbook_impl)


def _load_resolve_top_level_sheet_for_payload_impl() -> Callable[[object], object]:
    from exstruct.edit.normalize import (
        resolve_top_level_sheet_for_payload as resolve_top_level_sheet_for_payload_impl,
    )

    return cast(Callable[[object], object], resolve_top_level_sheet_for_payload_impl)


def _load_validate_input_impl() -> Callable[[object], object]:
    from exstruct.mcp.validate_input import validate_input as validate_input_impl

    return cast(Callable[[object], object], validate_input_impl)


def _load_on_conflict_policy_literal() -> object:
    from exstruct.edit.types import OnConflictPolicy

    return OnConflictPolicy


def _load_patch_backend_literal() -> object:
    from exstruct.edit.types import PatchBackend

    return PatchBackend


def _load_patch_request_model() -> Callable[..., object]:
    from exstruct.edit.models import PatchRequest

    return cast(Callable[..., object], PatchRequest)


def _load_make_request_model() -> Callable[..., object]:
    from exstruct.edit.models import MakeRequest

    return cast(Callable[..., object], MakeRequest)


def _load_patch_op_model() -> Callable[..., object]:
    from exstruct.edit.models import PatchOp

    return cast(Callable[..., object], PatchOp)


def _load_list_patch_op_schemas() -> Callable[[], list[object]]:
    from exstruct.edit.op_schema import list_patch_op_schemas

    return cast(Callable[[], list[object]], list_patch_op_schemas)


def _load_get_patch_op_schema() -> Callable[[str], object | None]:
    from exstruct.edit.op_schema import get_patch_op_schema

    return cast(Callable[[str], object | None], get_patch_op_schema)


def _load_validate_input_request_model() -> Callable[..., object]:
    from exstruct.mcp.validate_input import ValidateInputRequest

    return cast(Callable[..., object], ValidateInputRequest)


def _is_validation_error(exc: Exception) -> bool:
    from pydantic import ValidationError

    return isinstance(exc, ValidationError)


def _is_cli_runtime_error(exc: Exception) -> bool:
    return isinstance(exc, OSError | RuntimeError | ValueError) or _is_validation_error(
        exc
    )


def _is_validate_cli_error(exc: Exception) -> bool:
    return isinstance(exc, OSError | ValueError) or _is_validation_error(exc)


def patch_workbook(request: object) -> object:
    """Compatibility wrapper that resolves patch execution lazily."""

    return _load_patch_workbook_impl()(request)


def make_workbook(request: object) -> object:
    """Compatibility wrapper that resolves make execution lazily."""

    return _load_make_workbook_impl()(request)


def resolve_top_level_sheet_for_payload(payload: object) -> object:
    """Compatibility wrapper that resolves payload normalization lazily."""

    return _load_resolve_top_level_sheet_for_payload_impl()(payload)


def validate_input(request: object) -> object:
    """Compatibility wrapper that resolves validate execution lazily."""

    return _load_validate_input_impl()(request)


def _literal_choices(literal_type: object) -> tuple[str, ...]:
    """Return argparse choices derived from one string Literal type."""

    choices: list[str] = []
    for choice in get_args(literal_type):
        if not isinstance(choice, str):
            raise TypeError("CLI choices must derive from string Literal values.")
        choices.append(choice)
    return tuple(choices)


@lru_cache(maxsize=1)
def _on_conflict_choices() -> tuple[str, ...]:
    return _literal_choices(_load_on_conflict_policy_literal())


@lru_cache(maxsize=1)
def _backend_choices() -> tuple[str, ...]:
    return _literal_choices(_load_patch_backend_literal())


def is_edit_subcommand(argv: list[str]) -> bool:
    """Return whether argv targets the editing CLI."""

    if not argv:
        return False
    command = argv[0]
    if command not in _EDIT_SUBCOMMANDS:
        return False
    if _targets_edit_cli_explicitly(command, argv[1:]):
        return True
    return not Path(command).exists()


def _targets_edit_cli_explicitly(command: str, remainder: list[str]) -> bool:
    """Return whether argv contains edit-only syntax for one subcommand."""

    tokens = _EXPLICIT_EDIT_TOKENS[command]
    return any(_matches_explicit_edit_token(token, tokens) for token in remainder)


def _matches_explicit_edit_token(token: str, tokens: frozenset[str]) -> bool:
    """Return whether one argv token clearly targets the edit CLI."""

    if token in tokens:
        return True
    if not token.startswith("--"):
        return False
    option_name, _, _ = token.partition("=")
    return option_name in tokens


def build_edit_parser() -> argparse.ArgumentParser:
    """Build the edit-subcommand CLI parser."""

    parser = argparse.ArgumentParser(description="CLI for ExStruct workbook editing.")
    subparsers = parser.add_subparsers(dest="command")

    patch_parser = subparsers.add_parser("patch", help="Edit an existing workbook.")
    _add_patch_like_arguments(
        patch_parser,
        require_input=True,
        require_output=False,
        require_ops=True,
    )
    patch_parser.set_defaults(handler=_run_patch_command)

    make_parser = subparsers.add_parser("make", help="Create and edit a workbook.")
    _add_patch_like_arguments(
        make_parser,
        require_input=False,
        require_output=True,
        require_ops=False,
    )
    make_parser.set_defaults(handler=_run_make_command)

    ops_parser = subparsers.add_parser(
        "ops", help="Inspect supported patch operation schemas."
    )
    ops_subparsers = ops_parser.add_subparsers(dest="ops_command")

    ops_list_parser = ops_subparsers.add_parser(
        "list", help="List supported patch operations."
    )
    ops_list_parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON output.",
    )
    ops_list_parser.set_defaults(handler=_run_ops_list_command)

    ops_describe_parser = ops_subparsers.add_parser(
        "describe", help="Describe one patch operation."
    )
    ops_describe_parser.add_argument("op", help="Patch operation name.")
    ops_describe_parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON output.",
    )
    ops_describe_parser.set_defaults(handler=_run_ops_describe_command)

    validate_parser = subparsers.add_parser(
        "validate", help="Validate that an input workbook is readable."
    )
    validate_parser.add_argument(
        "--input",
        type=Path,
        required=True,
        help="Workbook path (.xlsx/.xlsm/.xls).",
    )
    validate_parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON output.",
    )
    validate_parser.set_defaults(handler=_run_validate_command)

    return parser


def run_edit_cli(argv: list[str]) -> int:
    """Run the edit-subcommand CLI."""

    parser = build_edit_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code in (None, 0) else 1
    handler = getattr(args, "handler", None)
    if handler is None:
        parser.print_help()
        return 1
    return int(handler(args))


def _add_patch_like_arguments(
    parser: argparse.ArgumentParser,
    *,
    require_input: bool,
    require_output: bool,
    require_ops: bool,
) -> None:
    """Add shared request arguments for patch/make commands."""

    if require_input:
        parser.add_argument(
            "--input",
            type=Path,
            required=True,
            help="Input workbook path.",
        )
    parser.add_argument(
        "--output",
        type=Path,
        required=require_output,
        help=(
            "Output workbook path."
            if require_output
            else "Optional output workbook path."
        ),
    )
    parser.add_argument(
        "--ops",
        required=require_ops,
        help="Path to a JSON array of patch ops, or '-' to read from stdin.",
    )
    parser.add_argument("--sheet", help="Top-level sheet fallback for patch ops.")
    parser.add_argument(
        "--on-conflict",
        choices=_on_conflict_choices(),
        default="overwrite",
        help="Conflict policy for output workbook paths.",
    )
    parser.add_argument(
        "--backend",
        choices=_backend_choices(),
        default="auto",
        help="Patch backend selection policy.",
    )
    parser.add_argument(
        "--auto-formula",
        action="store_true",
        help="Treat '=...' values as formulas in set_value ops.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Compute patch diff without saving workbook changes.",
    )
    parser.add_argument(
        "--return-inverse-ops",
        action="store_true",
        help="Return inverse ops when supported.",
    )
    parser.add_argument(
        "--preflight-formula-check",
        action="store_true",
        help="Run formula-health validation before saving when supported.",
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON output.",
    )


def _run_patch_command(args: argparse.Namespace) -> int:
    """Execute the patch subcommand."""

    try:
        ops = _load_patch_ops(args.ops, sheet=args.sheet)
        request_kwargs: dict[str, Any] = {
            "xlsx_path": args.input,
            "ops": ops,
            "sheet": args.sheet,
            "on_conflict": args.on_conflict,
            "auto_formula": args.auto_formula,
            "dry_run": args.dry_run,
            "return_inverse_ops": args.return_inverse_ops,
            "preflight_formula_check": args.preflight_formula_check,
            "backend": args.backend,
        }
        if args.output is not None:
            request_kwargs["out_dir"] = args.output.parent
            request_kwargs["out_name"] = args.output.name
        request = _load_patch_request_model()(**request_kwargs)
        result = cast(Any, patch_workbook(request))
    except Exception as exc:
        if not _is_cli_runtime_error(exc):
            raise
        _print_error(exc)
        return 1

    _print_json_payload(result, pretty=args.pretty)
    return 0 if result.error is None else 1


def _run_make_command(args: argparse.Namespace) -> int:
    """Execute the make subcommand."""

    try:
        ops = _load_patch_ops(args.ops, sheet=args.sheet)
        request = _load_make_request_model()(
            out_path=args.output,
            ops=ops,
            sheet=args.sheet,
            on_conflict=args.on_conflict,
            auto_formula=args.auto_formula,
            dry_run=args.dry_run,
            return_inverse_ops=args.return_inverse_ops,
            preflight_formula_check=args.preflight_formula_check,
            backend=args.backend,
        )
        result = cast(Any, make_workbook(request))
    except Exception as exc:
        if not _is_cli_runtime_error(exc):
            raise
        _print_error(exc)
        return 1

    _print_json_payload(result, pretty=args.pretty)
    return 0 if result.error is None else 1


def _run_ops_list_command(args: argparse.Namespace) -> int:
    """Execute the ops list subcommand."""

    schemas = cast(Any, _load_list_patch_op_schemas()())
    payload = {
        "ops": [
            {"op": schema.op, "description": schema.description} for schema in schemas
        ]
    }
    _print_json_payload(payload, pretty=args.pretty)
    return 0


def _run_ops_describe_command(args: argparse.Namespace) -> int:
    """Execute the ops describe subcommand."""

    schema = cast(Any, _load_get_patch_op_schema()(args.op))
    if schema is None:
        _print_error(ValueError(f"Unknown patch operation: {args.op}"))
        return 1
    _print_json_payload(schema, pretty=args.pretty)
    return 0


def _run_validate_command(args: argparse.Namespace) -> int:
    """Execute the validate subcommand."""

    try:
        result = cast(
            Any,
            validate_input(_load_validate_input_request_model()(xlsx_path=args.input)),
        )
    except Exception as exc:
        if not _is_validate_cli_error(exc):
            raise
        _print_error(exc)
        return 1

    _print_json_payload(result, pretty=args.pretty)
    return 0 if result.is_readable else 1


def _load_patch_ops(source: str | None, *, sheet: str | None = None) -> list[Any]:
    """Load patch ops from a JSON file or stdin."""

    if source is None:
        return []
    payload = _load_json_value(source)
    if not isinstance(payload, list):
        raise ValueError("--ops must contain a JSON array.")
    resolved_payload = resolve_top_level_sheet_for_payload(
        {"ops": payload, "sheet": sheet}
    )
    if not isinstance(resolved_payload, dict):
        raise ValueError("Top-level sheet resolution must return a dict payload.")
    resolved_ops = resolved_payload.get("ops")
    if not isinstance(resolved_ops, list):
        raise ValueError("Resolved patch ops payload must contain an ops list.")
    patch_op_model = _load_patch_op_model()
    return [patch_op_model(**op_payload) for op_payload in resolved_ops]


def _load_json_value(source: str) -> object:
    """Load a JSON value from a file path or stdin marker."""

    if source == "-":
        raw = sys.stdin.read()
    else:
        raw = Path(source).read_text(encoding="utf-8")
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON in --ops: {exc.msg}") from exc


def _print_json_payload(payload: object, *, pretty: bool) -> None:
    """Serialize one JSON payload to stdout."""

    serializable: object
    model_dump = getattr(payload, "model_dump", None)
    serializable = model_dump(mode="json") if callable(model_dump) else payload
    print(
        json.dumps(
            serializable,
            ensure_ascii=False,
            indent=2 if pretty else None,
        ),
        flush=True,
    )


def _print_error(exc: Exception) -> None:
    """Print one CLI error to stderr."""

    print(f"Error: {exc}", file=sys.stderr, flush=True)


__all__ = ["build_edit_parser", "is_edit_subcommand", "run_edit_cli"]
