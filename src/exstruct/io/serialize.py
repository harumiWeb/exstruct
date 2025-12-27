from __future__ import annotations

import importlib
import json
from types import ModuleType

from ..errors import MissingDependencyError, SerializationError
from ..models.types import JsonStructure

_FORMAT_HINTS: set[str] = {"json", "yaml", "toon"}


def _normalize_format_hint(fmt: str) -> str:
    """Normalize a format hint string.

    Args:
        fmt: Format string such as "json", "yaml", or "yml".

    Returns:
        Normalized format hint.
    """
    format_hint = fmt.lower()
    if format_hint == "yml":
        return "yaml"
    return format_hint


def _ensure_format_hint(
    fmt: str,
    *,
    allowed: set[str],
    error_type: type[Exception],
    error_message: str,
) -> str:
    """Validate and normalize a format hint.

    Args:
        fmt: Raw format string.
        allowed: Allowed format hints.
        error_type: Exception type to raise on error.
        error_message: Error message template with {fmt}.

    Returns:
        Normalized format hint.
    """
    format_hint = _normalize_format_hint(fmt)
    if format_hint not in allowed:
        raise error_type(error_message.format(fmt=fmt))
    return format_hint


def _serialize_payload_from_hint(
    payload: JsonStructure,
    format_hint: str,
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> str:
    """Serialize a payload using a normalized format hint.

    Args:
        payload: JSON-serializable payload.
        format_hint: Normalized format hint ("json", "yaml", "toon").
        pretty: Whether to pretty-print JSON.
        indent: Optional JSON indentation width.

    Returns:
        Serialized string for the requested format.
    """
    match format_hint:
        case "json":
            indent_val = 2 if pretty and indent is None else indent
            return json.dumps(payload, ensure_ascii=False, indent=indent_val)
        case "yaml":
            yaml = _require_yaml()
            return str(
                yaml.safe_dump(
                    payload,
                    allow_unicode=True,
                    sort_keys=False,
                    indent=2,
                )
            )
        case "toon":
            toon = _require_toon()
            return str(toon.encode(payload))
        case _:
            raise SerializationError(
                f"Unsupported export format '{format_hint}'. Allowed: json, yaml, yml, toon."
            )


def _require_yaml() -> ModuleType:
    """Ensure pyyaml is installed; otherwise raise with guidance."""
    try:
        module = importlib.import_module("yaml")
    except ImportError as e:
        raise MissingDependencyError(
            "YAML export requires pyyaml. Install it via `pip install pyyaml` or add the 'yaml' extra."
        ) from e
    return module


def _require_toon() -> ModuleType:
    """Ensure python-toon is installed; otherwise raise with guidance."""
    try:
        module = importlib.import_module("toon")
    except ImportError as e:
        raise MissingDependencyError(
            "TOON export requires python-toon. Install it via `pip install python-toon` or add the 'toon' extra."
        ) from e
    return module
