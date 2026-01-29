from __future__ import annotations

import json
import re
import unicodedata
from typing import Any

from pydantic import BaseModel


class NormalizedPayload(BaseModel):
    """Normalized JSON payload for deterministic comparison."""

    value: Any


_WS_PATTERN = re.compile(r"\s+")
_ZERO_WIDTH_PATTERN = re.compile(r"[\u200b\u200c\u200d\ufeff]")
_NON_ASCII_SPACE_PATTERN = re.compile(r"(?<=[^\x00-\x7F])\s+(?=[^\x00-\x7F])")


def _normalize_text(value: str) -> str:
    """Normalize a string for comparison.

    Args:
        value: Raw string value.

    Returns:
        Normalized string.
    """
    text = value.replace("\r\n", "\n").replace("\r", "\n")
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u3000", " ")
    text = _ZERO_WIDTH_PATTERN.sub("", text)
    text = text.strip()
    text = _WS_PATTERN.sub(" ", text)
    text = _NON_ASCII_SPACE_PATTERN.sub("", text)
    return text.strip()


def _maybe_parse_number(value: str) -> int | float | str:
    """Parse a numeric string when possible.

    Args:
        value: String value.

    Returns:
        int/float when value is numeric, otherwise original string.
    """
    if re.fullmatch(r"-?\d+", value):
        return int(value)
    if re.fullmatch(r"-?\d+\.\d+", value):
        return float(value)
    return value


def _canonical_json(value: Any) -> str:
    """Return a canonical JSON string for sorting.

    Args:
        value: JSON-serializable value.

    Returns:
        Canonical JSON string.
    """
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _normalize_value(value: Any, *, unordered_paths: set[str], path: str) -> Any:
    """Normalize a JSON-like value recursively.

    Args:
        value: Input value.
        unordered_paths: Set of list paths to sort.
        path: Dot-separated path for the current value.

    Returns:
        Normalized value.
    """
    if isinstance(value, dict):
        normalized: dict[str, Any] = {}
        for key in sorted(value.keys()):
            child_path = f"{path}.{key}" if path else key
            normalized[key] = _normalize_value(
                value[key], unordered_paths=unordered_paths, path=child_path
            )
        return normalized
    if isinstance(value, list):
        normalized_items = [
            _normalize_value(item, unordered_paths=unordered_paths, path=path)
            for item in value
        ]
        if path in unordered_paths:
            normalized_items.sort(key=_canonical_json)
        return normalized_items
    if isinstance(value, str):
        return _maybe_parse_number(_normalize_text(value))
    return value


def normalize_payload(
    payload: Any, *, unordered_paths: list[str] | None = None
) -> NormalizedPayload:
    """Normalize a JSON payload with deterministic rules.

    Args:
        payload: Raw JSON object.
        unordered_paths: Dot paths for lists that should be treated as unordered.

    Returns:
        NormalizedPayload with normalized value.
    """
    path_set = set(unordered_paths or [])
    normalized = _normalize_value(payload, unordered_paths=path_set, path="")
    return NormalizedPayload(value=normalized)
