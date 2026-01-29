from __future__ import annotations

import re
import unicodedata
from typing import Any

_WS_PATTERN = re.compile(r"\s+")
_NUMERIC_PATTERN = re.compile(r"[+-]?\d+(?:[.,]\d+)?")


def _normalize_raw_text(text: str) -> str:
    """Normalize text for raw coverage/precision matching.

    Args:
        text: Input string.

    Returns:
        Normalized string with whitespace removed and width normalized.
    """
    normalized = unicodedata.normalize("NFKC", text)
    normalized = normalized.replace("窶ｻ", "")
    normalized = _WS_PATTERN.sub("", normalized)
    return normalized.strip()


def _is_numeric_token(text: str) -> bool:
    """Return True if the text looks like a numeric token.

    Args:
        text: Token to check.

    Returns:
        True if the token matches a numeric pattern.
    """
    return _NUMERIC_PATTERN.fullmatch(text) is not None


def _flatten_scalars(
    value: Any, *, depth: int = 0, parent_is_list: bool = False
) -> list[str]:
    """Flatten nested payloads into a list of scalar strings.

    Keys are included for nested dicts that are not record-like (dicts inside lists)
    to capture table headers or row labels without pulling schema field names.

    Args:
        value: Arbitrary JSON-like value.
        depth: Current nesting depth.
        parent_is_list: Whether the parent container is a list.

    Returns:
        List of stringified scalar values (and selected keys).
    """
    if value is None:
        return []
    if isinstance(value, dict):
        items: list[str] = []
        if depth > 0 and not parent_is_list:
            items.extend([str(k) for k in value.keys()])
        for v in value.values():
            items.extend(_flatten_scalars(v, depth=depth + 1, parent_is_list=False))
        return items
    if isinstance(value, list):
        items: list[str] = []
        for v in value:
            items.extend(_flatten_scalars(v, depth=depth + 1, parent_is_list=True))
        return items
    return [str(value)]


def _dedupe_normalized(values: list[str]) -> list[str]:
    """Normalize and de-duplicate text values, dropping empty tokens.

    Args:
        values: List of raw string values.

    Returns:
        De-duplicated list of normalized tokens.
    """
    seen: set[str] = set()
    normalized: list[str] = []
    for value in values:
        token = _normalize_raw_text(value)
        if not token:
            continue
        if token not in seen:
            seen.add(token)
            normalized.append(token)
    return normalized


def _raw_match_token(truth_token: str, pred_token: str) -> bool:
    """Return True if tokens match under loose raw-data matching rules.

    Args:
        truth_token: Normalized truth token.
        pred_token: Normalized prediction token.

    Returns:
        True if tokens are considered a match.
    """
    if not truth_token or not pred_token:
        return False
    if _is_numeric_token(truth_token) or len(truth_token) == 1:
        return truth_token == pred_token
    return truth_token in pred_token or pred_token in truth_token


def raw_coverage_score(truth: Any, pred: Any) -> float:
    """Compute loose coverage of truth tokens in predictions.

    Args:
        truth: Ground-truth JSON payload.
        pred: Predicted JSON payload.

    Returns:
        Coverage in [0, 1].
    """
    truth_tokens = _dedupe_normalized(_flatten_scalars(truth))
    pred_tokens = _dedupe_normalized(_flatten_scalars(pred))
    if not truth_tokens:
        return 0.0
    matched = 0
    for t in truth_tokens:
        if any(_raw_match_token(t, p) for p in pred_tokens):
            matched += 1
    return matched / len(truth_tokens)


def raw_precision_score(truth: Any, pred: Any) -> float:
    """Compute loose precision of prediction tokens against truth.

    Args:
        truth: Ground-truth JSON payload.
        pred: Predicted JSON payload.

    Returns:
        Precision in [0, 1].
    """
    truth_tokens = _dedupe_normalized(_flatten_scalars(truth))
    pred_tokens = _dedupe_normalized(_flatten_scalars(pred))
    if not pred_tokens:
        return 0.0
    matched = 0
    for p in pred_tokens:
        if any(_raw_match_token(t, p) for t in truth_tokens):
            matched += 1
    return matched / len(pred_tokens)
