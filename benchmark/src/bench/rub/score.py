from __future__ import annotations

from collections import Counter
from typing import Any

from pydantic import BaseModel

from .normalize import normalize_payload


class RubScore(BaseModel):
    """Score result for a RUB task."""

    score: float
    ok: bool
    error: str | None = None


class RubPartialScore(BaseModel):
    """Partial match score for a RUB task."""

    precision: float
    recall: float
    f1: float


def _tokenize_scalar(value: Any) -> str | None:
    """Convert a scalar to a comparable token.

    Args:
        value: Scalar value.

    Returns:
        Token string or None for empty values.
    """
    if value is None:
        return None
    if isinstance(value, str):
        token = value.strip()
        return token or None
    return str(value)


def _flatten_tokens(value: Any) -> list[str]:
    """Flatten a JSON-like value into scalar tokens.

    Args:
        value: Normalized JSON value.

    Returns:
        List of scalar tokens.
    """
    tokens: list[str] = []
    if isinstance(value, dict):
        for v in value.values():
            tokens.extend(_flatten_tokens(v))
        return tokens
    if isinstance(value, list):
        for item in value:
            tokens.extend(_flatten_tokens(item))
        return tokens
    token = _tokenize_scalar(value)
    if token is not None:
        tokens.append(token)
    return tokens


def score_exact(
    truth: Any, pred: Any, *, unordered_paths: list[str] | None = None
) -> RubScore:
    """Compute exact-match score after normalization.

    Args:
        truth: Ground-truth JSON object.
        pred: Predicted JSON object.
        unordered_paths: Dot paths for unordered list comparison.

    Returns:
        RubScore with 1.0 for match, 0.0 otherwise.
    """
    truth_norm = normalize_payload(truth, unordered_paths=unordered_paths).value
    pred_norm = normalize_payload(pred, unordered_paths=unordered_paths).value
    ok = truth_norm == pred_norm
    return RubScore(score=1.0 if ok else 0.0, ok=ok)


def score_partial(
    truth: Any, pred: Any, *, unordered_paths: list[str] | None = None
) -> RubPartialScore:
    """Compute partial-match precision/recall/F1 after normalization.

    Args:
        truth: Ground-truth JSON object.
        pred: Predicted JSON object.
        unordered_paths: Dot paths for unordered list comparison.

    Returns:
        RubPartialScore with precision/recall/F1.
    """
    truth_norm = normalize_payload(truth, unordered_paths=unordered_paths).value
    pred_norm = normalize_payload(pred, unordered_paths=unordered_paths).value

    truth_tokens = _flatten_tokens(truth_norm)
    pred_tokens = _flatten_tokens(pred_norm)

    truth_counts = Counter(truth_tokens)
    pred_counts = Counter(pred_tokens)
    overlap = sum((truth_counts & pred_counts).values())

    truth_total = sum(truth_counts.values())
    pred_total = sum(pred_counts.values())

    if pred_total == 0:
        precision = 1.0 if truth_total == 0 else 0.0
    else:
        precision = overlap / pred_total
    if truth_total == 0:
        recall = 1.0 if pred_total == 0 else 0.0
    else:
        recall = overlap / truth_total

    if precision + recall == 0:
        f1 = 0.0
    else:
        f1 = 2 * precision * recall / (precision + recall)

    return RubPartialScore(precision=precision, recall=recall, f1=f1)
