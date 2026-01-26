from __future__ import annotations

from typing import Any

from pydantic import BaseModel

from .normalize import normalize_payload


class RubScore(BaseModel):
    """Score result for a RUB task."""

    score: float
    ok: bool
    error: str | None = None


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
