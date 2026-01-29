from __future__ import annotations

import re
import unicodedata

_TABLE_SEPARATOR = re.compile(r"^[\s|:-]+$")
_WS_PATTERN = re.compile(r"\s+")
_NUMERIC_PATTERN = re.compile(r"[+-]?\d+(?:[.,]\d+)?")
_DOT_SEPARATORS = re.compile(r"[\u30fb\uff65\u00b7\u2022\u2219]")
_ZERO_WIDTH_PATTERN = re.compile(r"[\u200b\u200c\u200d\ufeff]")
_WEEKDAY_PAREN = re.compile(
    r"(?:\uFF08|\()"
    r"(?:\u6708|\u706B|\u6C34|\u6728|\u91D1|\u571F|\u65E5)"
    r"(?:\uFF09|\))"
)
_PAREN = re.compile(r"[\uFF08\uFF09()]")
_NON_ASCII_SPACE_PATTERN = re.compile(r"(?<=[^\x00-\x7F])\s+(?=[^\x00-\x7F])")


def markdown_coverage_score(truth_md: str, pred_md: str) -> float:
    """Compute coverage of truth Markdown lines in prediction.

    Args:
        truth_md: Canonical Markdown from truth JSON.
        pred_md: Markdown output to evaluate.

    Returns:
        Coverage score in [0, 1].
    """
    truth_lines = _normalized_lines(truth_md)
    pred_lines = _normalized_lines(pred_md)
    if not truth_lines:
        return 0.0
    matched = 0
    for t in truth_lines:
        if any(_match_line(t, p) for p in pred_lines):
            matched += 1
    return matched / len(truth_lines)


def markdown_precision_score(truth_md: str, pred_md: str) -> float:
    """Compute precision of prediction Markdown lines against truth.

    Args:
        truth_md: Canonical Markdown from truth JSON.
        pred_md: Markdown output to evaluate.

    Returns:
        Precision score in [0, 1].
    """
    truth_lines = _normalized_lines(truth_md)
    pred_lines = _normalized_lines(pred_md)
    if not pred_lines:
        return 0.0
    matched = 0
    for p in pred_lines:
        if any(_match_line(t, p) for t in truth_lines):
            matched += 1
    return matched / len(pred_lines)


def _normalized_lines(markdown: str) -> list[str]:
    """Normalize Markdown into comparable text lines."""
    lines: list[str] = []
    in_code_block = False
    for raw in markdown.splitlines():
        stripped = raw.strip()
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            continue
        norm = _normalize_line(raw)
        if not norm:
            continue
        if _TABLE_SEPARATOR.fullmatch(norm):
            continue
        lines.append(norm)
    return lines


def _normalize_line(line: str) -> str:
    """Normalize a single Markdown line for matching."""
    text = line.strip()
    if not text:
        return ""
    text = re.sub(r"^\s*#{1,6}\s*", "", text)
    text = re.sub(r"^\s*[-*+]\s+", "", text)
    text = text.replace("|", " ")
    text = text.replace("`", "")
    text = text.replace("*", "")
    text = text.replace(">", "")
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u3000", " ")
    text = _ZERO_WIDTH_PATTERN.sub("", text)
    text = _WEEKDAY_PAREN.sub("", text)
    text = _PAREN.sub("", text)
    text = _DOT_SEPARATORS.sub("", text)
    text = _WS_PATTERN.sub(" ", text)
    text = _NON_ASCII_SPACE_PATTERN.sub("", text)
    return text.strip()


def _match_line(truth_line: str, pred_line: str) -> bool:
    """Return True if lines match under loose Markdown rules."""
    if not truth_line or not pred_line:
        return False
    if _is_numeric_line(truth_line) or len(truth_line) == 1:
        return truth_line == pred_line
    return truth_line in pred_line or pred_line in truth_line


def _is_numeric_line(text: str) -> bool:
    """Return True if the text is numeric-only."""
    return _NUMERIC_PATTERN.fullmatch(text) is not None
