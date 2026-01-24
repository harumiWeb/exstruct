from __future__ import annotations

import re
import unicodedata
from typing import Any

from .exact_match import canonical, exact_match
from .normalization_rules import (
    ListObjectRule,
    NormalizationRules,
    RuleIndex,
    build_rule_index,
    normalize_label,
)


def _list_score(truth_list: list[Any], pred_list: Any) -> float:
    """Compute a partial match score for lists.

    Args:
        truth_list: Ground-truth list.
        pred_list: Predicted list.

    Returns:
        Fraction of truth elements present in prediction (order-insensitive).
    """
    if not isinstance(pred_list, list):
        return 0.0
    if not truth_list:
        return 0.0
    # Use exact match on elements; ignore order and duplicates.
    truth_set = {_normalize_element(v) for v in truth_list}
    pred_set = {_normalize_element(v) for v in pred_list}
    if not truth_set:
        return 0.0
    return len(truth_set & pred_set) / len(truth_set)


def _normalize_scalar(value: Any) -> str:
    """Normalize scalar values for set comparison."""
    if value is None:
        return "null"
    text = str(value)
    text = _strip_circled_numbers(text)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("※", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _strip_circled_numbers(text: str) -> str:
    """Remove circled-number characters (e.g., ①②) for robust matching."""
    return "".join(ch for ch in text if unicodedata.category(ch) != "No")


def _normalize_element(value: Any) -> str:
    """Normalize list elements for comparison."""
    if isinstance(value, (dict, list)):
        return canonical(value)
    return _normalize_scalar(value)


def _normalize_scalar_with_rules(value: Any, index: RuleIndex | None) -> str:
    """Normalize scalar values with optional normalization rules."""
    text = normalize_label(str(value))
    if index is None:
        return text
    return index.alias_map.get(text, text)


def _expand_pred_item(value: Any, index: RuleIndex) -> list[str]:
    """Expand a predicted list item using split rules and aliases."""
    text = _normalize_scalar_with_rules(value, index)
    if text in index.split_map:
        return index.split_map[text]
    return [text]


_SPLIT_PATTERN = re.compile(r"[、,，・/／]+")


def _coerce_list(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(v) for v in value if v is not None]
    if isinstance(value, str):
        return [value]
    return []


def _normalize_text_field(
    value: Any, index: RuleIndex, *, prefix_pattern: str | None
) -> str:
    text = str(value)
    if prefix_pattern:
        text = re.sub(prefix_pattern, "", text).strip()
    return _normalize_scalar_with_rules(text, index)


def _normalize_list_field(
    value: Any, index: RuleIndex, rule: ListObjectRule, field_name: str
) -> list[str]:
    items = _coerce_list(value)
    if len(items) == 1 and field_name in rule.list_fields_contains:
        items = [t for t in _SPLIT_PATTERN.split(items[0]) if t.strip()]
    return [_normalize_scalar_with_rules(v, index) for v in items if str(v).strip()]


def _object_matches(
    truth_obj: dict[str, Any],
    pred_obj: dict[str, Any],
    rule: ListObjectRule,
    index: RuleIndex,
) -> bool:
    for field in rule.string_fields:
        if field not in truth_obj or field not in pred_obj:
            return False
        t_val = _normalize_text_field(
            truth_obj[field], index, prefix_pattern=rule.strip_prefix.get(field)
        )
        p_val = _normalize_text_field(
            pred_obj[field], index, prefix_pattern=rule.strip_prefix.get(field)
        )
        if t_val != p_val:
            return False

    for field in rule.string_fields_contains:
        if field not in truth_obj or field not in pred_obj:
            return False
        t_val = _normalize_text_field(
            truth_obj[field], index, prefix_pattern=rule.strip_prefix.get(field)
        )
        p_val = _normalize_text_field(
            pred_obj[field], index, prefix_pattern=rule.strip_prefix.get(field)
        )
        if t_val not in p_val and p_val not in t_val:
            return False

    for field in rule.list_fields:
        if field not in truth_obj or field not in pred_obj:
            return False
        t_list = _normalize_list_field(truth_obj[field], index, rule, field)
        p_list = _normalize_list_field(pred_obj[field], index, rule, field)
        if set(t_list) != set(p_list):
            return False

    for field in rule.list_fields_contains:
        if field not in truth_obj or field not in pred_obj:
            return False
        t_list = _normalize_list_field(truth_obj[field], index, rule, field)
        p_list = _normalize_list_field(pred_obj[field], index, rule, field)
        if t_list and not p_list:
            return False
        combined = " ".join(p_list)
        for t_val in t_list:
            if t_val not in combined:
                return False

    return True


def _lcs_length_objects(
    a: list[dict[str, Any]],
    b: list[dict[str, Any]],
    *,
    rule: ListObjectRule,
    index: RuleIndex,
) -> int:
    if not a or not b:
        return 0
    dp = [0] * (len(b) + 1)
    for i in range(1, len(a) + 1):
        prev = 0
        for j in range(1, len(b) + 1):
            temp = dp[j]
            if _object_matches(a[i - 1], b[j - 1], rule, index):
                dp[j] = prev + 1
            else:
                dp[j] = max(dp[j], dp[j - 1])
            prev = temp
    return dp[-1]


def _list_score_objects_normalized(
    truth_list: list[Any],
    pred_list: Any,
    *,
    rule: ListObjectRule,
    index: RuleIndex,
    ordered: bool,
) -> float:
    if not isinstance(pred_list, list):
        return 0.0
    if not truth_list:
        return 0.0
    truth_objs = [t for t in truth_list if isinstance(t, dict)]
    pred_objs = [p for p in pred_list if isinstance(p, dict)]
    if not truth_objs:
        return 0.0
    if ordered:
        lcs_len = _lcs_length_objects(truth_objs, pred_objs, rule=rule, index=index)
        return lcs_len / len(truth_objs)
    matched = 0
    used: set[int] = set()
    for t in truth_objs:
        for i, p in enumerate(pred_objs):
            if i in used:
                continue
            if _object_matches(t, p, rule, index):
                matched += 1
                used.add(i)
                break
    return matched / len(truth_objs)


def _list_score_ordered(truth_list: list[Any], pred_list: Any) -> float:
    """Compute an order-aware partial match score for lists.

    Args:
        truth_list: Ground-truth list.
        pred_list: Predicted list.

    Returns:
        LCS-based fraction of truth elements matched in order.
    """
    if not isinstance(pred_list, list):
        return 0.0
    if not truth_list:
        return 0.0
    truth_norm = [_normalize_element(v) for v in truth_list]
    pred_norm = [_normalize_element(v) for v in pred_list]
    lcs_len = _lcs_length(truth_norm, pred_norm)
    return lcs_len / len(truth_norm)


def _list_score_normalized(
    truth_list: list[Any], pred_list: Any, index: RuleIndex
) -> float:
    """Compute a partial match score for lists with normalization rules."""
    if not isinstance(pred_list, list):
        return 0.0
    if not truth_list:
        return 0.0
    truth_norm = [_normalize_scalar_with_rules(v, index) for v in truth_list]
    pred_expanded: list[str] = []
    for v in pred_list:
        pred_expanded.extend(_expand_pred_item(v, index))
    pred_set = set(pred_expanded)
    matched = 0
    for t in truth_norm:
        if t in pred_set:
            matched += 1
            continue
        if t in index.composite_map:
            for parts in index.composite_map[t]:
                if all(p in pred_set for p in parts):
                    matched += 1
                    break
    return matched / len(truth_norm)


def _list_score_ordered_normalized(
    truth_list: list[Any], pred_list: Any, index: RuleIndex
) -> float:
    """Compute order-aware list score with normalization rules."""
    if not isinstance(pred_list, list):
        return 0.0
    if not truth_list:
        return 0.0
    truth_norm = [_normalize_scalar_with_rules(v, index) for v in truth_list]
    pred_expanded: list[str] = []
    for v in pred_list:
        pred_expanded.extend(_expand_pred_item(v, index))
    lcs_len = _lcs_length(truth_norm, pred_expanded)
    return lcs_len / len(truth_norm)


def _lcs_length(a: list[str], b: list[str]) -> int:
    """Compute the length of the longest common subsequence."""
    if not a or not b:
        return 0
    dp = [0] * (len(b) + 1)
    for i in range(1, len(a) + 1):
        prev = 0
        for j in range(1, len(b) + 1):
            temp = dp[j]
            if a[i - 1] == b[j - 1]:
                dp[j] = prev + 1
            else:
                dp[j] = max(dp[j], dp[j - 1])
            prev = temp
    return dp[-1]


def _dict_score(truth_dict: dict[str, Any], pred_dict: dict[str, Any]) -> float:
    """Compute a key-level score for nested dicts (order-insensitive lists)."""
    total = len(truth_dict)
    if total == 0:
        return 0.0
    score_sum = 0.0
    for key, truth_val in truth_dict.items():
        if key not in pred_dict:
            continue
        pred_val = pred_dict[key]
        score_sum += _value_score(truth_val, pred_val, ordered=False)
    return score_sum / total


def _dict_score_ordered(truth_dict: dict[str, Any], pred_dict: dict[str, Any]) -> float:
    """Compute a key-level score for nested dicts (order-aware lists)."""
    total = len(truth_dict)
    if total == 0:
        return 0.0
    score_sum = 0.0
    for key, truth_val in truth_dict.items():
        if key not in pred_dict:
            continue
        pred_val = pred_dict[key]
        score_sum += _value_score(truth_val, pred_val, ordered=True)
    return score_sum / total


def _dict_score_normalized(
    truth_dict: dict[str, Any],
    pred_dict: dict[str, Any],
    index: RuleIndex,
    list_object_rules: dict[str, ListObjectRule],
) -> float:
    """Compute a key-level score for nested dicts with normalization rules."""
    total = len(truth_dict)
    if total == 0:
        return 0.0
    score_sum = 0.0
    for key, truth_val in truth_dict.items():
        if key not in pred_dict:
            continue
        pred_val = pred_dict[key]
        rule = list_object_rules.get(key)
        if rule and isinstance(truth_val, list) and isinstance(pred_val, list):
            score_sum += _list_score_objects_normalized(
                truth_val, pred_val, rule=rule, index=index, ordered=False
            )
            continue
        score_sum += _value_score_normalized(
            truth_val,
            pred_val,
            index,
            ordered=False,
            list_object_rules=list_object_rules,
        )
    return score_sum / total


def _dict_score_ordered_normalized(
    truth_dict: dict[str, Any],
    pred_dict: dict[str, Any],
    index: RuleIndex,
    list_object_rules: dict[str, ListObjectRule],
) -> float:
    """Compute a key-level score with normalized, order-aware list scoring."""
    total = len(truth_dict)
    if total == 0:
        return 0.0
    score_sum = 0.0
    for key, truth_val in truth_dict.items():
        if key not in pred_dict:
            continue
        pred_val = pred_dict[key]
        rule = list_object_rules.get(key)
        if rule and isinstance(truth_val, list) and isinstance(pred_val, list):
            score_sum += _list_score_objects_normalized(
                truth_val, pred_val, rule=rule, index=index, ordered=True
            )
            continue
        score_sum += _value_score_normalized(
            truth_val,
            pred_val,
            index,
            ordered=True,
            list_object_rules=list_object_rules,
        )
    return score_sum / total


def _value_score(truth: Any, pred: Any, *, ordered: bool) -> float:
    """Score a value with optional list ordering."""
    if isinstance(truth, dict):
        if not isinstance(pred, dict):
            return 0.0
        return _dict_score_ordered(truth, pred) if ordered else _dict_score(truth, pred)
    if isinstance(truth, list):
        return _list_score_ordered(truth, pred) if ordered else _list_score(truth, pred)
    return 1.0 if exact_match(truth, pred) else 0.0


def _value_score_normalized(
    truth: Any,
    pred: Any,
    index: RuleIndex,
    *,
    ordered: bool,
    list_object_rules: dict[str, ListObjectRule],
) -> float:
    """Score a value using normalization rules."""
    if isinstance(truth, dict):
        if not isinstance(pred, dict):
            return 0.0
        return (
            _dict_score_ordered_normalized(truth, pred, index, list_object_rules)
            if ordered
            else _dict_score_normalized(truth, pred, index, list_object_rules)
        )
    if isinstance(truth, list):
        return (
            _list_score_ordered_normalized(truth, pred, index)
            if ordered
            else _list_score_normalized(truth, pred, index)
        )
    truth_norm = _normalize_scalar_with_rules(truth, index)
    pred_norm = _normalize_scalar_with_rules(pred, index)
    return 1.0 if truth_norm == pred_norm else 0.0


def key_score(truth: Any, pred: Any) -> float:
    """Compute a key-level score against the truth payload.

    Args:
        truth: Ground-truth JSON payload.
        pred: Predicted JSON payload.

    Returns:
        Score in [0, 1]. For dict payloads, this is the fraction of truth keys
        that exactly match in the prediction. For non-dict payloads, this is
        1.0 if exactly equal, else 0.0.
    """
    if isinstance(truth, dict):
        if not isinstance(pred, dict):
            return 0.0
        return _dict_score(truth, pred)
    if isinstance(truth, list):
        return _list_score(truth, pred)
    return 1.0 if exact_match(truth, pred) else 0.0


def key_score_ordered(truth: Any, pred: Any) -> float:
    """Compute a key-level score that respects list order."""
    if isinstance(truth, dict):
        if not isinstance(pred, dict):
            return 0.0
        return _dict_score_ordered(truth, pred)
    if isinstance(truth, list):
        return _list_score_ordered(truth, pred)
    return 1.0 if exact_match(truth, pred) else 0.0


def key_score_normalized(truth: Any, pred: Any, rules: NormalizationRules) -> float:
    """Compute a normalized score using optional rules."""
    index = build_rule_index(rules)
    list_object_rules = rules.list_object_rule_map()
    if isinstance(truth, dict):
        if not isinstance(pred, dict):
            return 0.0
        return _dict_score_normalized(truth, pred, index, list_object_rules)
    if isinstance(truth, list):
        return _list_score_normalized(truth, pred, index)
    truth_norm = _normalize_scalar_with_rules(truth, index)
    pred_norm = _normalize_scalar_with_rules(pred, index)
    return 1.0 if truth_norm == pred_norm else 0.0


def key_score_ordered_normalized(
    truth: Any, pred: Any, rules: NormalizationRules
) -> float:
    """Compute an order-aware normalized score using optional rules."""
    index = build_rule_index(rules)
    list_object_rules = rules.list_object_rule_map()
    if isinstance(truth, dict):
        if not isinstance(pred, dict):
            return 0.0
        return _dict_score_ordered_normalized(truth, pred, index, list_object_rules)
    if isinstance(truth, list):
        return _list_score_ordered_normalized(truth, pred, index)
    truth_norm = _normalize_scalar_with_rules(truth, index)
    pred_norm = _normalize_scalar_with_rules(pred, index)
    return 1.0 if truth_norm == pred_norm else 0.0
