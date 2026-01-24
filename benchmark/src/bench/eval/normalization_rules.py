from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

from pydantic import BaseModel, Field


class AliasRule(BaseModel):
    """Canonical label with its acceptable aliases."""

    canonical: str
    aliases: list[str] = Field(default_factory=list)


class SplitRule(BaseModel):
    """Split a combined label into multiple canonical labels."""

    trigger: str
    parts: list[str]


class CompositeRule(BaseModel):
    """Match a canonical label when all parts appear in prediction."""

    canonical: str
    parts: list[str]


class ListObjectRule(BaseModel):
    """Normalization rules for lists of objects."""

    list_key: str
    string_fields: list[str] = Field(default_factory=list)
    string_fields_contains: list[str] = Field(default_factory=list)
    list_fields: list[str] = Field(default_factory=list)
    list_fields_contains: list[str] = Field(default_factory=list)
    strip_prefix: dict[str, str] = Field(default_factory=dict)


class NormalizationRules(BaseModel):
    """Normalization rules for a single case."""

    alias_rules: list[AliasRule] = Field(default_factory=list)
    split_rules: list[SplitRule] = Field(default_factory=list)
    composite_rules: list[CompositeRule] = Field(default_factory=list)
    list_object_rules: list[ListObjectRule] = Field(default_factory=list)

    def list_object_rule_map(self) -> dict[str, ListObjectRule]:
        """Return list-object rules keyed by list key."""
        return {rule.list_key: rule for rule in self.list_object_rules}


class NormalizationRuleset(BaseModel):
    """Normalization rules keyed by case id."""

    cases: dict[str, NormalizationRules] = Field(default_factory=dict)

    def for_case(self, case_id: str) -> NormalizationRules:
        """Return rules for the given case id (or empty rules if missing)."""
        return self.cases.get(case_id, NormalizationRules())


class RuleIndex(BaseModel):
    """Prebuilt normalized lookup tables for scoring."""

    alias_map: dict[str, str] = Field(default_factory=dict)
    split_map: dict[str, list[str]] = Field(default_factory=dict)
    composite_map: dict[str, list[list[str]]] = Field(default_factory=dict)


def _strip_circled_numbers(text: str) -> str:
    """Remove circled-number characters for robust matching."""
    return "".join(ch for ch in text if unicodedata.category(ch) != "No")


def normalize_label(text: str) -> str:
    """Normalize labels for comparison."""
    text = _strip_circled_numbers(text)
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def build_rule_index(rules: NormalizationRules) -> RuleIndex:
    """Build normalized lookup tables from rules."""
    alias_map: dict[str, str] = {}
    for rule in rules.alias_rules:
        canonical = normalize_label(rule.canonical)
        alias_map[canonical] = canonical
        for alias in rule.aliases:
            alias_map[normalize_label(alias)] = canonical

    split_map: dict[str, list[str]] = {
        normalize_label(rule.trigger): [normalize_label(p) for p in rule.parts]
        for rule in rules.split_rules
    }

    composite_map: dict[str, list[list[str]]] = {}
    for rule in rules.composite_rules:
        canonical = normalize_label(rule.canonical)
        parts = [normalize_label(p) for p in rule.parts]
        composite_map.setdefault(canonical, []).append(parts)

    return RuleIndex(
        alias_map=alias_map,
        split_map=split_map,
        composite_map=composite_map,
    )


def load_ruleset(path: Path) -> NormalizationRuleset:
    """Load normalization ruleset from JSON file."""
    if not path.exists():
        return NormalizationRuleset()
    payload = json.loads(path.read_text(encoding="utf-8"))
    return NormalizationRuleset(**payload)
