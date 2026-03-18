"""Canonical openpyxl engine boundary for workbook editing."""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any, TypeVar, cast

from pydantic import BaseModel, ValidationError

from exstruct.edit import internal as _internal
from exstruct.edit.models import (
    FormulaIssue,
    OpenpyxlEngineResult,
    PatchDiffItem,
    PatchOp,
    PatchRequest,
)

TModel = TypeVar("TModel", bound=BaseModel)


def apply_openpyxl_engine(
    request: PatchRequest,
    input_path: Path,
    output_path: Path,
) -> OpenpyxlEngineResult:
    """Apply patch operations using the edit-owned openpyxl implementation."""
    diff, inverse_ops, formula_issues, op_warnings = _internal._apply_ops_openpyxl(
        cast(Any, request),
        input_path,
        output_path,
    )
    return OpenpyxlEngineResult(
        patch_diff=_coerce_model_list(diff, PatchDiffItem),
        inverse_ops=_coerce_model_list(inverse_ops, PatchOp),
        formula_issues=_coerce_model_list(formula_issues, FormulaIssue),
        op_warnings=list(op_warnings),
    )


def _coerce_model_list(
    items: Sequence[object], model_cls: type[TModel]
) -> list[TModel]:
    """Normalize model-like payloads into canonical Pydantic models."""
    coerced: list[TModel] = []
    for item in items:
        try:
            source: object
            if isinstance(item, model_cls):
                coerced.append(item)
                continue
            if isinstance(item, BaseModel):
                source = item.model_dump(mode="python")
            else:
                source = item
            coerced.append(model_cls.model_validate(source))
        except ValidationError:
            continue
    return coerced


__all__ = ["apply_openpyxl_engine"]
