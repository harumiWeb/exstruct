"""First-class workbook editing API for ExStruct."""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .api import make_workbook, patch_workbook
    from .chart_types import (
        CHART_TYPE_ALIASES,
        CHART_TYPE_TO_COM_ID,
        SUPPORTED_CHART_TYPES,
        SUPPORTED_CHART_TYPES_CSV,
        SUPPORTED_CHART_TYPES_SET,
        normalize_chart_type,
        resolve_chart_type_id,
    )
    from .errors import PatchOpError
    from .models import (
        AlignmentSnapshot,
        BorderSideSnapshot,
        BorderSnapshot,
        ColumnDimensionSnapshot,
        DesignSnapshot,
        FillSnapshot,
        FontSnapshot,
        FormulaIssue,
        MakeRequest,
        MergeStateSnapshot,
        OpenpyxlEngineResult,
        OpenpyxlWorksheetProtocol,
        PatchDiffItem,
        PatchErrorDetail,
        PatchOp,
        PatchRequest,
        PatchResult,
        PatchValue,
        RowDimensionSnapshot,
        XlwingsRangeProtocol,
    )
    from .normalize import (
        alias_to_canonical_with_conflict_check,
        build_missing_sheet_message,
        build_patch_op_error_message,
        coerce_patch_ops,
        normalize_draw_grid_border_range,
        normalize_patch_op_aliases,
        normalize_top_level_sheet,
        parse_patch_op_json,
        resolve_top_level_sheet_for_payload,
    )
    from .op_schema import (
        PatchOpSchema,
        build_patch_tool_mini_schema,
        get_patch_op_schema,
        list_patch_op_schemas,
        schema_with_sheet_resolution_rules,
    )
    from .specs import PATCH_OP_SPECS, PatchOpSpec, get_alias_map_for_op
    from .types import (
        FormulaIssueCode,
        FormulaIssueLevel,
        HorizontalAlignType,
        OnConflictPolicy,
        PatchBackend,
        PatchEngine,
        PatchOpType,
        PatchStatus,
        PatchValueKind,
        VerticalAlignType,
    )

__all__ = [
    "AlignmentSnapshot",
    "BorderSideSnapshot",
    "BorderSnapshot",
    "CHART_TYPE_ALIASES",
    "CHART_TYPE_TO_COM_ID",
    "ColumnDimensionSnapshot",
    "DesignSnapshot",
    "FillSnapshot",
    "FontSnapshot",
    "FormulaIssue",
    "FormulaIssueCode",
    "FormulaIssueLevel",
    "HorizontalAlignType",
    "MakeRequest",
    "MergeStateSnapshot",
    "OnConflictPolicy",
    "OpenpyxlEngineResult",
    "OpenpyxlWorksheetProtocol",
    "PATCH_OP_SPECS",
    "PatchBackend",
    "PatchDiffItem",
    "PatchEngine",
    "PatchErrorDetail",
    "PatchOp",
    "PatchOpError",
    "PatchOpSchema",
    "PatchOpSpec",
    "PatchOpType",
    "PatchRequest",
    "PatchResult",
    "PatchStatus",
    "PatchValue",
    "PatchValueKind",
    "RowDimensionSnapshot",
    "SUPPORTED_CHART_TYPES",
    "SUPPORTED_CHART_TYPES_CSV",
    "SUPPORTED_CHART_TYPES_SET",
    "VerticalAlignType",
    "XlwingsRangeProtocol",
    "alias_to_canonical_with_conflict_check",
    "build_missing_sheet_message",
    "build_patch_op_error_message",
    "build_patch_tool_mini_schema",
    "coerce_patch_ops",
    "get_alias_map_for_op",
    "get_patch_op_schema",
    "list_patch_op_schemas",
    "make_workbook",
    "normalize_chart_type",
    "normalize_draw_grid_border_range",
    "normalize_patch_op_aliases",
    "normalize_top_level_sheet",
    "parse_patch_op_json",
    "patch_workbook",
    "resolve_chart_type_id",
    "resolve_top_level_sheet_for_payload",
    "schema_with_sheet_resolution_rules",
]

LazyExportLoader = Callable[[], object]


def _load_api_attr(name: str) -> object:
    from . import api as api_module

    return getattr(api_module, name)


def _load_chart_type_attr(name: str) -> object:
    from . import chart_types as chart_types_module

    return getattr(chart_types_module, name)


def _load_error_attr(name: str) -> object:
    from . import errors as errors_module

    return getattr(errors_module, name)


def _load_model_attr(name: str) -> object:
    from . import models as models_module

    return getattr(models_module, name)


def _load_normalize_attr(name: str) -> object:
    from . import normalize as normalize_module

    return getattr(normalize_module, name)


def _load_op_schema_attr(name: str) -> object:
    from . import op_schema as op_schema_module

    return getattr(op_schema_module, name)


def _load_specs_attr(name: str) -> object:
    from . import specs as specs_module

    return getattr(specs_module, name)


def _load_type_attr(name: str) -> object:
    from . import types as types_module

    return getattr(types_module, name)


_LAZY_EXPORTS: dict[str, LazyExportLoader] = {
    "AlignmentSnapshot": lambda: _load_model_attr("AlignmentSnapshot"),
    "BorderSideSnapshot": lambda: _load_model_attr("BorderSideSnapshot"),
    "BorderSnapshot": lambda: _load_model_attr("BorderSnapshot"),
    "CHART_TYPE_ALIASES": lambda: _load_chart_type_attr("CHART_TYPE_ALIASES"),
    "CHART_TYPE_TO_COM_ID": lambda: _load_chart_type_attr("CHART_TYPE_TO_COM_ID"),
    "ColumnDimensionSnapshot": lambda: _load_model_attr("ColumnDimensionSnapshot"),
    "DesignSnapshot": lambda: _load_model_attr("DesignSnapshot"),
    "FillSnapshot": lambda: _load_model_attr("FillSnapshot"),
    "FontSnapshot": lambda: _load_model_attr("FontSnapshot"),
    "FormulaIssue": lambda: _load_model_attr("FormulaIssue"),
    "FormulaIssueCode": lambda: _load_type_attr("FormulaIssueCode"),
    "FormulaIssueLevel": lambda: _load_type_attr("FormulaIssueLevel"),
    "HorizontalAlignType": lambda: _load_type_attr("HorizontalAlignType"),
    "MakeRequest": lambda: _load_model_attr("MakeRequest"),
    "MergeStateSnapshot": lambda: _load_model_attr("MergeStateSnapshot"),
    "OnConflictPolicy": lambda: _load_type_attr("OnConflictPolicy"),
    "OpenpyxlEngineResult": lambda: _load_model_attr("OpenpyxlEngineResult"),
    "OpenpyxlWorksheetProtocol": lambda: _load_model_attr("OpenpyxlWorksheetProtocol"),
    "PATCH_OP_SPECS": lambda: _load_specs_attr("PATCH_OP_SPECS"),
    "PatchBackend": lambda: _load_type_attr("PatchBackend"),
    "PatchDiffItem": lambda: _load_model_attr("PatchDiffItem"),
    "PatchEngine": lambda: _load_type_attr("PatchEngine"),
    "PatchErrorDetail": lambda: _load_model_attr("PatchErrorDetail"),
    "PatchOp": lambda: _load_model_attr("PatchOp"),
    "PatchOpError": lambda: _load_error_attr("PatchOpError"),
    "PatchOpSchema": lambda: _load_op_schema_attr("PatchOpSchema"),
    "PatchOpSpec": lambda: _load_specs_attr("PatchOpSpec"),
    "PatchOpType": lambda: _load_type_attr("PatchOpType"),
    "PatchRequest": lambda: _load_model_attr("PatchRequest"),
    "PatchResult": lambda: _load_model_attr("PatchResult"),
    "PatchStatus": lambda: _load_type_attr("PatchStatus"),
    "PatchValue": lambda: _load_model_attr("PatchValue"),
    "PatchValueKind": lambda: _load_type_attr("PatchValueKind"),
    "RowDimensionSnapshot": lambda: _load_model_attr("RowDimensionSnapshot"),
    "SUPPORTED_CHART_TYPES": lambda: _load_chart_type_attr("SUPPORTED_CHART_TYPES"),
    "SUPPORTED_CHART_TYPES_CSV": lambda: _load_chart_type_attr(
        "SUPPORTED_CHART_TYPES_CSV"
    ),
    "SUPPORTED_CHART_TYPES_SET": lambda: _load_chart_type_attr(
        "SUPPORTED_CHART_TYPES_SET"
    ),
    "VerticalAlignType": lambda: _load_type_attr("VerticalAlignType"),
    "XlwingsRangeProtocol": lambda: _load_model_attr("XlwingsRangeProtocol"),
    "alias_to_canonical_with_conflict_check": lambda: _load_normalize_attr(
        "alias_to_canonical_with_conflict_check"
    ),
    "build_missing_sheet_message": lambda: _load_normalize_attr(
        "build_missing_sheet_message"
    ),
    "build_patch_op_error_message": lambda: _load_normalize_attr(
        "build_patch_op_error_message"
    ),
    "build_patch_tool_mini_schema": lambda: _load_op_schema_attr(
        "build_patch_tool_mini_schema"
    ),
    "coerce_patch_ops": lambda: _load_normalize_attr("coerce_patch_ops"),
    "get_alias_map_for_op": lambda: _load_specs_attr("get_alias_map_for_op"),
    "get_patch_op_schema": lambda: _load_op_schema_attr("get_patch_op_schema"),
    "list_patch_op_schemas": lambda: _load_op_schema_attr("list_patch_op_schemas"),
    "make_workbook": lambda: _load_api_attr("make_workbook"),
    "normalize_chart_type": lambda: _load_chart_type_attr("normalize_chart_type"),
    "normalize_draw_grid_border_range": lambda: _load_normalize_attr(
        "normalize_draw_grid_border_range"
    ),
    "normalize_patch_op_aliases": lambda: _load_normalize_attr(
        "normalize_patch_op_aliases"
    ),
    "normalize_top_level_sheet": lambda: _load_normalize_attr(
        "normalize_top_level_sheet"
    ),
    "parse_patch_op_json": lambda: _load_normalize_attr("parse_patch_op_json"),
    "patch_workbook": lambda: _load_api_attr("patch_workbook"),
    "resolve_chart_type_id": lambda: _load_chart_type_attr("resolve_chart_type_id"),
    "resolve_top_level_sheet_for_payload": lambda: _load_normalize_attr(
        "resolve_top_level_sheet_for_payload"
    ),
    "schema_with_sheet_resolution_rules": lambda: _load_op_schema_attr(
        "schema_with_sheet_resolution_rules"
    ),
}


def _resolve_lazy_export(name: str) -> object:
    value = _LAZY_EXPORTS[name]()
    globals()[name] = value
    return value


def __getattr__(name: str) -> object:
    if name not in _LAZY_EXPORTS:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    return _resolve_lazy_export(name)


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(__all__))
