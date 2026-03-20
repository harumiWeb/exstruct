"""First-class workbook editing API for ExStruct."""

from __future__ import annotations

from importlib import import_module
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

_LAZY_EXPORTS: dict[str, tuple[str, str]] = {
    "AlignmentSnapshot": (".models", "AlignmentSnapshot"),
    "BorderSideSnapshot": (".models", "BorderSideSnapshot"),
    "BorderSnapshot": (".models", "BorderSnapshot"),
    "CHART_TYPE_ALIASES": (".chart_types", "CHART_TYPE_ALIASES"),
    "CHART_TYPE_TO_COM_ID": (".chart_types", "CHART_TYPE_TO_COM_ID"),
    "ColumnDimensionSnapshot": (".models", "ColumnDimensionSnapshot"),
    "DesignSnapshot": (".models", "DesignSnapshot"),
    "FillSnapshot": (".models", "FillSnapshot"),
    "FontSnapshot": (".models", "FontSnapshot"),
    "FormulaIssue": (".models", "FormulaIssue"),
    "FormulaIssueCode": (".types", "FormulaIssueCode"),
    "FormulaIssueLevel": (".types", "FormulaIssueLevel"),
    "HorizontalAlignType": (".types", "HorizontalAlignType"),
    "MakeRequest": (".models", "MakeRequest"),
    "MergeStateSnapshot": (".models", "MergeStateSnapshot"),
    "OnConflictPolicy": (".types", "OnConflictPolicy"),
    "OpenpyxlEngineResult": (".models", "OpenpyxlEngineResult"),
    "OpenpyxlWorksheetProtocol": (".models", "OpenpyxlWorksheetProtocol"),
    "PATCH_OP_SPECS": (".specs", "PATCH_OP_SPECS"),
    "PatchBackend": (".types", "PatchBackend"),
    "PatchDiffItem": (".models", "PatchDiffItem"),
    "PatchEngine": (".types", "PatchEngine"),
    "PatchErrorDetail": (".models", "PatchErrorDetail"),
    "PatchOp": (".models", "PatchOp"),
    "PatchOpError": (".errors", "PatchOpError"),
    "PatchOpSchema": (".op_schema", "PatchOpSchema"),
    "PatchOpSpec": (".specs", "PatchOpSpec"),
    "PatchOpType": (".types", "PatchOpType"),
    "PatchRequest": (".models", "PatchRequest"),
    "PatchResult": (".models", "PatchResult"),
    "PatchStatus": (".types", "PatchStatus"),
    "PatchValue": (".models", "PatchValue"),
    "PatchValueKind": (".types", "PatchValueKind"),
    "RowDimensionSnapshot": (".models", "RowDimensionSnapshot"),
    "SUPPORTED_CHART_TYPES": (".chart_types", "SUPPORTED_CHART_TYPES"),
    "SUPPORTED_CHART_TYPES_CSV": (".chart_types", "SUPPORTED_CHART_TYPES_CSV"),
    "SUPPORTED_CHART_TYPES_SET": (".chart_types", "SUPPORTED_CHART_TYPES_SET"),
    "VerticalAlignType": (".types", "VerticalAlignType"),
    "XlwingsRangeProtocol": (".models", "XlwingsRangeProtocol"),
    "alias_to_canonical_with_conflict_check": (
        ".normalize",
        "alias_to_canonical_with_conflict_check",
    ),
    "build_missing_sheet_message": (".normalize", "build_missing_sheet_message"),
    "build_patch_op_error_message": (
        ".normalize",
        "build_patch_op_error_message",
    ),
    "build_patch_tool_mini_schema": (".op_schema", "build_patch_tool_mini_schema"),
    "coerce_patch_ops": (".normalize", "coerce_patch_ops"),
    "get_alias_map_for_op": (".specs", "get_alias_map_for_op"),
    "get_patch_op_schema": (".op_schema", "get_patch_op_schema"),
    "list_patch_op_schemas": (".op_schema", "list_patch_op_schemas"),
    "make_workbook": (".api", "make_workbook"),
    "normalize_chart_type": (".chart_types", "normalize_chart_type"),
    "normalize_draw_grid_border_range": (
        ".normalize",
        "normalize_draw_grid_border_range",
    ),
    "normalize_patch_op_aliases": (".normalize", "normalize_patch_op_aliases"),
    "normalize_top_level_sheet": (".normalize", "normalize_top_level_sheet"),
    "parse_patch_op_json": (".normalize", "parse_patch_op_json"),
    "patch_workbook": (".api", "patch_workbook"),
    "resolve_chart_type_id": (".chart_types", "resolve_chart_type_id"),
    "resolve_top_level_sheet_for_payload": (
        ".normalize",
        "resolve_top_level_sheet_for_payload",
    ),
    "schema_with_sheet_resolution_rules": (
        ".op_schema",
        "schema_with_sheet_resolution_rules",
    ),
}


def _resolve_lazy_export(name: str) -> object:
    module_name, attr_name = _LAZY_EXPORTS[name]
    value = getattr(import_module(module_name, __name__), attr_name)
    globals()[name] = value
    return value


def __getattr__(name: str) -> object:
    if name not in _LAZY_EXPORTS:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    return _resolve_lazy_export(name)


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(__all__))
