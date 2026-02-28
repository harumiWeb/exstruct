from __future__ import annotations

from collections.abc import Callable, Iterator
from pathlib import Path
import re
from typing import Protocol, runtime_checkable

from pydantic import BaseModel, Field, field_validator, model_validator

from ..extract_runner import OnConflictPolicy
from ..shared.a1 import (
    column_index_to_label as _shared_column_index_to_label,
    column_label_to_index as _shared_column_label_to_index,
    range_cell_count as _shared_range_cell_count,
    split_a1 as _shared_split_a1,
)
from .chart_types import SUPPORTED_CHART_TYPES_CSV, normalize_chart_type
from .types import (
    FormulaIssueCode,
    FormulaIssueLevel,
    HorizontalAlignType,
    PatchBackend,
    PatchEngine,
    PatchOpType,
    PatchStatus,
    PatchValueKind,
    VerticalAlignType,
)

_A1_PATTERN = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*$")
_A1_RANGE_PATTERN = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*:[A-Za-z]{1,3}[1-9][0-9]*$")
_SHEET_QUALIFIED_A1_RANGE_PATTERN = re.compile(
    r"^(?P<sheet>(?:'(?:(?:[^']|'')+)'|[^!]+)!)?"
    r"(?P<start>[A-Za-z]{1,3}[1-9][0-9]*):(?P<end>[A-Za-z]{1,3}[1-9][0-9]*)$"
)
_HEX_COLOR_PATTERN = re.compile(r"^#?(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$")
_COLUMN_LABEL_PATTERN = re.compile(r"^[A-Za-z]{1,3}$")
_MAX_STYLE_TARGET_CELLS = 10_000


class BorderSideSnapshot(BaseModel):
    """Serializable border side state for inverse restoration."""

    style: str | None = None
    color: str | None = None


class BorderSnapshot(BaseModel):
    """Serializable border state for one cell."""

    cell: str
    top: BorderSideSnapshot = Field(default_factory=BorderSideSnapshot)
    right: BorderSideSnapshot = Field(default_factory=BorderSideSnapshot)
    bottom: BorderSideSnapshot = Field(default_factory=BorderSideSnapshot)
    left: BorderSideSnapshot = Field(default_factory=BorderSideSnapshot)


class FontSnapshot(BaseModel):
    """Serializable font state for one cell."""

    cell: str
    bold: bool | None = None
    size: float | None = None
    color: str | None = None


class FillSnapshot(BaseModel):
    """Serializable fill state for one cell."""

    cell: str
    fill_type: str | None = None
    start_color: str | None = None
    end_color: str | None = None


class AlignmentSnapshot(BaseModel):
    """Serializable alignment state for one cell."""

    cell: str
    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool | None = None


class MergeStateSnapshot(BaseModel):
    """Serializable merged-range state for deterministic restoration."""

    scope: str
    ranges: list[str] = Field(default_factory=list)


class RowDimensionSnapshot(BaseModel):
    """Serializable row height state."""

    row: int
    height: float | None = None


class ColumnDimensionSnapshot(BaseModel):
    """Serializable column width state."""

    column: str
    width: float | None = None


class DesignSnapshot(BaseModel):
    """Serializable style/dimension snapshot for inverse restore."""

    borders: list[BorderSnapshot] = Field(default_factory=list)
    fonts: list[FontSnapshot] = Field(default_factory=list)
    fills: list[FillSnapshot] = Field(default_factory=list)
    alignments: list[AlignmentSnapshot] = Field(default_factory=list)
    merge_state: MergeStateSnapshot | None = None
    row_dimensions: list[RowDimensionSnapshot] = Field(default_factory=list)
    column_dimensions: list[ColumnDimensionSnapshot] = Field(default_factory=list)


@runtime_checkable
class OpenpyxlCellProtocol(Protocol):
    """Protocol for openpyxl cell access used by patch runner."""

    value: str | int | float | None
    data_type: str | None
    font: OpenpyxlFontProtocol
    fill: OpenpyxlFillProtocol
    border: OpenpyxlBorderProtocol
    alignment: OpenpyxlAlignmentProtocol


@runtime_checkable
class OpenpyxlColorProtocol(Protocol):
    """Protocol for openpyxl color access."""

    rgb: object | None


@runtime_checkable
class OpenpyxlSideProtocol(Protocol):
    """Protocol for openpyxl border side access."""

    style: str | None
    color: OpenpyxlColorProtocol | None


@runtime_checkable
class OpenpyxlBorderProtocol(Protocol):
    """Protocol for openpyxl border access."""

    top: OpenpyxlSideProtocol
    right: OpenpyxlSideProtocol
    bottom: OpenpyxlSideProtocol
    left: OpenpyxlSideProtocol


@runtime_checkable
class OpenpyxlFontProtocol(Protocol):
    """Protocol for openpyxl font access."""

    bold: bool | None
    size: float | None
    color: object | None


@runtime_checkable
class OpenpyxlFillProtocol(Protocol):
    """Protocol for openpyxl fill access."""

    fill_type: str | None
    start_color: OpenpyxlColorProtocol | None
    end_color: OpenpyxlColorProtocol | None


@runtime_checkable
class OpenpyxlAlignmentProtocol(Protocol):
    """Protocol for openpyxl alignment access."""

    horizontal: str | None
    vertical: str | None
    wrap_text: bool | None


@runtime_checkable
class OpenpyxlRowDimensionProtocol(Protocol):
    """Protocol for openpyxl row dimension access."""

    height: float | None


@runtime_checkable
class OpenpyxlColumnDimensionProtocol(Protocol):
    """Protocol for openpyxl column dimension access."""

    width: float | None


@runtime_checkable
class OpenpyxlRowDimensionsProtocol(Protocol):
    """Protocol for openpyxl row dimensions collection."""

    def __getitem__(self, key: int) -> OpenpyxlRowDimensionProtocol: ...


@runtime_checkable
class OpenpyxlColumnDimensionsProtocol(Protocol):
    """Protocol for openpyxl column dimensions collection."""

    def __getitem__(self, key: str) -> OpenpyxlColumnDimensionProtocol: ...


@runtime_checkable
class OpenpyxlWorksheetProtocol(Protocol):
    """Protocol for openpyxl worksheet access used by patch runner."""

    row_dimensions: OpenpyxlRowDimensionsProtocol
    column_dimensions: OpenpyxlColumnDimensionsProtocol

    def __getitem__(self, key: str) -> OpenpyxlCellProtocol: ...

    def merge_cells(self, range_string: str) -> None: ...

    def unmerge_cells(self, range_string: str) -> None: ...


@runtime_checkable
class OpenpyxlTablesProtocol(Protocol):
    """Protocol for openpyxl worksheet tables collection."""

    def items(self) -> Iterator[tuple[object, object]]: ...


@runtime_checkable
class OpenpyxlWorkbookProtocol(Protocol):
    """Protocol for openpyxl workbook access used by patch runner."""

    sheetnames: list[str]

    def __getitem__(self, key: str) -> OpenpyxlWorksheetProtocol: ...

    def create_sheet(self, title: str) -> OpenpyxlWorksheetProtocol: ...

    def save(self, filename: str | Path) -> None: ...

    def close(self) -> None: ...


@runtime_checkable
class XlwingsRangeProtocol(Protocol):
    """Protocol for xlwings range access used by patch runner."""

    value: object | None
    formula: str | None
    api: object


@runtime_checkable
class XlwingsSheetProtocol(Protocol):
    """Protocol for xlwings sheet access used by patch runner."""

    name: str
    api: object

    def range(self, cell: str) -> XlwingsRangeProtocol: ...


@runtime_checkable
class XlwingsSheetsProtocol(Protocol):
    """Protocol for xlwings sheets collection."""

    def __iter__(self) -> Iterator[XlwingsSheetProtocol]: ...

    def __len__(self) -> int: ...

    def __getitem__(self, index: int) -> XlwingsSheetProtocol: ...

    def add(
        self, name: str, after: XlwingsSheetProtocol | None = None
    ) -> XlwingsSheetProtocol: ...


@runtime_checkable
class XlwingsWorkbookProtocol(Protocol):
    """Protocol for xlwings workbook access used by patch runner."""

    sheets: XlwingsSheetsProtocol

    def save(self, filename: str) -> None: ...

    def close(self) -> None: ...


@runtime_checkable
class XlwingsFontApiProtocol(Protocol):
    """Protocol for xlwings COM font API."""

    Bold: bool
    Size: float
    Color: int


@runtime_checkable
class XlwingsInteriorApiProtocol(Protocol):
    """Protocol for xlwings COM interior API."""

    Color: int


@runtime_checkable
class XlwingsBorderApiProtocol(Protocol):
    """Protocol for xlwings COM border API."""

    LineStyle: int
    Color: int


@runtime_checkable
class XlwingsMergeAreaApiProtocol(Protocol):
    """Protocol for xlwings COM merged-area API."""

    def Address(self, row_absolute: bool, column_absolute: bool) -> str: ...  # noqa: N802


@runtime_checkable
class XlwingsRangeApiProtocol(Protocol):
    """Protocol for xlwings COM range API."""

    Font: XlwingsFontApiProtocol
    Interior: XlwingsInteriorApiProtocol
    MergeCells: bool
    MergeArea: XlwingsMergeAreaApiProtocol
    HorizontalAlignment: int
    VerticalAlignment: int
    WrapText: bool

    def Borders(self, edge: int) -> XlwingsBorderApiProtocol: ...  # noqa: N802

    def Merge(self) -> None: ...  # noqa: N802

    def UnMerge(self) -> None: ...  # noqa: N802


@runtime_checkable
class XlwingsRowApiProtocol(Protocol):
    """Protocol for xlwings COM row API."""

    RowHeight: float


@runtime_checkable
class XlwingsColumnApiProtocol(Protocol):
    """Protocol for xlwings COM column API."""

    ColumnWidth: float

    def AutoFit(self) -> None: ...  # noqa: N802


@runtime_checkable
class XlwingsSheetApiProtocol(Protocol):
    """Protocol for xlwings COM sheet API."""

    def Rows(self, index: int) -> XlwingsRowApiProtocol: ...  # noqa: N802

    def Columns(self, key: str) -> XlwingsColumnApiProtocol: ...  # noqa: N802


class PatchOp(BaseModel):
    """Single patch operation for an Excel workbook.

    Operation types and their required fields:

    - ``set_value``: Set a cell value. Requires ``sheet``, ``cell``, ``value``.
    - ``set_formula``: Set a cell formula. Requires ``sheet``, ``cell``, ``formula`` (must start with ``=``).
    - ``add_sheet``: Add a new worksheet. Requires ``sheet`` (new sheet name). No ``cell``/``value``/``formula``.
    - ``set_range_values``: Set values for a rectangular range. Requires ``sheet``, ``range`` (e.g. ``A1:C3``), ``values`` (2D list matching range shape).
    - ``fill_formula``: Fill a formula across a single row or column. Requires ``sheet``, ``range``, ``base_cell``, ``formula``.
    - ``set_value_if``: Conditionally set value. Requires ``sheet``, ``cell``, ``value``. ``expected`` is optional; ``null`` matches an empty cell. Skips if current value != expected.
    - ``set_formula_if``: Conditionally set formula. Requires ``sheet``, ``cell``, ``formula``. ``expected`` is optional; ``null`` matches an empty cell. Skips if current value != expected.
    - ``draw_grid_border``: Draw thin black borders on a target rectangle.
    - ``set_bold``: Set bold style for one cell or one range.
    - ``set_font_size``: Set font size for one cell or one range.
    - ``set_font_color``: Set font color for one cell or one range.
    - ``set_fill_color``: Set solid fill color for one cell or one range.
    - ``set_dimensions``: Set row height and/or column width.
    - ``auto_fit_columns``: Auto-fit column widths with optional bounds.
    - ``merge_cells``: Merge a rectangular range.
    - ``unmerge_cells``: Unmerge all merged ranges intersecting target range.
    - ``set_alignment``: Set horizontal/vertical alignment and/or wrap_text.
    - ``set_style``: Set multiple style attributes in one operation.
    - ``apply_table_style``: Create an Excel table and apply table style.
    - ``create_chart``: Create a new chart from source ranges (COM only).
    - ``restore_design_snapshot``: Restore style/dimension snapshot (internal inverse op).
    """

    op: PatchOpType = Field(
        description=(
            "Operation type: 'set_value', 'set_formula', 'add_sheet', "
            "'set_range_values', 'fill_formula', 'set_value_if', 'set_formula_if', "
            "'draw_grid_border', 'set_bold', 'set_font_size', 'set_font_color', "
            "'set_fill_color', "
            "'set_dimensions', "
            "'auto_fit_columns', "
            "'merge_cells', 'unmerge_cells', 'set_alignment', 'set_style', "
            "'apply_table_style', "
            "'create_chart', "
            "or 'restore_design_snapshot'."
        )
    )
    sheet: str = Field(
        description="Target sheet name. For add_sheet, this is the new sheet name."
    )
    cell: str | None = Field(
        default=None,
        description="Cell reference in A1 notation (e.g. 'B2'). Required for set_value, set_formula, set_value_if, set_formula_if.",
    )
    range: str | None = Field(
        default=None,
        description="Range reference in A1 notation (e.g. 'A1:C3'). Required for set_range_values and fill_formula.",
    )
    base_cell: str | None = Field(
        default=None,
        description="Base cell for formula translation in fill_formula (e.g. 'C2').",
    )
    expected: str | int | float | None = Field(
        default=None,
        description="Expected current value for conditional ops (set_value_if, set_formula_if). Operation is skipped if mismatch.",
    )
    value: str | int | float | None = Field(
        default=None,
        description="Value to set. Use null to clear a cell. For set_value and set_value_if.",
    )
    values: list[list[str | int | float | None]] | None = Field(
        default=None,
        description="2D list of values for set_range_values. Shape must match the range dimensions.",
    )
    formula: str | None = Field(
        default=None,
        description="Formula string starting with '=' (e.g. '=SUM(A1:A10)'). For set_formula, set_formula_if, fill_formula.",
    )
    row_count: int | None = Field(
        default=None,
        description="Row count for draw_grid_border.",
    )
    col_count: int | None = Field(
        default=None,
        description="Column count for draw_grid_border.",
    )
    bold: bool | None = Field(
        default=None,
        description="Bold flag for set_bold. Defaults to true.",
    )
    font_size: float | None = Field(
        default=None,
        description="Font size for set_font_size. Must be > 0.",
    )
    color: str | None = Field(
        default=None,
        description="Font color for set_font_color in RRGGBB/AARRGGBB (with optional '#').",
    )
    fill_color: str | None = Field(
        default=None,
        description="Fill color for set_fill_color in RRGGBB/AARRGGBB (with optional '#').",
    )
    rows: list[int] | None = Field(
        default=None,
        description="Row indexes for set_dimensions.",
    )
    columns: list[str | int] | None = Field(
        default=None,
        description="Column identifiers for set_dimensions. Accepts letters (A/AA) or positive indexes.",
    )
    row_height: float | None = Field(
        default=None,
        description="Target row height for set_dimensions.",
    )
    column_width: float | None = Field(
        default=None,
        description="Target column width for set_dimensions.",
    )
    min_width: float | None = Field(
        default=None,
        description="Optional minimum width bound for auto_fit_columns.",
    )
    max_width: float | None = Field(
        default=None,
        description="Optional maximum width bound for auto_fit_columns.",
    )
    horizontal_align: HorizontalAlignType | None = Field(
        default=None,
        description="Horizontal alignment for set_alignment/set_style.",
    )
    vertical_align: VerticalAlignType | None = Field(
        default=None,
        description="Vertical alignment for set_alignment/set_style.",
    )
    wrap_text: bool | None = Field(
        default=None,
        description="Wrap text flag for set_alignment/set_style.",
    )
    style: str | None = Field(
        default=None,
        description="Table style name for apply_table_style.",
    )
    table_name: str | None = Field(
        default=None,
        description="Optional table name for apply_table_style.",
    )
    design_snapshot: DesignSnapshot | None = Field(
        default=None,
        description="Design snapshot payload for restore_design_snapshot.",
    )
    chart_type: str | None = Field(
        default=None,
        description=(
            "Chart type for create_chart: line, column, bar, area, pie, "
            "doughnut, scatter, radar."
        ),
    )
    data_range: str | list[str] | None = Field(
        default=None,
        description=(
            "Data range in A1 notation for create_chart. "
            "Accepts a single range or a list of ranges."
        ),
    )
    category_range: str | None = Field(
        default=None,
        description="Optional category range in A1 notation for create_chart.",
    )
    anchor_cell: str | None = Field(
        default=None,
        description="Top-left anchor cell in A1 notation for chart placement.",
    )
    chart_name: str | None = Field(
        default=None,
        description="Optional chart object name for create_chart.",
    )
    width: float | None = Field(
        default=None,
        description="Optional chart width (points) for create_chart.",
    )
    height: float | None = Field(
        default=None,
        description="Optional chart height (points) for create_chart.",
    )
    titles_from_data: bool | None = Field(
        default=None,
        description="Whether to infer titles from source data for create_chart.",
    )
    series_from_rows: bool | None = Field(
        default=None,
        description="Whether chart series are oriented by rows for create_chart.",
    )
    chart_title: str | None = Field(
        default=None,
        description="Optional chart title text for create_chart.",
    )
    x_axis_title: str | None = Field(
        default=None,
        description="Optional X-axis title text for create_chart.",
    )
    y_axis_title: str | None = Field(
        default=None,
        description="Optional Y-axis title text for create_chart.",
    )

    @field_validator("sheet")
    @classmethod
    def _validate_sheet(cls, value: str) -> str:
        if not value.strip():
            raise ValueError("sheet must not be empty.")
        return value

    @field_validator("cell")
    @classmethod
    def _validate_cell(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not _A1_PATTERN.match(candidate):
            raise ValueError(f"Invalid cell reference: {value}")
        return candidate.upper()

    @field_validator("base_cell")
    @classmethod
    def _validate_base_cell(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not _A1_PATTERN.match(candidate):
            raise ValueError(f"Invalid base_cell reference: {value}")
        return candidate.upper()

    @field_validator("range")
    @classmethod
    def _validate_range(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not _A1_RANGE_PATTERN.match(candidate):
            raise ValueError(f"Invalid range reference: {value}")
        start, end = candidate.split(":", maxsplit=1)
        return f"{start.upper()}:{end.upper()}"

    @field_validator("data_range")
    @classmethod
    def _validate_data_range(
        cls, value: str | list[str] | None
    ) -> str | list[str] | None:
        if value is None:
            return None
        if isinstance(value, str):
            return _normalize_chart_range_reference(value)
        if not value:
            raise ValueError("data_range list must not be empty.")
        normalized: list[str] = []
        for item in value:
            normalized.append(_normalize_chart_range_reference(item))
        return normalized

    @field_validator("category_range")
    @classmethod
    def _validate_category_range(cls, value: str | None) -> str | None:
        if value is None:
            return None
        return _normalize_chart_range_reference(value)

    @field_validator("anchor_cell")
    @classmethod
    def _validate_anchor_cell(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not _A1_PATTERN.match(candidate):
            raise ValueError(f"Invalid anchor_cell reference: {value}")
        return candidate.upper()

    @field_validator("chart_type")
    @classmethod
    def _validate_chart_type(cls, value: str | None) -> str | None:
        if value is None:
            return None
        normalized = normalize_chart_type(value)
        if normalized is None:
            raise ValueError(f"chart_type must be one of: {SUPPORTED_CHART_TYPES_CSV}.")
        return normalized

    @field_validator("fill_color")
    @classmethod
    def _validate_fill_color(cls, value: str | None) -> str | None:
        if value is None:
            return None
        return _normalize_hex_input(value, field_name="fill_color")

    @field_validator("color")
    @classmethod
    def _validate_color(cls, value: str | None) -> str | None:
        if value is None:
            return None
        return _normalize_hex_input(value, field_name="color")

    @field_validator("rows")
    @classmethod
    def _validate_rows(cls, value: list[int] | None) -> list[int] | None:
        if value is None:
            return None
        if not value:
            raise ValueError("rows must not be empty.")
        normalized: list[int] = []
        for row in value:
            if row < 1:
                raise ValueError("rows must contain positive integers.")
            normalized.append(row)
        return normalized

    @field_validator("columns")
    @classmethod
    def _validate_columns(cls, value: list[str | int] | None) -> list[str | int] | None:
        if value is None:
            return None
        if not value:
            raise ValueError("columns must not be empty.")
        normalized: list[str | int] = []
        for column in value:
            normalized.append(_normalize_column_identifier(column))
        return normalized

    @field_validator(
        "style",
        "table_name",
        "chart_name",
        "chart_title",
        "x_axis_title",
        "y_axis_title",
    )
    @classmethod
    def _validate_non_empty_optional_text(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip()
        if not candidate:
            raise ValueError(
                "style/table_name/chart_name/chart_title/x_axis_title/y_axis_title "
                "must not be empty when provided."
            )
        return candidate

    @field_validator("min_width", "max_width", "width", "height")
    @classmethod
    def _validate_optional_positive_width(cls, value: float | None) -> float | None:
        if value is None:
            return None
        if value <= 0:
            raise ValueError("min_width/max_width/width/height must be > 0.")
        return value

    @model_validator(mode="after")
    def _validate_op(self) -> PatchOp:
        validator = _validator_for_op(self.op)
        if validator is None:
            return self
        if self.op in _CELL_REQUIRED_OPS:
            _validate_cell_required(self)
        validator(self)
        return self


_CELL_REQUIRED_OPS: set[PatchOpType] = {
    "set_value",
    "set_formula",
    "set_value_if",
    "set_formula_if",
}


def _validator_for_op(op_type: PatchOpType) -> Callable[[PatchOp], None] | None:
    """Return per-op validator function."""
    validators: dict[PatchOpType, Callable[[PatchOp], None]] = {
        "add_sheet": _validate_add_sheet,
        "set_value": _validate_set_value,
        "set_formula": _validate_set_formula,
        "set_range_values": _validate_set_range_values,
        "fill_formula": _validate_fill_formula,
        "set_value_if": _validate_set_value_if,
        "set_formula_if": _validate_set_formula_if,
        "draw_grid_border": _validate_draw_grid_border,
        "set_bold": _validate_set_bold,
        "set_font_size": _validate_set_font_size,
        "set_font_color": _validate_set_font_color,
        "set_fill_color": _validate_set_fill_color,
        "set_dimensions": _validate_set_dimensions,
        "auto_fit_columns": _validate_auto_fit_columns,
        "merge_cells": _validate_merge_cells,
        "unmerge_cells": _validate_unmerge_cells,
        "set_alignment": _validate_set_alignment,
        "set_style": _validate_set_style,
        "apply_table_style": _validate_apply_table_style,
        "create_chart": _validate_create_chart,
        "restore_design_snapshot": _validate_restore_design_snapshot,
    }
    return validators.get(op_type)


def _validate_add_sheet(op: PatchOp) -> None:
    """Validate add_sheet operation."""
    _validate_no_design_fields(op, op_name="add_sheet")
    if op.cell is not None:
        raise ValueError("add_sheet does not accept cell.")
    if op.range is not None:
        raise ValueError("add_sheet does not accept range.")
    if op.base_cell is not None:
        raise ValueError("add_sheet does not accept base_cell.")
    if op.expected is not None:
        raise ValueError("add_sheet does not accept expected.")
    if op.value is not None:
        raise ValueError("add_sheet does not accept value.")
    if op.values is not None:
        raise ValueError("add_sheet does not accept values.")
    if op.formula is not None:
        raise ValueError("add_sheet does not accept formula.")


def _validate_cell_required(op: PatchOp) -> None:
    """Validate that the operation has a cell value."""
    if op.cell is None:
        raise ValueError(f"{op.op} requires cell.")


def _validate_set_value(op: PatchOp) -> None:
    """Validate set_value operation."""
    _validate_no_design_fields(op, op_name="set_value")
    if op.range is not None:
        raise ValueError("set_value does not accept range.")
    if op.base_cell is not None:
        raise ValueError("set_value does not accept base_cell.")
    if op.expected is not None:
        raise ValueError("set_value does not accept expected.")
    if op.values is not None:
        raise ValueError("set_value does not accept values.")
    if op.formula is not None:
        raise ValueError("set_value does not accept formula.")


def _validate_set_formula(op: PatchOp) -> None:
    """Validate set_formula operation."""
    _validate_no_design_fields(op, op_name="set_formula")
    if op.range is not None:
        raise ValueError("set_formula does not accept range.")
    if op.base_cell is not None:
        raise ValueError("set_formula does not accept base_cell.")
    if op.expected is not None:
        raise ValueError("set_formula does not accept expected.")
    if op.values is not None:
        raise ValueError("set_formula does not accept values.")
    if op.value is not None:
        raise ValueError("set_formula does not accept value.")
    if op.formula is None:
        raise ValueError("set_formula requires formula.")
    if not op.formula.startswith("="):
        raise ValueError("set_formula requires formula starting with '='.")


def _validate_set_range_values(op: PatchOp) -> None:
    """Validate set_range_values operation."""
    _validate_no_design_fields(op, op_name="set_range_values")
    if op.cell is not None:
        raise ValueError("set_range_values does not accept cell.")
    if op.base_cell is not None:
        raise ValueError("set_range_values does not accept base_cell.")
    if op.expected is not None:
        raise ValueError("set_range_values does not accept expected.")
    if op.formula is not None:
        raise ValueError("set_range_values does not accept formula.")
    if op.range is None:
        raise ValueError("set_range_values requires range.")
    if op.values is None:
        raise ValueError("set_range_values requires values.")
    if not op.values:
        raise ValueError("set_range_values requires non-empty values.")
    if not all(op.values):
        raise ValueError("set_range_values values rows must not be empty.")
    expected_width = len(op.values[0])
    if any(len(row) != expected_width for row in op.values):
        raise ValueError("set_range_values requires rectangular values.")


def _validate_fill_formula(op: PatchOp) -> None:
    """Validate fill_formula operation."""
    _validate_no_design_fields(op, op_name="fill_formula")
    if op.cell is not None:
        raise ValueError("fill_formula does not accept cell.")
    if op.expected is not None:
        raise ValueError("fill_formula does not accept expected.")
    if op.value is not None:
        raise ValueError("fill_formula does not accept value.")
    if op.values is not None:
        raise ValueError("fill_formula does not accept values.")
    if op.range is None:
        raise ValueError("fill_formula requires range.")
    if op.base_cell is None:
        raise ValueError("fill_formula requires base_cell.")
    if op.formula is None:
        raise ValueError("fill_formula requires formula.")
    if not op.formula.startswith("="):
        raise ValueError("fill_formula requires formula starting with '='.")


def _validate_set_value_if(op: PatchOp) -> None:
    """Validate set_value_if operation."""
    _validate_no_design_fields(op, op_name="set_value_if")
    if op.formula is not None:
        raise ValueError("set_value_if does not accept formula.")
    if op.range is not None:
        raise ValueError("set_value_if does not accept range.")
    if op.values is not None:
        raise ValueError("set_value_if does not accept values.")
    if op.base_cell is not None:
        raise ValueError("set_value_if does not accept base_cell.")


def _validate_set_formula_if(op: PatchOp) -> None:
    """Validate set_formula_if operation."""
    _validate_no_design_fields(op, op_name="set_formula_if")
    if op.value is not None:
        raise ValueError("set_formula_if does not accept value.")
    if op.range is not None:
        raise ValueError("set_formula_if does not accept range.")
    if op.values is not None:
        raise ValueError("set_formula_if does not accept values.")
    if op.base_cell is not None:
        raise ValueError("set_formula_if does not accept base_cell.")
    if op.formula is None:
        raise ValueError("set_formula_if requires formula.")
    if not op.formula.startswith("="):
        raise ValueError("set_formula_if requires formula starting with '='.")


def _validate_draw_grid_border(op: PatchOp) -> None:
    """Validate draw_grid_border operation."""
    _validate_no_legacy_edit_fields(op, op_name="draw_grid_border")
    if op.cell is not None or op.range is not None:
        raise ValueError("draw_grid_border does not accept cell or range.")
    if op.bold is not None or op.color is not None or op.fill_color is not None:
        raise ValueError("draw_grid_border does not accept bold, color, or fill_color.")
    if op.font_size is not None:
        raise ValueError("draw_grid_border does not accept font_size.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("draw_grid_border does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("draw_grid_border does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("draw_grid_border does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="draw_grid_border")
    if op.base_cell is None:
        raise ValueError("draw_grid_border requires base_cell.")
    if op.row_count is None or op.col_count is None:
        raise ValueError("draw_grid_border requires row_count and col_count.")
    if op.row_count < 1 or op.col_count < 1:
        raise ValueError("draw_grid_border requires row_count >= 1 and col_count >= 1.")
    if op.row_count * op.col_count > _MAX_STYLE_TARGET_CELLS:
        raise ValueError(
            f"draw_grid_border target exceeds max cells: {_MAX_STYLE_TARGET_CELLS}."
        )


def _validate_set_bold(op: PatchOp) -> None:
    """Validate set_bold operation."""
    _validate_no_legacy_edit_fields(op, op_name="set_bold")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("set_bold does not accept row_count or col_count.")
    if op.color is not None or op.fill_color is not None:
        raise ValueError("set_bold does not accept color or fill_color.")
    if op.font_size is not None:
        raise ValueError("set_bold does not accept font_size.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("set_bold does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("set_bold does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("set_bold does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="set_bold")
    _validate_exactly_one_cell_or_range(op, op_name="set_bold")
    if op.bold is None:
        op.bold = True
    _validate_style_target_size(op, op_name="set_bold")


def _validate_set_font_size(op: PatchOp) -> None:
    """Validate set_font_size operation."""
    _validate_no_legacy_edit_fields(op, op_name="set_font_size")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("set_font_size does not accept row_count or col_count.")
    if op.bold is not None or op.color is not None or op.fill_color is not None:
        raise ValueError("set_font_size does not accept bold, color, or fill_color.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("set_font_size does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("set_font_size does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("set_font_size does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="set_font_size")
    _validate_exactly_one_cell_or_range(op, op_name="set_font_size")
    if op.font_size is None:
        raise ValueError("set_font_size requires font_size.")
    if op.font_size <= 0:
        raise ValueError("set_font_size font_size must be > 0.")
    _validate_style_target_size(op, op_name="set_font_size")


def _validate_set_font_color(op: PatchOp) -> None:
    """Validate set_font_color operation."""
    _validate_no_legacy_edit_fields(op, op_name="set_font_color")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("set_font_color does not accept row_count or col_count.")
    if op.bold is not None:
        raise ValueError("set_font_color does not accept bold.")
    if op.font_size is not None:
        raise ValueError("set_font_color does not accept font_size.")
    if op.fill_color is not None:
        raise ValueError("set_font_color does not accept fill_color.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("set_font_color does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("set_font_color does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("set_font_color does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="set_font_color")
    _validate_exactly_one_cell_or_range(op, op_name="set_font_color")
    if op.color is None:
        raise ValueError("set_font_color requires color.")
    _validate_style_target_size(op, op_name="set_font_color")


def _validate_set_fill_color(op: PatchOp) -> None:
    """Validate set_fill_color operation."""
    _validate_no_legacy_edit_fields(op, op_name="set_fill_color")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("set_fill_color does not accept row_count or col_count.")
    if op.bold is not None:
        raise ValueError("set_fill_color does not accept bold.")
    if op.color is not None:
        raise ValueError("set_fill_color does not accept color.")
    if op.font_size is not None:
        raise ValueError("set_fill_color does not accept font_size.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("set_fill_color does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("set_fill_color does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("set_fill_color does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="set_fill_color")
    _validate_exactly_one_cell_or_range(op, op_name="set_fill_color")
    if op.fill_color is None:
        raise ValueError("set_fill_color requires fill_color.")
    _validate_style_target_size(op, op_name="set_fill_color")


def _validate_set_dimensions(op: PatchOp) -> None:
    """Validate set_dimensions operation."""
    _validate_no_legacy_edit_fields(op, op_name="set_dimensions")
    if op.cell is not None or op.range is not None or op.base_cell is not None:
        raise ValueError("set_dimensions does not accept cell/range/base_cell.")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("set_dimensions does not accept row_count or col_count.")
    if op.bold is not None or op.color is not None or op.fill_color is not None:
        raise ValueError("set_dimensions does not accept bold, color, or fill_color.")
    if op.font_size is not None:
        raise ValueError("set_dimensions does not accept font_size.")
    if op.design_snapshot is not None:
        raise ValueError("set_dimensions does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="set_dimensions")
    has_rows = op.rows is not None
    has_columns = op.columns is not None
    if not has_rows and not has_columns:
        raise ValueError("set_dimensions requires rows and/or columns.")
    if has_rows and op.row_height is None:
        raise ValueError("set_dimensions requires row_height when rows is provided.")
    if has_columns and op.column_width is None:
        raise ValueError(
            "set_dimensions requires column_width when columns is provided."
        )
    if op.row_height is not None and op.row_height <= 0:
        raise ValueError("set_dimensions row_height must be > 0.")
    if op.column_width is not None and op.column_width <= 0:
        raise ValueError("set_dimensions column_width must be > 0.")


def _validate_auto_fit_columns(op: PatchOp) -> None:
    """Validate auto_fit_columns operation."""
    _validate_no_legacy_edit_fields(
        op, op_name="auto_fit_columns", allow_auto_fit_fields=True
    )
    if op.cell is not None or op.range is not None or op.base_cell is not None:
        raise ValueError("auto_fit_columns does not accept cell/range/base_cell.")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("auto_fit_columns does not accept row_count or col_count.")
    if op.bold is not None or op.color is not None or op.fill_color is not None:
        raise ValueError("auto_fit_columns does not accept bold, color, or fill_color.")
    if op.font_size is not None:
        raise ValueError("auto_fit_columns does not accept font_size.")
    if op.rows is not None or op.row_height is not None or op.column_width is not None:
        raise ValueError(
            "auto_fit_columns does not accept rows, row_height, or column_width."
        )
    if op.design_snapshot is not None:
        raise ValueError("auto_fit_columns does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="auto_fit_columns")
    if (
        op.min_width is not None
        and op.max_width is not None
        and op.min_width > op.max_width
    ):
        raise ValueError("auto_fit_columns requires min_width <= max_width.")


def _validate_merge_cells(op: PatchOp) -> None:
    """Validate merge_cells operation."""
    _validate_no_legacy_edit_fields(op, op_name="merge_cells")
    if op.cell is not None or op.base_cell is not None:
        raise ValueError("merge_cells does not accept cell or base_cell.")
    if op.range is None:
        raise ValueError("merge_cells requires range.")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("merge_cells does not accept row_count or col_count.")
    if op.bold is not None or op.color is not None or op.fill_color is not None:
        raise ValueError("merge_cells does not accept bold, color, or fill_color.")
    if op.font_size is not None:
        raise ValueError("merge_cells does not accept font_size.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("merge_cells does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("merge_cells does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("merge_cells does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="merge_cells")
    if _range_cell_count(op.range) < 2:
        raise ValueError("merge_cells requires a multi-cell range.")


def _validate_unmerge_cells(op: PatchOp) -> None:
    """Validate unmerge_cells operation."""
    _validate_no_legacy_edit_fields(op, op_name="unmerge_cells")
    if op.cell is not None or op.base_cell is not None:
        raise ValueError("unmerge_cells does not accept cell or base_cell.")
    if op.range is None:
        raise ValueError("unmerge_cells requires range.")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("unmerge_cells does not accept row_count or col_count.")
    if op.bold is not None or op.color is not None or op.fill_color is not None:
        raise ValueError("unmerge_cells does not accept bold, color, or fill_color.")
    if op.font_size is not None:
        raise ValueError("unmerge_cells does not accept font_size.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("unmerge_cells does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("unmerge_cells does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("unmerge_cells does not accept design_snapshot.")
    _validate_no_alignment_fields(op, op_name="unmerge_cells")


def _validate_set_alignment(op: PatchOp) -> None:
    """Validate set_alignment operation."""
    _validate_no_legacy_edit_fields(op, op_name="set_alignment")
    if op.base_cell is not None:
        raise ValueError("set_alignment does not accept base_cell.")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("set_alignment does not accept row_count or col_count.")
    if op.bold is not None or op.color is not None or op.fill_color is not None:
        raise ValueError("set_alignment does not accept bold, color, or fill_color.")
    if op.font_size is not None:
        raise ValueError("set_alignment does not accept font_size.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("set_alignment does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("set_alignment does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("set_alignment does not accept design_snapshot.")
    _validate_exactly_one_cell_or_range(op, op_name="set_alignment")
    if (
        op.horizontal_align is None
        and op.vertical_align is None
        and op.wrap_text is None
    ):
        raise ValueError(
            "set_alignment requires at least one of horizontal_align, vertical_align, or wrap_text."
        )
    _validate_style_target_size(op, op_name="set_alignment")


def _validate_set_style(op: PatchOp) -> None:
    """Validate set_style operation."""
    _validate_no_legacy_edit_fields(op, op_name="set_style")
    if op.base_cell is not None:
        raise ValueError("set_style does not accept base_cell.")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("set_style does not accept row_count or col_count.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("set_style does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("set_style does not accept row_height or column_width.")
    if op.design_snapshot is not None:
        raise ValueError("set_style does not accept design_snapshot.")
    _validate_exactly_one_cell_or_range(op, op_name="set_style")
    if (
        op.bold is None
        and op.font_size is None
        and op.color is None
        and op.fill_color is None
        and op.horizontal_align is None
        and op.vertical_align is None
        and op.wrap_text is None
    ):
        raise ValueError(
            "set_style requires at least one style field from: "
            "bold, font_size, color, fill_color, horizontal_align, vertical_align, wrap_text."
        )
    if op.font_size is not None and op.font_size <= 0:
        raise ValueError("set_style font_size must be > 0.")
    _validate_style_target_size(op, op_name="set_style")


def _validate_apply_table_style(op: PatchOp) -> None:
    """Validate apply_table_style operation."""
    _validate_no_legacy_edit_fields(
        op, op_name="apply_table_style", allow_table_fields=True
    )
    if op.cell is not None or op.base_cell is not None:
        raise ValueError("apply_table_style does not accept cell or base_cell.")
    if op.range is None:
        raise ValueError("apply_table_style requires range.")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("apply_table_style does not accept row_count or col_count.")
    if (
        op.bold is not None
        or op.color is not None
        or op.fill_color is not None
        or op.font_size is not None
    ):
        raise ValueError(
            "apply_table_style does not accept bold, color, fill_color, or font_size."
        )
    if op.rows is not None or op.columns is not None:
        raise ValueError("apply_table_style does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError(
            "apply_table_style does not accept row_height or column_width."
        )
    _validate_no_alignment_fields(op, op_name="apply_table_style")
    if op.design_snapshot is not None:
        raise ValueError("apply_table_style does not accept design_snapshot.")
    if op.style is None:
        raise ValueError("apply_table_style requires style.")


def _validate_restore_design_snapshot(op: PatchOp) -> None:
    """Validate restore_design_snapshot operation."""
    _validate_no_legacy_edit_fields(op, op_name="restore_design_snapshot")
    if op.cell is not None or op.range is not None or op.base_cell is not None:
        raise ValueError(
            "restore_design_snapshot does not accept cell/range/base_cell."
        )
    if op.row_count is not None or op.col_count is not None:
        raise ValueError(
            "restore_design_snapshot does not accept row_count or col_count."
        )
    if op.bold is not None or op.color is not None or op.fill_color is not None:
        raise ValueError(
            "restore_design_snapshot does not accept bold, color, or fill_color."
        )
    if op.font_size is not None:
        raise ValueError("restore_design_snapshot does not accept font_size.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("restore_design_snapshot does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError(
            "restore_design_snapshot does not accept row_height or column_width."
        )
    _validate_no_alignment_fields(op, op_name="restore_design_snapshot")
    if op.design_snapshot is None:
        raise ValueError("restore_design_snapshot requires design_snapshot.")


def _validate_create_chart(op: PatchOp) -> None:
    """Validate create_chart operation."""
    _validate_no_legacy_edit_fields(op, op_name="create_chart", allow_chart_fields=True)
    if op.cell is not None or op.range is not None or op.base_cell is not None:
        raise ValueError("create_chart does not accept cell/range/base_cell.")
    if op.row_count is not None or op.col_count is not None:
        raise ValueError("create_chart does not accept row_count or col_count.")
    if (
        op.bold is not None
        or op.color is not None
        or op.fill_color is not None
        or op.font_size is not None
    ):
        raise ValueError("create_chart does not accept style fields.")
    if op.rows is not None or op.columns is not None:
        raise ValueError("create_chart does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError("create_chart does not accept row_height or column_width.")
    _validate_no_alignment_fields(op, op_name="create_chart")
    if op.design_snapshot is not None:
        raise ValueError("create_chart does not accept design_snapshot.")
    if op.chart_type is None:
        raise ValueError("create_chart requires chart_type.")
    if op.data_range is None:
        raise ValueError("create_chart requires data_range.")
    if op.anchor_cell is None:
        raise ValueError("create_chart requires anchor_cell.")
    if op.titles_from_data is None:
        op.titles_from_data = True
    if op.series_from_rows is None:
        op.series_from_rows = False


def _validate_no_legacy_edit_fields(
    op: PatchOp,
    *,
    op_name: str,
    allow_table_fields: bool = False,
    allow_auto_fit_fields: bool = False,
    allow_chart_fields: bool = False,
) -> None:
    """Reject fields that are unrelated to design operations."""
    if op.expected is not None:
        raise ValueError(f"{op_name} does not accept expected.")
    if op.value is not None:
        raise ValueError(f"{op_name} does not accept value.")
    if op.values is not None:
        raise ValueError(f"{op_name} does not accept values.")
    if op.formula is not None:
        raise ValueError(f"{op_name} does not accept formula.")
    if not allow_table_fields:
        if op.style is not None:
            raise ValueError(f"{op_name} does not accept style.")
        if op.table_name is not None:
            raise ValueError(f"{op_name} does not accept table_name.")
    if not allow_auto_fit_fields:
        if op.min_width is not None:
            raise ValueError(f"{op_name} does not accept min_width.")
        if op.max_width is not None:
            raise ValueError(f"{op_name} does not accept max_width.")
    if not allow_chart_fields:
        _reject_optional_field(op_name, "chart_type", op.chart_type)
        _reject_optional_field(op_name, "data_range", op.data_range)
        _reject_optional_field(op_name, "category_range", op.category_range)
        _reject_optional_field(op_name, "anchor_cell", op.anchor_cell)
        _reject_optional_field(op_name, "chart_name", op.chart_name)
        _reject_optional_field(op_name, "width", op.width)
        _reject_optional_field(op_name, "height", op.height)
        _reject_optional_field(op_name, "titles_from_data", op.titles_from_data)
        _reject_optional_field(op_name, "series_from_rows", op.series_from_rows)
        _reject_optional_field(op_name, "chart_title", op.chart_title)
        _reject_optional_field(op_name, "x_axis_title", op.x_axis_title)
        _reject_optional_field(op_name, "y_axis_title", op.y_axis_title)


def _validate_no_design_fields(op: PatchOp, *, op_name: str) -> None:
    """Reject design-only fields for legacy value edit operations."""
    if op.row_count is not None or op.col_count is not None:
        raise ValueError(f"{op_name} does not accept row_count or col_count.")
    if op.rows is not None or op.columns is not None:
        raise ValueError(f"{op_name} does not accept rows or columns.")
    if op.row_height is not None or op.column_width is not None:
        raise ValueError(f"{op_name} does not accept row_height or column_width.")
    _reject_optional_field(op_name, "bold", op.bold)
    _reject_optional_field(op_name, "color", op.color)
    _reject_optional_field(op_name, "font_size", op.font_size)
    _reject_optional_field(op_name, "fill_color", op.fill_color)
    _reject_optional_field(op_name, "style", op.style)
    _reject_optional_field(op_name, "table_name", op.table_name)
    _validate_no_alignment_fields(op, op_name=op_name)
    _reject_optional_field(op_name, "design_snapshot", op.design_snapshot)
    _reject_optional_field(op_name, "min_width", op.min_width)
    _reject_optional_field(op_name, "max_width", op.max_width)
    _reject_optional_field(op_name, "chart_type", op.chart_type)
    _reject_optional_field(op_name, "data_range", op.data_range)
    _reject_optional_field(op_name, "category_range", op.category_range)
    _reject_optional_field(op_name, "anchor_cell", op.anchor_cell)
    _reject_optional_field(op_name, "chart_name", op.chart_name)
    _reject_optional_field(op_name, "width", op.width)
    _reject_optional_field(op_name, "height", op.height)
    _reject_optional_field(op_name, "titles_from_data", op.titles_from_data)
    _reject_optional_field(op_name, "series_from_rows", op.series_from_rows)
    _reject_optional_field(op_name, "chart_title", op.chart_title)
    _reject_optional_field(op_name, "x_axis_title", op.x_axis_title)
    _reject_optional_field(op_name, "y_axis_title", op.y_axis_title)


def _reject_optional_field(op_name: str, field_name: str, value: object) -> None:
    """Raise when an optional field is provided for an unsupported op."""
    if value is not None:
        raise ValueError(f"{op_name} does not accept {field_name}.")


def _validate_no_alignment_fields(op: PatchOp, *, op_name: str) -> None:
    """Reject alignment-only fields for unrelated operations."""
    if op.horizontal_align is not None:
        raise ValueError(f"{op_name} does not accept horizontal_align.")
    if op.vertical_align is not None:
        raise ValueError(f"{op_name} does not accept vertical_align.")
    if op.wrap_text is not None:
        raise ValueError(f"{op_name} does not accept wrap_text.")


def _validate_exactly_one_cell_or_range(op: PatchOp, *, op_name: str) -> None:
    """Ensure exactly one of cell/range is provided."""
    if op.base_cell is not None:
        raise ValueError(f"{op_name} does not accept base_cell.")
    has_cell = op.cell is not None
    has_range = op.range is not None
    if has_cell == has_range:
        raise ValueError(f"{op_name} requires exactly one of cell or range.")


def _validate_style_target_size(op: PatchOp, *, op_name: str) -> None:
    """Guard style edits against accidental huge targets."""
    target_count = 1 if op.cell is not None else _range_cell_count(op.range)
    if target_count > _MAX_STYLE_TARGET_CELLS:
        raise ValueError(
            f"{op_name} target exceeds max cells: {_MAX_STYLE_TARGET_CELLS}."
        )


def _range_cell_count(range_ref: str | None) -> int:
    """Return the number of cells represented by an A1 range."""
    if range_ref is None:
        raise ValueError("range is required.")
    return _shared_range_cell_count(range_ref)


def _split_a1(value: str) -> tuple[str, int]:
    """Split A1 notation into normalized (column_label, row_index)."""
    return _shared_split_a1(value)


def _normalize_column_identifier(value: str | int) -> str | int:
    """Normalize a column identifier preserving letter/index semantics."""
    if isinstance(value, int):
        if value < 1:
            raise ValueError("columns numeric values must be positive.")
        return value
    label = value.strip().upper()
    if not _COLUMN_LABEL_PATTERN.match(label):
        raise ValueError(f"Invalid column identifier: {value}")
    return label


def _column_label_to_index(label: str) -> int:
    """Convert Excel-style column label (A/AA) to 1-based index."""
    return _shared_column_label_to_index(label)


def _column_index_to_label(index: int) -> str:
    """Convert 1-based column index to Excel-style column label."""
    return _shared_column_index_to_label(index)


class PatchValue(BaseModel):
    """Normalized before/after value in patch diff."""

    kind: PatchValueKind
    value: str | int | float | None


class PatchDiffItem(BaseModel):
    """Applied change record for patch operations."""

    op_index: int
    op: PatchOpType
    sheet: str
    cell: str | None = None
    before: PatchValue | None = None
    after: PatchValue | None = None
    status: PatchStatus = "applied"


class PatchErrorDetail(BaseModel):
    """Structured error details for patch failures."""

    op_index: int
    op: PatchOpType
    sheet: str
    cell: str | None
    message: str
    hint: str | None = None
    expected_fields: list[str] = Field(default_factory=list)
    example_op: str | None = None
    error_code: str | None = None
    failed_field: str | None = None
    raw_com_message: str | None = None


class FormulaIssue(BaseModel):
    """Formula health-check finding."""

    sheet: str
    cell: str
    level: FormulaIssueLevel
    code: FormulaIssueCode
    message: str


class PatchRequest(BaseModel):
    """Input model for ExStruct MCP patch."""

    xlsx_path: Path
    ops: list[PatchOp]
    sheet: str | None = None
    out_dir: Path | None = None
    out_name: str | None = None
    on_conflict: OnConflictPolicy = "overwrite"
    auto_formula: bool = False
    dry_run: bool = False
    return_inverse_ops: bool = False
    preflight_formula_check: bool = False
    backend: PatchBackend = "auto"

    @model_validator(mode="after")
    def _validate_backend_constraints(self) -> PatchRequest:
        has_create_chart = any(op.op == "create_chart" for op in self.ops)
        if has_create_chart and self.backend == "openpyxl":
            raise ValueError(
                "create_chart is supported only on COM backend; backend='openpyxl' is not allowed."
            )
        if self.backend == "com":
            if self.dry_run or self.return_inverse_ops or self.preflight_formula_check:
                raise ValueError(
                    "backend='com' does not support dry_run, return_inverse_ops, "
                    "or preflight_formula_check."
                )
            if any(op.op == "restore_design_snapshot" for op in self.ops):
                raise ValueError(
                    "backend='com' does not support restore_design_snapshot operation."
                )
        if has_create_chart and (
            self.dry_run or self.return_inverse_ops or self.preflight_formula_check
        ):
            raise ValueError(
                "create_chart does not support dry_run, return_inverse_ops, or preflight_formula_check."
            )
        return self


class MakeRequest(BaseModel):
    """Input model for ExStruct MCP workbook creation."""

    out_path: Path
    ops: list[PatchOp] = Field(default_factory=list)
    sheet: str | None = None
    on_conflict: OnConflictPolicy = "overwrite"
    auto_formula: bool = False
    dry_run: bool = False
    return_inverse_ops: bool = False
    preflight_formula_check: bool = False
    backend: PatchBackend = "auto"

    @model_validator(mode="after")
    def _validate_backend_constraints(self) -> MakeRequest:
        has_create_chart = any(op.op == "create_chart" for op in self.ops)
        if has_create_chart and self.backend == "openpyxl":
            raise ValueError(
                "create_chart is supported only on COM backend; backend='openpyxl' is not allowed."
            )
        if self.backend == "com":
            if self.dry_run or self.return_inverse_ops or self.preflight_formula_check:
                raise ValueError(
                    "backend='com' does not support dry_run, return_inverse_ops, "
                    "or preflight_formula_check."
                )
            if any(op.op == "restore_design_snapshot" for op in self.ops):
                raise ValueError(
                    "backend='com' does not support restore_design_snapshot operation."
                )
        if has_create_chart and (
            self.dry_run or self.return_inverse_ops or self.preflight_formula_check
        ):
            raise ValueError(
                "create_chart does not support dry_run, return_inverse_ops, or preflight_formula_check."
            )
        return self


class OpenpyxlEngineResult(BaseModel):
    """Structured result returned by the openpyxl engine boundary.

    Attributes:
        patch_diff: Backend patch diff payload items.
        inverse_ops: Backend inverse operation payload items.
        formula_issues: Backend formula issue payload items.
        op_warnings: Backend warning messages emitted during apply.
    """

    patch_diff: list[PatchDiffItem] = Field(default_factory=list)
    inverse_ops: list[PatchOp] = Field(default_factory=list)
    formula_issues: list[FormulaIssue] = Field(default_factory=list)
    op_warnings: list[str] = Field(default_factory=list)


class PatchResult(BaseModel):
    """Output model for ExStruct MCP patch."""

    out_path: str
    patch_diff: list[PatchDiffItem] = Field(default_factory=list)
    inverse_ops: list[PatchOp] = Field(default_factory=list)
    formula_issues: list[FormulaIssue] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    error: PatchErrorDetail | None = None
    engine: PatchEngine


def _normalize_chart_range_reference(value: str) -> str:
    """Normalize chart range reference with optional sheet qualifier."""
    candidate = value.strip()
    match = _SHEET_QUALIFIED_A1_RANGE_PATTERN.match(candidate)
    if match is None:
        raise ValueError(f"Invalid chart range reference: {value}")
    sheet_prefix = match.group("sheet") or ""
    start = match.group("start").upper()
    end = match.group("end").upper()
    return f"{sheet_prefix}{start}:{end}"


def _normalize_hex_input(value: str, *, field_name: str) -> str:
    """Normalize HEX input into #RRGGBB or #AARRGGBB form."""
    text = value.strip().upper()
    if not _HEX_COLOR_PATTERN.match(text):
        raise ValueError(
            f"Invalid {field_name} format. Use 'RRGGBB', 'AARRGGBB', "
            "'#RRGGBB', or '#AARRGGBB'."
        )
    return text if text.startswith("#") else f"#{text}"


__all__ = [
    "AlignmentSnapshot",
    "BorderSideSnapshot",
    "BorderSnapshot",
    "ColumnDimensionSnapshot",
    "DesignSnapshot",
    "FillSnapshot",
    "FontSnapshot",
    "FormulaIssue",
    "MakeRequest",
    "MergeStateSnapshot",
    "OpenpyxlWorksheetProtocol",
    "OpenpyxlEngineResult",
    "PatchDiffItem",
    "PatchErrorDetail",
    "PatchOp",
    "PatchRequest",
    "PatchResult",
    "PatchValue",
    "RowDimensionSnapshot",
    "XlwingsRangeProtocol",
]
