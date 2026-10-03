"""Dependency-free value contracts shared by extraction backends."""

from __future__ import annotations

from dataclasses import dataclass


# Use dataclasses for lightweight models
@dataclass(frozen=True)
class SheetColorsMap:
    """Background color map for a single worksheet."""

    sheet_name: str
    colors_map: dict[str, list[tuple[int, int]]]


@dataclass(frozen=True)
class WorkbookColorsMap:
    """Background color maps for all worksheets in a workbook."""

    sheets: dict[str, SheetColorsMap]

    def get_sheet(self, sheet_name: str) -> SheetColorsMap | None:
        """
        Retrieve the SheetColorsMap for a worksheet by name.

        Parameters:
            sheet_name (str): Name of the worksheet to retrieve.

        Returns:
            SheetColorsMap | None: The sheet's color map if present, `None` otherwise.
        """
        return self.sheets.get(sheet_name)


@dataclass(frozen=True)
class SheetFormulasMap:
    """Formula map for a single worksheet."""

    sheet_name: str
    formulas_map: dict[str, list[tuple[int, int]]]


@dataclass(frozen=True)
class WorkbookFormulasMap:
    """Formula maps for all worksheets in a workbook."""

    sheets: dict[str, SheetFormulasMap]

    def get_sheet(self, sheet_name: str) -> SheetFormulasMap | None:
        """
        Retrieve the formulas map for a worksheet.

        Parameters:
            sheet_name (str): Name of the worksheet to look up.

        Returns:
            SheetFormulasMap | None: The sheet's formulas map if present, `None` if the worksheet is not found.
        """
        return self.sheets.get(sheet_name)


@dataclass(frozen=True)
class MergedCellRange:
    """Merged cell range with normalized value."""

    r1: int
    c1: int
    r2: int
    c2: int
    v: str
