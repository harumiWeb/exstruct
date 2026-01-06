from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from ...models import CellRow, PrintArea
from ..cells import MergedCellRange, WorkbookColorsMap

CellData = dict[str, list[CellRow]]
PrintAreaData = dict[str, list[PrintArea]]
MergedCellData = dict[str, list[MergedCellRange]]


@dataclass(frozen=True)
class BackendConfig:
    """Configuration options shared across backends.

    Attributes:
        include_default_background: Whether to include default background colors.
        ignore_colors: Optional set of color keys to ignore.
    """

    include_default_background: bool
    ignore_colors: set[str] | None


class Backend(Protocol):
    """Protocol for backend implementations."""

    def extract_cells(self, *, include_links: bool) -> CellData:
        """Extract cell rows from the workbook."""

    def extract_print_areas(self) -> PrintAreaData:
        """Extract print areas from the workbook."""

    def extract_colors_map(
        self, *, include_default_background: bool, ignore_colors: set[str] | None
    ) -> WorkbookColorsMap | None:
        """Extract colors map from the workbook."""

    def extract_merged_cells(self) -> MergedCellData:
        """Extract merged cell ranges from the workbook."""
