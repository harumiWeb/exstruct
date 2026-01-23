from __future__ import annotations

from pathlib import Path
from typing import Literal

from ..models import WorkbookData
from .pipeline import resolve_extraction_inputs, run_extraction_pipeline


def extract_workbook(  # noqa: C901
    file_path: str | Path,
    mode: Literal["light", "standard", "verbose"] = "standard",
    *,
    include_cell_links: bool | None = None,
    include_print_areas: bool | None = None,
    include_auto_page_breaks: bool = False,
    include_colors_map: bool | None = None,
    include_default_background: bool = False,
    ignore_colors: set[str] | None = None,
    include_formulas_map: bool | None = None,
    include_merged_cells: bool | None = None,
    include_merged_values_in_rows: bool = True,
) -> WorkbookData:
    """
    Extract a workbook into a structured WorkbookData representation.
    
    May fall back to cells+tables extraction if Excel COM automation is unavailable.
    
    Parameters:
        file_path (str | Path): Path to the workbook file.
        mode (Literal['light', 'standard', 'verbose']): Extraction mode that controls detail level.
        include_cell_links (bool | None): Include cell hyperlinks; `None` uses mode defaults.
        include_print_areas (bool | None): Include print areas; `None` defaults to True.
        include_auto_page_breaks (bool): Include automatic page break information.
        include_colors_map (bool | None): Include a colors map; `None` uses mode defaults.
        include_default_background (bool): Include default background color when present.
        ignore_colors (set[str] | None): Set of color keys to ignore during color mapping.
        include_formulas_map (bool | None): Include a map of cell formulas; `None` uses mode defaults.
        include_merged_cells (bool | None): Include merged cell ranges; `None` uses mode defaults.
        include_merged_values_in_rows (bool): Preserve merged cell values in row-wise output.
    
    Returns:
        WorkbookData: The extracted workbook representation.
    
    Raises:
        ValueError: If `mode` is not one of "light", "standard", or "verbose".
    """
    inputs = resolve_extraction_inputs(
        file_path,
        mode=mode,
        include_cell_links=include_cell_links,
        include_print_areas=include_print_areas,
        include_auto_page_breaks=include_auto_page_breaks,
        include_colors_map=include_colors_map,
        include_default_background=include_default_background,
        ignore_colors=ignore_colors,
        include_formulas_map=include_formulas_map,
        include_merged_cells=include_merged_cells,
        include_merged_values_in_rows=include_merged_values_in_rows,
    )
    result = run_extraction_pipeline(inputs)
    return result.workbook