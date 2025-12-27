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
) -> WorkbookData:
    """Extract workbook and return WorkbookData.

    Falls back to cells+tables if Excel COM is unavailable.

    Args:
        file_path: Workbook path.
        mode: Extraction mode.
        include_cell_links: Whether to include cell hyperlinks; None uses mode defaults.
        include_print_areas: Whether to include print areas; None defaults to True.
        include_auto_page_breaks: Whether to include auto page breaks.
        include_colors_map: Whether to include colors map; None uses mode defaults.
        include_default_background: Whether to include default background color.
        ignore_colors: Optional set of color keys to ignore.

    Returns:
        Extracted WorkbookData.

    Raises:
        ValueError: If mode is unsupported.
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
    )
    result = run_extraction_pipeline(inputs)
    return result.workbook
