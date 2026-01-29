from __future__ import annotations

import logging
from pathlib import Path

from pydantic import BaseModel

from exstruct import (
    ExtractionMode,
    ExStructEngine,
    StructOptions,
)
from exstruct.models import SheetData, WorkbookData

from .common import write_text

logger = logging.getLogger(__name__)


class ExstructTextConfig(BaseModel):
    """Configuration for ExStruct text extraction output."""

    mode: ExtractionMode = "standard"
    pretty: bool = False
    indent: int | None = None


def _filter_workbook_sheets(
    workbook: WorkbookData, sheet_scope: list[str] | None
) -> WorkbookData:
    """Return a workbook filtered to the requested sheet scope.

    Args:
        workbook: Extracted workbook payload from ExStruct.
        sheet_scope: Optional list of sheet names to keep.

    Returns:
        WorkbookData filtered to the requested sheets, or the original workbook if none match.
    """
    if not sheet_scope:
        return workbook
    sheets: dict[str, SheetData] = {
        name: sheet
        for name, sheet in workbook.sheets.items()
        if name in set(sheet_scope)
    }
    if not sheets:
        logger.warning("No matching sheets found for scope: %s", sheet_scope)
        return workbook
    return WorkbookData(book_name=workbook.book_name, sheets=sheets)


def extract_exstruct(
    xlsx_path: Path,
    out_txt: Path,
    sheet_scope: list[str] | None = None,
    *,
    config: ExstructTextConfig | None = None,
) -> None:
    """Extract workbook with ExStruct and write JSON text for LLM context.

    Args:
        xlsx_path: Excel workbook path.
        out_txt: Destination text file path.
        sheet_scope: Optional list of sheet names to keep.
        config: Optional ExStruct text extraction configuration.
    """
    resolved_config = config or ExstructTextConfig()
    engine = ExStructEngine(options=StructOptions(include_merged_values_in_rows=True))
    workbook = engine.extract(xlsx_path, mode=resolved_config.mode)
    workbook = _filter_workbook_sheets(workbook, sheet_scope)
    payload = workbook.to_json(
        pretty=resolved_config.pretty, indent=resolved_config.indent
    )

    lines = [
        "[DOC_META]",
        f"source={xlsx_path.name}",
        "method=exstruct",
        f"mode={resolved_config.mode}",
        "",
        "[CONTENT]",
        payload,
    ]
    write_text(out_txt, "\n".join(lines).strip() + "\n")
