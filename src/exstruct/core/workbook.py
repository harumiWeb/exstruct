from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
import logging
from pathlib import Path
from typing import Any
import warnings

from openpyxl import load_workbook
import xlwings as xw

logger = logging.getLogger(__name__)


@contextmanager
def openpyxl_workbook(
    file_path: Path, *, data_only: bool, read_only: bool
) -> Iterator[Any]:
    """Open an openpyxl workbook and ensure it is closed.

    Args:
        file_path: Workbook path.
        data_only: Whether to read formula results.
        read_only: Whether to open in read-only mode.

    Yields:
        openpyxl workbook instance.
    """
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="Unknown extension is not supported and will be removed",
            category=UserWarning,
            module="openpyxl",
        )
        warnings.filterwarnings(
            "ignore",
            message="Conditional Formatting extension is not supported and will be removed",
            category=UserWarning,
            module="openpyxl",
        )
        warnings.filterwarnings(
            "ignore",
            message="Cannot parse header or footer so it will be ignored",
            category=UserWarning,
            module="openpyxl",
        )
        wb = load_workbook(file_path, data_only=data_only, read_only=read_only)
    try:
        yield wb
    finally:
        try:
            wb.close()
        except Exception as exc:
            logger.debug("Failed to close openpyxl workbook. (%r)", exc)


@contextmanager
def xlwings_workbook(file_path: Path, *, visible: bool = False) -> Iterator[xw.Book]:
    """Open an Excel workbook via xlwings and close if created.

    Args:
        file_path: Workbook path.
        visible: Whether to show the Excel application window.

    Yields:
        xlwings workbook instance.
    """
    existing = _find_open_workbook(file_path)
    if existing:
        yield existing
        return

    app = xw.App(add_book=False, visible=visible)
    wb = app.books.open(str(file_path))
    try:
        yield wb
    finally:
        try:
            wb.close()
        except Exception as exc:
            logger.debug("Failed to close Excel workbook. (%r)", exc)
        try:
            app.quit()
        except Exception as exc:
            logger.debug("Failed to quit Excel application. (%r)", exc)


def _find_open_workbook(file_path: Path) -> xw.Book | None:
    """Return an existing workbook if already open in Excel.

    Args:
        file_path: Workbook path to search for.

    Returns:
        Existing xlwings workbook if open; otherwise None.
    """
    try:
        for app in xw.apps:
            for wb in app.books:
                resolved_path: Path | None = None
                try:
                    resolved_path = Path(wb.fullname).resolve()
                except Exception as exc:
                    logger.debug("Failed to resolve workbook path. (%r)", exc)
                if resolved_path is None:
                    continue
                if resolved_path == file_path.resolve():
                    return wb
    except Exception as exc:
        logger.debug("Failed to inspect open Excel workbooks. (%r)", exc)
        return None
    return None
