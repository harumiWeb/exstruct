"""Explicit, extraction-scoped ownership of lazy openpyxl workbooks."""

from __future__ import annotations

from contextlib import ExitStack
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.workbook.workbook import Workbook

from . import workbook


class OpenpyxlExtractionSession:
    """Share compatible workbook variants until final model construction ends."""

    def __init__(self, file_path: Path) -> None:
        self.file_path = file_path
        self._stack = ExitStack()
        self._workbooks: dict[bool, Workbook] = {}
        self._closed = False

    def __enter__(self) -> OpenpyxlExtractionSession:
        if self._closed:
            raise RuntimeError("Openpyxl extraction session is closed")
        return self

    def workbook(self, *, data_only: bool = True) -> Workbook:
        """Load each regular worksheet-capable variant only on first use."""
        if self._closed:
            raise RuntimeError("Openpyxl extraction session is closed")
        if data_only not in self._workbooks:
            self._workbooks[data_only] = self._stack.enter_context(
                workbook.openpyxl_workbook(
                    self.file_path, data_only=data_only, read_only=False
                )
            )
        return self._workbooks[data_only]

    def close(self) -> None:
        """Release all successfully loaded variants, including after errors."""
        if self._closed:
            return
        self._closed = True
        self._workbooks.clear()
        self._stack.close()

    def __exit__(self, *exc_info: object) -> None:
        self.close()
