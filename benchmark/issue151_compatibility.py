"""Reproduce Value2 compatibility limits without modifying source workbooks."""

from __future__ import annotations

import argparse
from datetime import datetime
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from benchmark.issue151_com_first import com_first, inputs_for
from exstruct.core import pipeline
from exstruct.core.backends.com_backend import ComBackend
from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend
from exstruct.core.workbook import xlwings_workbook


def probe() -> dict[str, object]:
    from openpyxl import Workbook

    with TemporaryDirectory(prefix="exstruct-issue151-") as directory:
        path = Path(directory) / "compatibility.xlsx"
        book = Workbook()
        sheet = book.active
        sheet["A1"] = "label"
        sheet["B2"] = datetime(2024, 1, 2, 3, 4, 5)
        sheet["C3"] = "=1+2"  # No stored formula cache in generated OOXML.
        sheet["D4"] = "#DIV/0!"
        sheet["E5"] = -2146826281  # Legitimate number, equal to a COM error code.
        sheet["F6"] = True
        sheet["G7"] = "external"
        sheet["G7"].hyperlink = "https://example.test/"
        sheet["H8"] = "internal"
        from openpyxl.worksheet.hyperlink import Hyperlink

        sheet["H8"].hyperlink = Hyperlink(ref="H8", location="Sheet!A1")
        book.save(path)
        book.close()
        saved = OpenpyxlBackend(path).extract_cells(include_links=True)
        with xlwings_workbook(path) as live:
            bulk = ComBackend(live).extract_cells(include_links=True)
            live.sheets[0].range("A1").value = "unsaved"
            unsaved = ComBackend(live).extract_cells(include_links=True)
        inputs = inputs_for(path, "standard")
        current = pipeline.run_extraction_pipeline(inputs)
        candidate = com_first(inputs)
        return {
            "file_cells": {
                name: [row.model_dump() for row in rows] for name, rows in saved.items()
            },
            "com_cells": {
                name: [row.model_dump() for row in rows] for name, rows in bulk.items()
            },
            "unsaved_com_cells": {
                name: [row.model_dump() for row in rows]
                for name, rows in unsaved.items()
            },
            "current_com_succeeded": current.state.com_succeeded,
            "candidate_com_succeeded": candidate.state.com_succeeded,
            "full_output_equal": current.workbook == candidate.workbook,
        }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = probe()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
