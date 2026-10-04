"""Sequential real-Excel benchmark for Issue #159 color extraction."""

from __future__ import annotations

import argparse
from collections.abc import Iterator
from contextlib import ExitStack, contextmanager
from datetime import UTC, datetime
import hashlib
import json
import logging
import platform
from pathlib import Path
import re
import sys
from time import perf_counter
from typing import Any
from unittest.mock import patch

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Color, PatternFill

from exstruct.core import cells

_COLUMNS = 20
_MULTI_SHEET_COUNT = 8
_FIXTURE_CASES = (
    "static",
    "sparse_cf",
    "dense_cf",
    "overlapping_formula",
    "theme_indexed",
    "merged",
    "multiple_sheets",
)
_IGNORE_RED = {"FF0000"}
_CF_RENDER_EXPECTATIONS: dict[str, dict[str, int]] = {
    "sparse_cf": {"FF0000": 1},
    "dense_cf": {"0070C0": 1},
    # A1/B1 are direct-fill probe controls; require additional red and blue
    # cells so those controls alone cannot make a broken CF fixture pass.
    "overlapping_formula": {"FF0000": 2, "0070C0": 2},
}
_FIXED_PROPERTIES_TIME = datetime(2020, 1, 1, tzinfo=UTC)
_METRIC_KEYS = (
    "used_cells",
    "conditional_candidates",
    "display_format_calls",
    "rendered_candidates",
    "fallback",
    "static_duration_ms",
    "com_duration_ms",
    "duration_ms",
    "fallback_duration",
    "fallback_durations",
    "fallback_duration_ms",
    "fallback_durations_ms",
    "fallback_ms",
)
_METRIC_PATTERN = re.compile(
    r"\b(" + "|".join(re.escape(key) for key in _METRIC_KEYS) + r")\s*[=:]\s*([^,;\s]+)",
    re.IGNORECASE,
)


class _DebugRecordCollector(logging.Handler):
    """Keep the hybrid extractor's DEBUG records in the output JSON."""

    def __init__(self) -> None:
        super().__init__(logging.DEBUG)
        self.records: list[dict[str, Any]] = []

    def emit(self, record: logging.LogRecord) -> None:
        standard_fields = set(logging.makeLogRecord({}).__dict__)
        extras = {
            key: _jsonable(value)
            for key, value in record.__dict__.items()
            if key not in standard_fields and key not in {"message", "asctime"}
        }
        message = record.getMessage()
        fields = {key: value for key, value in extras.items() if key in _METRIC_KEYS}
        fields.update({key: value for key, value in _METRIC_PATTERN.findall(message)})
        self.records.append(
            {
                "logger": record.name,
                "level": record.levelname,
                "message": message,
                "metric_fields": fields,
                "extra": extras,
            }
        )


@contextmanager
def _capture_hybrid_debug() -> Iterator[_DebugRecordCollector]:
    logger = logging.getLogger("exstruct.core.color_hybrid")
    previous_level = logger.level
    previous_propagate = logger.propagate
    collector = _DebugRecordCollector()
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    logger.addHandler(collector)
    try:
        yield collector
    finally:
        logger.removeHandler(collector)
        logger.setLevel(previous_level)
        logger.propagate = previous_propagate


@contextmanager
def _count_display_format_wrapper() -> Iterator[dict[str, int]]:
    """Count calls through cells._get_display_format_color on either path."""
    original = cells._get_display_format_color
    counter = {"calls": 0}

    def counted(sheet: object, row: int, col: int) -> int | None:
        counter["calls"] += 1
        return original(sheet, row, col)  # type: ignore[arg-type]

    with ExitStack() as stack:
        stack.enter_context(patch.object(cells, "_get_display_format_color", counted))
        # The optimized module may cache the helper at import time. Patch that
        # alias as well so both paths are measured at the same call boundary.
        hybrid = sys.modules.get("exstruct.core.color_hybrid")
        if hybrid is not None and getattr(hybrid, "_get_display_format_color", None) is original:
            stack.enter_context(
                patch.object(hybrid, "_get_display_format_color", counted)
            )
        yield counter


def _jsonable(value: object) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_jsonable(item) for item in value]
    return repr(value)


def _populate_grid(sheet: Any, rows: int) -> None:
    for row in range(1, rows + 1):
        sheet.append([row * 100 + col for col in range(1, _COLUMNS + 1)])


def _single_sheet_fixture(name: str, rows: int) -> tuple[Workbook, list[str]]:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    _populate_grid(sheet, rows)
    red = PatternFill(fill_type="solid", fgColor="FFFF0000", bgColor="FFFF0000")
    blue = PatternFill(fill_type="solid", fgColor="FF0070C0", bgColor="FF0070C0")
    yellow = PatternFill(fill_type="solid", fgColor="FFFFFF00", bgColor="FFFFFF00")
    rules: list[str] = []
    extent = f"A1:T{rows}"

    if name == "static":
        for row in range(1, rows + 1):
            for col in range(1, _COLUMNS + 1):
                serial = (row - 1) * _COLUMNS + col
                if serial % 23 == 0:
                    sheet.cell(row, col).fill = red
                elif serial % 79 == 0:
                    sheet.cell(row, col).fill = yellow
    elif name == "sparse_cf":
        sparse_start = 2 if rows >= 2 else 1
        sparse_end = min(rows, sparse_start + 29)
        sparse_extent = f"D{sparse_start}:D{sparse_end}"
        sheet.conditional_formatting.add(
            sparse_extent,
            FormulaRule(
                formula=["MOD(ROW(),2)=0"],
                fill=red,
            ),
        )
        rules.append(
            f"one formula rule limited to {sparse_extent}; every other row matches"
        )
    elif name == "dense_cf":
        sheet.conditional_formatting.add(
            extent,
            FormulaRule(formula=["MOD(ROW()+COLUMN(),2)=0"], fill=blue),
        )
        rules.append("one formula rule across the grid; alternating cells match")
    elif name == "overlapping_formula":
        # Distinct saved fills make the range-vs-cell probe meaningful even
        # if Excel does not evaluate a conditional-formatting rule.
        sheet["A1"].fill = red
        sheet["B1"].fill = blue
        for row in range(2, rows + 1, 4):
            sheet.cell(row, 1).value = f"=SUM(B{row}:C{row})"
        sheet.conditional_formatting.add(
            extent,
            FormulaRule(formula=["MOD(ROW(),7)=0"], fill=red),
        )
        sheet.conditional_formatting.add(
            extent,
            FormulaRule(formula=["MOD(COLUMN(),5)=0"], fill=blue),
        )
        sheet.conditional_formatting.add(
            extent,
            FormulaRule(formula=["A1>1500"], fill=yellow),
        )
        # Make A1 and B1 a deterministic mixed-color probe. These highest-
        # priority rules stop the broader overlapping rules on those cells.
        sheet.conditional_formatting.add(
            "A1:B1",
            FormulaRule(formula=["COLUMN()=1"], fill=red, stopIfTrue=True),
        )
        sheet.conditional_formatting.add(
            "A1:B1",
            FormulaRule(formula=["COLUMN()=2"], fill=blue, stopIfTrue=True),
        )
        rules.extend(
            [
                "three overlapping formula rules across the grid",
                "two stop-if-true rules on A1:B1 for the mixed-range probe",
            ]
        )
    elif name == "theme_indexed":
        themed = PatternFill(
            fill_type="solid", fgColor=Color(theme=4, tint=0.25)
        )
        indexed = PatternFill(fill_type="solid", fgColor=Color(indexed=3))
        for row in range(1, rows + 1):
            for col in range(1, _COLUMNS + 1):
                serial = (row - 1) * _COLUMNS + col
                if serial % 17 == 0:
                    sheet.cell(row, col).fill = themed
                elif serial % 19 == 0:
                    sheet.cell(row, col).fill = indexed
    elif name == "merged":
        for row in range(1, rows + 1, 30):
            sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            sheet.cell(row, 1).fill = blue
    else:
        raise ValueError(f"Unknown single-sheet fixture: {name}")

    return workbook, rules


def _multiple_sheet_fixture(rows: int) -> tuple[Workbook, list[str]]:
    sheet_rows = min(rows, 30)
    workbook = Workbook()
    rules = [
        f"eight populated worksheets of {sheet_rows}x{_COLUMNS}; direct fills vary by sheet"
    ]
    for index in range(_MULTI_SHEET_COUNT):
        sheet = workbook.active if index == 0 else workbook.create_sheet()
        sheet.title = f"Sheet{index + 1:02d}"
        _populate_grid(sheet, sheet_rows)
        fill = PatternFill(
            fill_type="solid",
            fgColor=("FFFF0000" if index % 2 == 0 else "FF0070C0"),
        )
        for row in range(1, sheet_rows + 1):
            for col in range(1, _COLUMNS + 1):
                if (row * 3 + col + index) % 31 == 0:
                    sheet.cell(row, col).fill = fill
    return workbook, rules


def generate_fixtures(
    directory: Path, rows: int, cases: tuple[str, ...]
) -> list[dict[str, Any]]:
    directory.mkdir(parents=True, exist_ok=True)
    fixture_specs = (
        ("static", "rows x 20 values with deterministic direct fills"),
        ("sparse_cf", "conditional formatting with sparse matches"),
        ("dense_cf", "conditional formatting with dense alternating matches"),
        ("overlapping_formula", "overlapping CF rules, formulas, and mixed-range probe"),
        ("theme_indexed", "theme and indexed direct fills"),
        ("merged", "periodic horizontal merged ranges"),
        ("multiple_sheets", "eight populated worksheets capped at 30 rows each"),
    )
    fixtures: list[dict[str, Any]] = []
    for name, description in fixture_specs:
        if name not in cases:
            continue
        if name == "multiple_sheets":
            workbook, rules = _multiple_sheet_fixture(rows)
        else:
            workbook, rules = _single_sheet_fixture(name, rows)
        workbook.properties.created = _FIXED_PROPERTIES_TIME
        workbook.properties.modified = _FIXED_PROPERTIES_TIME
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        path = directory / f"{name}.xlsx"
        workbook.save(path)
        workbook.close()
        fixtures.append(
            {
                "name": name,
                "description": description,
                "path": str(path.resolve()),
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "rows": min(rows, 30) if name == "multiple_sheets" else rows,
                "columns": _COLUMNS,
                "sheets": _MULTI_SHEET_COUNT if name == "multiple_sheets" else 1,
                "conditional_formatting": rules,
                "expected_conditional_colors_min_cells": _CF_RENDER_EXPECTATIONS.get(
                    name
                ),
            }
        )
    return fixtures


def _reference_extract(
    book: Any, include_default_background: bool, ignore_colors: set[str] | None
) -> Any:
    cells._prepare_workbook_for_display_format(book)
    sheets = {}
    for sheet in book.sheets:
        cells._prepare_sheet_for_display_format(sheet)
        sheets[sheet.name] = cells._extract_sheet_colors_com(
            sheet, include_default_background, ignore_colors
        )
    return cells.WorkbookColorsMap(sheets=sheets)


def _map_fingerprint(color_map: Any) -> str:
    normalized = {
        sheet_name: {
            color: [list(position) for position in positions]
            for color, positions in sorted(sheet.colors_map.items())
        }
        for sheet_name, sheet in sorted(color_map.sheets.items())
    }
    encoded = json.dumps(normalized, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _map_summary(color_map: Any) -> dict[str, int]:
    return {
        "sheets": len(color_map.sheets),
        "color_keys": sum(len(sheet.colors_map) for sheet in color_map.sheets.values()),
        "mapped_cells": sum(
            len(positions)
            for sheet in color_map.sheets.values()
            for positions in sheet.colors_map.values()
        ),
    }


def _conditional_render_validation(
    fixture_name: str,
    color_map: Any | None,
    ignore_colors: set[str] | None,
) -> dict[str, Any] | None:
    """Require visible CF colors so equal empty maps cannot pass correctness."""
    expectations = _CF_RENDER_EXPECTATIONS.get(fixture_name)
    if expectations is None:
        return None
    ignored = {
        color.strip().lstrip("#").upper()[-6:] for color in (ignore_colors or set())
    }
    active = {color: minimum for color, minimum in expectations.items() if color not in ignored}
    report: dict[str, Any] = {
        "expected_min_cells_by_color": expectations,
        "ignored_expected_colors": sorted(set(expectations) & ignored),
        "checked_expected_colors": active,
    }
    if not active:
        report["status"] = "skipped_all_expected_colors_ignored"
        report["actual_cells_by_color"] = {}
        return report
    if color_map is None:
        report["status"] = "unavailable"
        report["actual_cells_by_color"] = {}
        return report

    sheet = color_map.sheets.get("Data")
    colors_map = sheet.colors_map if sheet is not None else {}
    actual = {color: len(colors_map.get(color, [])) for color in active}
    missing = {
        color: {"minimum": minimum, "actual": actual[color]}
        for color, minimum in active.items()
        if actual[color] < minimum
    }
    report["actual_cells_by_color"] = actual
    report["missing_minimums"] = missing
    report["status"] = "failed" if missing else "passed"
    return report


def _measure(
    book: Any,
    strategy: str,
    include_default_background: bool,
    ignore_colors: set[str] | None,
    fixture_name: str,
    repeat: int,
) -> tuple[dict[str, Any], Any | None]:
    context_counter = getattr(cells, "_display_format_calls", None)
    context_token = context_counter.set(0) if context_counter is not None else None
    collector: _DebugRecordCollector | None = None
    color_map = None
    error = None
    with _count_display_format_wrapper() as calls:
        if strategy == "hybrid":
            with _capture_hybrid_debug() as collector:
                started = perf_counter()
                try:
                    color_map = cells.extract_sheet_colors_map_com(
                        book,
                        include_default_background=include_default_background,
                        ignore_colors=ignore_colors,
                    )
                except Exception as exc:  # Retain failures beside successful timings.
                    error = f"{type(exc).__name__}: {exc}"
                elapsed_ms = (perf_counter() - started) * 1000
        else:
            started = perf_counter()
            try:
                color_map = _reference_extract(
                    book, include_default_background, ignore_colors
                )
            except Exception as exc:  # Retain failures beside successful timings.
                error = f"{type(exc).__name__}: {exc}"
            elapsed_ms = (perf_counter() - started) * 1000
    context_calls = (
        context_counter.get()
        if context_counter is not None and strategy != "hybrid"
        else None
    )
    if context_token is not None:
        context_counter.reset(context_token)

    record: dict[str, Any] = {
        "fixture": fixture_name,
        "repeat": repeat,
        "strategy": strategy,
        "include_default_background": include_default_background,
        "ignore_colors": sorted(ignore_colors) if ignore_colors else None,
        "elapsed_ms": elapsed_ms,
        "wrapper_display_format_calls": calls["calls"],
        "hybrid_debug_records": collector.records if collector is not None else [],
        "error": error,
    }
    if strategy != "hybrid":
        record["cells_context_display_format_calls"] = context_calls
    if color_map is not None:
        record["map_sha256"] = _map_fingerprint(color_map)
        record["map_summary"] = _map_summary(color_map)
    record["conditional_render_validation"] = _conditional_render_validation(
        fixture_name, color_map, ignore_colors
    )
    return record, color_map


def _probe_mixed_range(book: Any) -> dict[str, Any]:
    sheet = book.sheets["Data"]
    saved_before = bool(book.api.Saved)
    sheet.api.Activate()
    sheet.api.Calculate()
    result: dict[str, Any] = {
        "sheet": sheet.name,
        "range": "A1:B1",
        "workbook_saved_flag_before_calculate": saved_before,
        "workbook_saved_flag_after_calculate": bool(book.api.Saved),
    }
    per_cell: list[dict[str, Any]] = []
    for address in ("A1", "B1"):
        try:
            value = sheet.api.Range(address).DisplayFormat.Interior.Color
            per_cell.append(
                {
                    "address": address,
                    "value": int(value),
                    "repr": repr(value),
                    "type": type(value).__name__,
                }
            )
        except Exception as exc:
            per_cell.append({"address": address, "error": f"{type(exc).__name__}: {exc}"})
    try:
        value = sheet.api.Range("A1:B1").DisplayFormat.Interior.Color
        result["range_value"] = {
            "value": int(value),
            "repr": repr(value),
            "type": type(value).__name__,
        }
    except Exception as exc:
        result["range_value"] = {"error": f"{type(exc).__name__}: {exc}"}
    result["per_cell_values"] = per_cell
    values = [cell.get("value") for cell in per_cell]
    result["per_cell_colors_differ"] = (
        len(values) == 2 and None not in values and values[0] != values[1]
    )
    range_value = result["range_value"].get("value")
    result["range_matches_each_cell"] = (
        range_value is not None and len(values) == 2 and values == [range_value, range_value]
    )
    return result


def _write_output(path: Path, result: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(result, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    temporary.replace(path)


def _run(args: argparse.Namespace) -> dict[str, Any]:
    import xlwings as xw

    before = sorted(xw.apps.keys())
    if before:
        raise RuntimeError(
            f"Close existing Excel instances before this sequential benchmark: {before}"
        )

    fixtures = generate_fixtures(
        args.fixtures_dir.resolve(), args.rows, tuple(args.cases)
    )
    result: dict[str, Any] = {
        "schema_version": 1,
        "recorded_at": datetime.now(UTC).isoformat(),
        "platform": platform.platform(),
        "python": sys.version,
        "xlwings": xw.__version__,
        "excel_version": None,
        "configuration": {
            "rows": args.rows,
            "columns": _COLUMNS,
            "cases": args.cases,
            "repeats": args.repeats,
            "execution": "sequential; one owned Excel application; one workbook open at a time",
            "include_default_background_values": [False, True],
            "ignore_color_cases": [None, sorted(_IGNORE_RED)],
        },
        "fixtures": fixtures,
        "mixed_range_probe": None,
        "measurements": [],
        "comparisons": [],
    }

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = True
        try:
            app.api.EnableEvents = False
            app.api.AskToUpdateLinks = False
        except Exception:
            pass
        result["excel_version"] = str(app.version)

        for fixture_index, fixture in enumerate(fixtures, start=1):
            print(
                f"[issue159] fixture {fixture_index}/{len(fixtures)}: "
                f"{fixture['name']} ({fixture['rows']}x{_COLUMNS}, "
                f"{fixture['sheets']} sheet(s)); opening",
                flush=True,
            )
            book = app.books.open(
                fixture["path"], update_links=False, read_only=True
            )
            try:
                if fixture["name"] == "overlapping_formula":
                    try:
                        result["mixed_range_probe"] = _probe_mixed_range(book)
                    except Exception as exc:
                        result["mixed_range_probe"] = {
                            "error": f"{type(exc).__name__}: {exc}"
                        }

                for include_default_background in (False, True):
                    for ignore_colors in (None, _IGNORE_RED):
                        ignore_label = (
                            "none"
                            if ignore_colors is None
                            else ",".join(sorted(ignore_colors))
                        )
                        print(
                            f"  [config] include_default_background="
                            f"{include_default_background}, ignore_colors={ignore_label}",
                            flush=True,
                        )
                        pair: dict[str, Any] = {
                            "fixture": fixture["name"],
                            "include_default_background": include_default_background,
                            "ignore_colors": sorted(ignore_colors) if ignore_colors else None,
                            "repeats": [],
                        }
                        for repeat in range(1, args.repeats + 1):
                            order = (
                                ("reference", "hybrid")
                                if repeat % 2 == 1
                                else ("hybrid", "reference")
                            )
                            by_strategy: dict[str, Any | None] = {}
                            repeat_records: dict[str, dict[str, Any]] = {}
                            print(
                                f"    [repeat {repeat}/{args.repeats}] order="
                                f"{' -> '.join(order)}",
                                flush=True,
                            )
                            for strategy in order:
                                print(f"      [{strategy}] started", flush=True)
                                record, color_map = _measure(
                                    book,
                                    strategy,
                                    include_default_background,
                                    ignore_colors,
                                    fixture["name"],
                                    repeat,
                                )
                                result["measurements"].append(record)
                                repeat_records[strategy] = record
                                by_strategy[strategy] = color_map
                                outcome = (
                                    f"error={record['error']}"
                                    if record["error"]
                                    else "complete"
                                )
                                render_check = record[
                                    "conditional_render_validation"
                                ]
                                if render_check is not None:
                                    outcome += (
                                        "; CF visibility="
                                        f"{render_check['status']}"
                                    )
                                print(
                                    f"      [{strategy}] {outcome}; "
                                    f"{record['elapsed_ms']:.1f} ms, "
                                    f"wrapper calls={record['wrapper_display_format_calls']}",
                                    flush=True,
                                )
                            reference_map = by_strategy["reference"]
                            hybrid_map = by_strategy["hybrid"]
                            exact_equal = (
                                reference_map == hybrid_map
                                if reference_map is not None and hybrid_map is not None
                                else None
                            )
                            for record in repeat_records.values():
                                record["exact_workbook_colors_map_equal"] = exact_equal
                            print(
                                f"    [repeat {repeat}/{args.repeats}] exact map equality="
                                f"{exact_equal}",
                                flush=True,
                            )
                            pair["repeats"].append(
                                {
                                    "repeat": repeat,
                                    "order": list(order),
                                    "exact_workbook_colors_map_equal": exact_equal,
                                    "errors": {
                                        strategy: repeat_records[strategy]["error"]
                                        for strategy in ("reference", "hybrid")
                                        if repeat_records[strategy]["error"]
                                    },
                                    "conditional_render_validation": {
                                        strategy: repeat_records[strategy][
                                            "conditional_render_validation"
                                        ]
                                        for strategy in ("reference", "hybrid")
                                    },
                                }
                            )
                        result["comparisons"].append(pair)
                        _write_output(args.output.resolve(), result)
            finally:
                book.close()
    finally:
        if app is not None:
            app.quit()
        after = sorted(xw.apps.keys())
        if after != before:
            raise RuntimeError(
                f"Excel process cleanup mismatch: before={before}, after={after}"
            )

    _write_output(args.output.resolve(), result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="JSON result path")
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--rows", type=int, default=300)
    parser.add_argument(
        "--cases",
        nargs="+",
        choices=_FIXTURE_CASES,
        default=list(_FIXTURE_CASES),
        help="Fixture cases to run; for example: --cases static sparse_cf",
    )
    parser.add_argument(
        "--fixtures-dir",
        type=Path,
        help="Directory for generated xlsx fixtures (default: <output-stem>-fixtures)",
    )
    args = parser.parse_args()
    if args.repeats < 1:
        parser.error("--repeats must be positive")
    if args.rows < 1:
        parser.error("--rows must be positive")
    if args.fixtures_dir is None:
        args.fixtures_dir = args.output.with_name(args.output.stem + "-fixtures")
    if args.output.resolve() == args.fixtures_dir.resolve():
        parser.error("--output must not be the fixture directory")

    result = _run(args)
    pairs = [
        repeat["exact_workbook_colors_map_equal"]
        for comparison in result["comparisons"]
        for repeat in comparison["repeats"]
    ]
    failed = any(value is False or value is None for value in pairs) or any(
        record["error"]
        or (
            record["conditional_render_validation"] is not None
            and record["conditional_render_validation"]["status"]
            in {"failed", "unavailable"}
        )
        for record in result["measurements"]
    )
    print(
        json.dumps(
            {
                "output": str(args.output.resolve()),
                "fixtures_dir": str(args.fixtures_dir.resolve()),
                "fixture_count": len(result["fixtures"]),
                "measurement_count": len(result["measurements"]),
                "all_exact_comparisons_equal": not failed,
                "all_conditional_render_checks_passed": not any(
                    record["conditional_render_validation"] is not None
                    and record["conditional_render_validation"]["status"]
                    in {"failed", "unavailable"}
                    for record in result["measurements"]
                ),
                "mixed_range_probe": result["mixed_range_probe"],
            },
            indent=2,
            ensure_ascii=False,
        )
    )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
