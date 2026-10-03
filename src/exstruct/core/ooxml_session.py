"""Extraction-scoped, streaming access to core OOXML workbook parts (#149)."""

from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections.abc import Iterator
from dataclasses import dataclass, field
from heapq import heappop, heappush
import logging
from pathlib import Path
from typing import TYPE_CHECKING, Literal
from zipfile import ZipFile

from defusedxml import ElementTree

from ..models import CellRow, PrintArea
from .cell_types import MergedCellRange, SheetFormulasMap, WorkbookFormulasMap
from .cells import _normalize_cell_value, _normalize_formula_value
from .ooxml_package import OoxmlRelationship, read_relationships
from .ooxml_scalars import (
    BUILTIN_FORMATS,
    MAC_EPOCH,
    WINDOWS_EPOCH,
    UnsupportedOoxmlError,
    coordinate_to_tuple,
    from_excel,
    from_iso8601,
    get_column_letter,
    is_date_format,
    is_timedelta_format,
    range_boundaries,
    translate_formula,
)

if TYPE_CHECKING:
    from .ooxml_drawing import SheetDrawingData

logger = logging.getLogger(__name__)
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS = {"s": _MAIN}
_XML_PREFIX = f"{{{_MAIN}}}"


@dataclass(frozen=True)
class OoxmlSheet:
    """Sheet identity, visibility and resolved part path in workbook order."""

    name: str
    sheet_id: str
    state: str
    part_path: str
    kind: str


@dataclass(frozen=True)
class OoxmlDefinedName:
    """Raw defined-name text and optional workbook-local sheet index."""

    name: str
    text: str
    local_sheet_id: int | None = None


@dataclass(frozen=True)
class OoxmlTable:
    """Explicit table identity, extent and column names (no heuristics)."""

    name: str
    display_name: str
    ref: str
    columns: tuple[str, ...]


@dataclass(frozen=True)
class _Formula:
    """Stored formula text and optional shared-formula anchor identity."""

    coordinate: str
    text: str | None
    kind: str
    shared_id: str | None


@dataclass
class _SheetData:
    """Sparse worksheet artifacts cached only after successful XML parsing."""

    values: dict[tuple[int, int], object] = field(default_factory=dict)
    formulas: list[_Formula] = field(default_factory=list)
    merged: list[str] = field(default_factory=list)
    links: list[tuple[str, str]] = field(default_factory=list)
    table_ids: list[str] = field(default_factory=list)
    styles: dict[tuple[int, int], int] = field(default_factory=dict)
    has_row_styles: bool = False
    has_column_styles: bool = False
    covered_followers: set[tuple[int, int]] | None = None


def _string_text(node: ElementTree.Element) -> str:
    """Join plain/rich text runs without phonetic annotations."""
    parts: list[str] = []
    for child in node:
        if child.tag == f"{{{_MAIN}}}t":
            parts.append(child.text or "")
        elif child.tag == f"{{{_MAIN}}}r":
            parts.extend(text.text or "" for text in child.findall("s:t", _NS))
    return "".join(parts)


def _covered_followers(data: _SheetData) -> set[tuple[int, int]]:
    """Index stored merge followers once without expanding empty coordinates."""
    if data.covered_followers is not None:
        return data.covered_followers
    if not data.merged:
        data.covered_followers = set()
        return data.covered_followers
    columns_by_row: dict[int, list[int]] = {}
    for row, col in sorted(data.values):
        columns_by_row.setdefault(row, []).append(col)
    row_keys = list(columns_by_row)
    covered: set[tuple[int, int]] = set()
    for ref in data.merged:
        c1, r1, c2, r2 = _bounds(ref)
        start = bisect_left(row_keys, r1)
        stop = bisect_right(row_keys, r2)
        for row in row_keys[start:stop]:
            columns = columns_by_row[row]
            left = bisect_left(columns, c1)
            right = bisect_right(columns, c2)
            covered.update(
                (row, col) for col in columns[left:right] if (row, col) != (r1, c1)
            )
    data.covered_followers = covered
    return covered


def _bounds(ref: str) -> tuple[int, int, int, int]:
    """Validate finite cell bounds rather than expanding whole-row references."""
    c1, r1, c2, r2 = range_boundaries(ref)
    if c1 is None or r1 is None or c2 is None or r2 is None:
        raise ValueError(f"Unsupported cell range: {ref}")
    if not (1 <= c1 <= c2 <= 16384 and 1 <= r1 <= r2 <= 1048576):
        raise ValueError(f"Invalid cell range: {ref}")
    return c1, r1, c2, r2


def _split_defined_name_union(expression: str) -> list[str]:
    """Split a defined-name union at commas outside quoted sheet names."""
    parts: list[str] = []
    start = 0
    index = 0
    quoted = False
    while index < len(expression):
        char = expression[index]
        if char == "'":
            if quoted and index + 1 < len(expression) and expression[index + 1] == "'":
                index += 2
                continue
            quoted = not quoted
        elif char == "," and not quoted:
            parts.append(expression[start:index].strip())
            start = index + 1
        index += 1
    if quoted:
        raise ValueError("Unterminated quoted sheet name in defined name")
    parts.append(expression[start:].strip())
    return parts


def _sheet_name_separator(part: str) -> int:
    """Find the first exclamation mark outside a quoted sheet name."""
    quoted = False
    index = 0
    while index < len(part):
        if part[index] == "'":
            if quoted and index + 1 < len(part) and part[index + 1] == "'":
                index += 2
                continue
            quoted = not quoted
        elif part[index] == "!" and not quoted:
            return index
        index += 1
    return -1


def _defined_name_destinations(expression: str) -> list[tuple[str, str]]:
    """Extract worksheet/range pairs from a print-area defined name."""
    destinations: list[tuple[str, str]] = []
    for part in _split_defined_name_union(expression):
        separator = _sheet_name_separator(part)
        if separator < 0:
            continue
        title = part[:separator].strip()
        if title.startswith("'") and title.endswith("'"):
            title = title[1:-1].replace("''", "'")
        elif "'" in title:
            continue
        reference = part[separator + 1 :].strip()
        if title and reference:
            destinations.append((title, reference))
    return destinations


@dataclass(frozen=True)
class _BorderSideView:
    """The only border property consumed by the shared table heuristic."""

    style: str | None


@dataclass(frozen=True)
class _BorderView:
    """Sparse OOXML border representation matching worksheet cell access."""

    top: _BorderSideView
    bottom: _BorderSideView
    left: _BorderSideView
    right: _BorderSideView


_NO_BORDER_SIDE = _BorderSideView(None)
_NO_BORDER = _BorderView(
    top=_NO_BORDER_SIDE,
    bottom=_NO_BORDER_SIDE,
    left=_NO_BORDER_SIDE,
    right=_NO_BORDER_SIDE,
)


@dataclass(frozen=True)
class _TableRefView:
    """Minimal explicit-table object understood by cells.py."""

    ref: str


@dataclass(frozen=True)
class _OoxmlCellView:
    """Minimal cell surface used by cells.py border scanning."""

    value: object
    border: _BorderView


class _MergeIndex:
    """Index merge boundaries without allocating every covered coordinate.

    Row bands share the same active merges. Within each band, column segments
    select the earliest input merge, preserving the previous overlap semantics.
    Lookup uses two binary searches rather than scanning merge rectangles.
    """

    def __init__(self, merges: list[tuple[int, int, int, int]]) -> None:
        """Build row bands and their first-match column segments once."""
        events: dict[int, list[tuple[int, bool]]] = {}
        for index, (_, top, _, bottom) in enumerate(merges):
            events.setdefault(top, []).append((index, True))
            events.setdefault(bottom + 1, []).append((index, False))
        self.rows = sorted(events)
        self.columns: list[list[int]] = []
        self.matches: list[list[tuple[int, int, int, int] | None]] = []
        active: set[int] = set()
        for row in self.rows:
            for index, starts in events[row]:
                if starts:
                    active.add(index)
                else:
                    active.remove(index)
            columns, matches = self._column_segments(merges, active)
            self.columns.append(columns)
            self.matches.append(matches)

    @staticmethod
    def _column_segments(
        merges: list[tuple[int, int, int, int]], active_rows: set[int]
    ) -> tuple[list[int], list[tuple[int, int, int, int] | None]]:
        """Sweep column endpoints, resolving overlaps in input order."""
        events: dict[int, list[tuple[int, bool]]] = {}
        for index in active_rows:
            left, _, right, _ = merges[index]
            events.setdefault(left, []).append((index, True))
            events.setdefault(right + 1, []).append((index, False))
        columns = sorted(events)
        matches: list[tuple[int, int, int, int] | None] = []
        active: set[int] = set()
        priority: list[int] = []
        for column in columns:
            for index, starts in events[column]:
                if starts:
                    active.add(index)
                    heappush(priority, index)
                else:
                    active.remove(index)
            while priority and priority[0] not in active:
                heappop(priority)
            matches.append(merges[priority[0]] if priority else None)
        return columns, matches

    def at(self, row: int, column: int) -> tuple[int, int, int, int] | None:
        """Find the first covering merge in logarithmic lookup work."""
        band = bisect_right(self.rows, row) - 1
        if band < 0:
            return None
        segment = bisect_right(self.columns[band], column) - 1
        return self.matches[band][segment] if segment >= 0 else None


class _OoxmlWorksheetView:
    """Worksheet-shaped view over parsed OOXML values, styles and tables."""

    def __init__(
        self,
        data: _SheetData,
        borders: tuple[_BorderView, ...],
        tables: list[OoxmlTable],
    ) -> None:
        self._data = data
        self._borders = borders
        self._merges = [_bounds(reference) for reference in data.merged]
        self._merge_index = _MergeIndex(self._merges)
        self.tables = {
            str(index): _TableRefView(table.ref) for index, table in enumerate(tables)
        }
        occupied = list(data.values)
        for c1, r1, c2, r2 in self._merges:
            occupied.extend(((r1, c1), (r2, c2)))
        self._dimension = self._make_dimension(occupied)

    @staticmethod
    def _make_dimension(coordinates: list[tuple[int, int]]) -> str:
        """Compute the worksheet extent from stored and merged coordinates."""
        if not coordinates:
            return "A1:A1"
        rows = [row for row, _ in coordinates]
        columns = [column for _, column in coordinates]
        return (
            f"{get_column_letter(min(columns))}{min(rows)}:"
            f"{get_column_letter(max(columns))}{max(rows)}"
        )

    @property
    def max_row(self) -> int:
        """Return the bottom row of the represented worksheet extent."""
        _, _, _, max_row = range_boundaries(self._dimension)
        if max_row is None:
            raise ValueError("Worksheet dimension must have finite rows")
        return max_row

    @property
    def max_column(self) -> int:
        """Return the right column of the represented worksheet extent."""
        _, _, max_column, _ = range_boundaries(self._dimension)
        if max_column is None:
            raise ValueError("Worksheet dimension must have finite columns")
        return max_column

    def calculate_dimension(self) -> str:
        """Return the extent string consumed by border-map loading."""
        return self._dimension

    def _merge_at(self, row: int, column: int) -> tuple[int, int, int, int] | None:
        """Return the merge covering a coordinate, if present."""
        return self._merge_index.at(row, column)

    def _raw_border(self, row: int, column: int) -> _BorderView:
        """Resolve a cell's direct border style record."""
        style_id = self._data.styles.get((row, column), 0)
        if not 0 <= style_id < len(self._borders):
            raise UnsupportedOoxmlError(
                f"Cell style {style_id} has no corresponding border record"
            )
        return self._borders[style_id]

    def _cell_border(
        self, row: int, column: int, merged: tuple[int, int, int, int] | None
    ) -> _BorderView:
        """Apply the border inheritance used by openpyxl merged cells."""
        if merged is None:
            return self._raw_border(row, column)
        left, top, right, bottom = merged
        anchor = self._raw_border(top, left)
        if row == top and column == left:
            end_border = self._raw_border(bottom, right)
            return _BorderView(
                top=anchor.top,
                bottom=(
                    end_border.bottom
                    if end_border.bottom.style is not None
                    else anchor.bottom
                ),
                left=anchor.left,
                right=(
                    end_border.right
                    if end_border.right.style is not None
                    else anchor.right
                ),
            )
        bottom_right = self._raw_border(bottom, right)
        inherited_bottom = (
            bottom_right.bottom
            if bottom_right.bottom.style is not None
            else anchor.bottom
        )
        inherited_right = (
            bottom_right.right if bottom_right.right.style is not None else anchor.right
        )
        return _BorderView(
            top=anchor.top if row == top else _NO_BORDER_SIDE,
            bottom=inherited_bottom if row == bottom else _NO_BORDER_SIDE,
            left=anchor.left if column == left else _NO_BORDER_SIDE,
            right=inherited_right if column == right else _NO_BORDER_SIDE,
        )

    def cell(self, row: int, column: int) -> _OoxmlCellView:
        """Return a minimal cell with effective value and border semantics."""
        merged = self._merge_at(row, column)
        is_anchor = merged is None or (row == merged[1] and column == merged[0])
        value = self._data.values.get((row, column)) if is_anchor else None
        return _OoxmlCellView(value, self._cell_border(row, column, merged))

    def iter_rows(
        self,
        *,
        min_row: int,
        max_row: int,
        min_col: int,
        max_col: int,
        values_only: bool = False,
    ) -> Iterator[tuple[object, ...]]:
        """Yield a rectangular value view for cells.py's shared heuristic."""
        if not values_only:
            raise UnsupportedOoxmlError("OOXML worksheet view only yields values")
        for row in range(min_row, max_row + 1):
            yield tuple(
                self.cell(row, column).value for column in range(min_col, max_col + 1)
            )


class OoxmlExtractionSession:
    """Reuse one ZIP and parsed metadata; never calculate or edit a workbook.

    The light pipeline selects this session for supported OOXML workbooks.
    Worksheet XML and shared strings are streamed; returned sparse data is cached
    until close. Scalar format, coordinate and formula utilities use pure
    Python helpers; no openpyxl module or workbook model is required.
    """

    def __init__(self, file_path: Path) -> None:
        """Initialize lazy archive ownership and extraction-local caches."""
        self.file_path = file_path
        self._archive: ZipFile | None = None
        self._closed = False
        self._workbook_root: ElementTree.Element | None = None
        self._workbook_path = "xl/workbook.xml"
        self._relationships: dict[str, dict[str, OoxmlRelationship]] = {}
        self._sheet_cache: dict[str, _SheetData] = {}
        self._table_cache: dict[str, list[OoxmlTable]] = {}
        self._strings: list[str] | None = None
        self._formats: list[str] | None = None
        self._date_styles: set[int] = set()
        self._duration_styles: set[int] = set()
        self._epoch = WINDOWS_EPOCH
        self._styles_root: ElementTree.Element | None = None
        self._borders: tuple[_BorderView, ...] | None = None
        self._drawings: dict[str, SheetDrawingData] | None = None

    def __enter__(self) -> OoxmlExtractionSession:
        """Enter without opening a ZIP until a feature needs its parts."""
        self._ensure_open()
        return self

    def _ensure_open(self) -> None:
        """Reject access after the session has released its resources."""
        if self._closed:
            raise RuntimeError("OOXML extraction session is closed")

    def archive(self) -> ZipFile:
        """Open the archive lazily, once per session."""
        self._ensure_open()
        if self._archive is None:
            if self.file_path.suffix.lower() not in {".xlsx", ".xlsm"}:
                raise ValueError("OOXML extraction requires .xlsx or .xlsm")
            self._archive = ZipFile(self.file_path)
        return self._archive

    def relationships(self, source_path: str) -> dict[str, OoxmlRelationship]:
        """Resolve relationships once per source part."""
        self._ensure_open()
        if source_path not in self._relationships:
            self._relationships[source_path] = read_relationships(
                self.archive(), source_path
            )
        return self._relationships[source_path]

    def _workbook(self) -> ElementTree.Element:
        """Resolve the package officeDocument relationship and cache its XML."""
        self._ensure_open()
        if self._workbook_root is None:
            for rel in self.relationships("").values():
                if (
                    not rel.external
                    and rel.relationship_type == f"{_REL}/officeDocument"
                ):
                    self._workbook_path = rel.target
                    break
            self._workbook_root = ElementTree.fromstring(
                self.archive().read(self._workbook_path)
            )
        return self._workbook_root

    def sheets(self) -> list[OoxmlSheet]:
        """Return all sheet metadata, including non-worksheet sheet kinds."""
        root = self._workbook()
        rels = self.relationships(self._workbook_path)
        result: list[OoxmlSheet] = []
        for node in root.findall("s:sheets/s:sheet", _NS):
            rel = rels.get(node.get(f"{{{_REL}}}id", ""))
            if rel is None or rel.external:
                continue
            result.append(
                OoxmlSheet(
                    name=node.attrib["name"],
                    sheet_id=node.attrib["sheetId"],
                    state=node.get("state", "visible"),
                    part_path=rel.target,
                    kind=rel.relationship_type.rsplit("/", 1)[-1],
                )
            )
        return result

    def defined_names(self) -> list[OoxmlDefinedName]:
        """Preserve names and scope without evaluating named formulas."""
        return [
            OoxmlDefinedName(
                name=node.attrib["name"],
                text=node.text or "",
                local_sheet_id=int(node.attrib["localSheetId"])
                if "localSheetId" in node.attrib
                else None,
            )
            for node in self._workbook().findall("s:definedNames/s:definedName", _NS)
        ]

    def _related_part(self, kind: str) -> str | None:
        """Find a workbook-owned internal relationship of the requested kind."""
        self._workbook()
        return next(
            (
                rel.target
                for rel in self.relationships(self._workbook_path).values()
                if not rel.external and rel.relationship_type == f"{_REL}/{kind}"
            ),
            None,
        )

    def _shared_strings(self) -> list[str]:
        """Stream the shared-string dictionary once without retaining XML."""
        if self._strings is None:
            strings: list[str] = []
            part = self._related_part("sharedStrings")
            if part is not None:
                with self.archive().open(part) as stream:
                    root: ElementTree.Element | None = None
                    for event, node in ElementTree.iterparse(
                        stream, events=("start", "end")
                    ):
                        if root is None:
                            root = node
                        if event == "end" and node.tag == f"{{{_MAIN}}}si":
                            strings.append(_string_text(node))
                            root.remove(node)
            self._strings = strings
        return self._strings

    def _number_formats(self) -> list[str]:
        """Cache scalar number formats from built-in and custom style records."""
        if self._formats is None:
            formats: list[str] = []
            root = self._styles()
            if root is not None:
                custom = {
                    int(node.attrib["numFmtId"]): node.attrib["formatCode"]
                    for node in root.findall("s:numFmts/s:numFmt", _NS)
                }
                formats = [
                    custom.get(fmt_id, BUILTIN_FORMATS.get(fmt_id, "General"))
                    for node in root.findall("s:cellXfs/s:xf", _NS)
                    for fmt_id in [int(node.get("numFmtId", "0"))]
                ]
            self._formats = formats
            self._date_styles = {
                index for index, fmt in enumerate(formats) if is_date_format(fmt)
            }
            self._duration_styles = {
                index for index, fmt in enumerate(formats) if is_timedelta_format(fmt)
            }
            props = self._workbook().find(f"{_XML_PREFIX}workbookPr")
            self._epoch = (
                MAC_EPOCH
                if props is not None and props.get("date1904") in {"1", "true"}
                else WINDOWS_EPOCH
            )
        return self._formats

    def _styles(self) -> ElementTree.Element | None:
        """Read and cache the workbook's shared style part once."""
        if self._styles_root is None:
            part = self._related_part("styles")
            if part is not None:
                self._styles_root = ElementTree.fromstring(self.archive().read(part))
        return self._styles_root

    def _border_styles(self) -> tuple[_BorderView, ...]:
        """Map cell style IDs to border edges without constructing style models."""
        if self._borders is None:
            root = self._styles()
            if root is None:
                self._borders = (_NO_BORDER,)
            else:
                border_nodes = root.findall("s:borders/s:border", _NS)
                border_records = tuple(
                    _BorderView(
                        top=_BorderSideView(
                            node.find("s:top", _NS).get("style")
                            if node.find("s:top", _NS) is not None
                            else None
                        ),
                        bottom=_BorderSideView(
                            node.find("s:bottom", _NS).get("style")
                            if node.find("s:bottom", _NS) is not None
                            else None
                        ),
                        left=_BorderSideView(
                            node.find("s:left", _NS).get("style")
                            if node.find("s:left", _NS) is not None
                            else None
                        ),
                        right=_BorderSideView(
                            node.find("s:right", _NS).get("style")
                            if node.find("s:right", _NS) is not None
                            else None
                        ),
                    )
                    for node in border_nodes
                ) or (_NO_BORDER,)
                style_borders: list[_BorderView] = []
                for node in root.findall("s:cellXfs/s:xf", _NS):
                    border_id = int(node.get("borderId", "0"))
                    if not 0 <= border_id < len(border_records):
                        raise UnsupportedOoxmlError(
                            f"Cell style references missing border {border_id}"
                        )
                    style_borders.append(border_records[border_id])
                self._borders = tuple(style_borders) or (_NO_BORDER,)
        return self._borders

    def _cell_value(self, node: ElementTree.Element) -> object:
        """Decode the stored scalar cache and honor date/time number formats."""
        kind = node.get("t", "n")
        text = node.findtext(f"{_XML_PREFIX}v")
        if kind == "inlineStr":
            inline = node.find(f"{_XML_PREFIX}is")
            return _string_text(inline) if inline is not None else None
        if text in (None, "") or kind == "e":
            return None
        if kind == "s":
            index = int(text)
            if index < 0:
                raise ValueError("Negative OOXML shared-string index")
            return self._shared_strings()[index]
        if kind == "b":
            return bool(int(text))
        if kind == "d":
            return from_iso8601(text)
        if kind != "n":
            return text
        value = float(text) if "." in text or "e" in text or "E" in text else int(text)
        style = int(node.get("s", "0"))
        if self._formats is None:
            self._number_formats()
        if style in self._date_styles:
            try:
                return from_excel(
                    value, epoch=self._epoch, timedelta=style in self._duration_styles
                )
            except (OverflowError, ValueError):
                logger.warning("Skipping out-of-range OOXML date at %s", node.get("r"))
                return None
        return int(value) if int(value) == value else value

    def _read_sheet(self, sheet: OoxmlSheet) -> _SheetData:
        """Stream worksheet artifacts, publishing the cache only after success."""
        self._ensure_open()
        if sheet.name in self._sheet_cache:
            return self._sheet_cache[sheet.name]
        data = _SheetData()
        rels = self.relationships(sheet.part_path)
        stack: list[ElementTree.Element] = []
        row_number = 0
        column_number = 0
        with self.archive().open(sheet.part_path) as stream:
            for event, node in ElementTree.iterparse(stream, events=("start", "end")):
                if event == "start":
                    stack.append(node)
                    if node.tag == f"{{{_MAIN}}}row":
                        data.has_row_styles = data.has_row_styles or "s" in node.attrib
                        number = float(node.get("r", str(row_number + 1)))
                        if not number.is_integer():
                            raise ValueError("Invalid OOXML row number")
                        row_number = int(number)
                        column_number = 0
                    elif node.tag == f"{{{_MAIN}}}col":
                        data.has_column_styles = (
                            data.has_column_styles or "style" in node.attrib
                        )
                    continue
                tag = node.tag.removeprefix(_XML_PREFIX)
                position = None
                if tag == "c":
                    coordinate = node.get("r")
                    if coordinate:
                        position = coordinate_to_tuple(coordinate)
                        _, column_number = position
                    else:
                        column_number += 1
                        node.set("r", f"{get_column_letter(column_number)}{row_number}")
                        position = (row_number, column_number)
                if tag in {"c", "mergeCell", "hyperlink", "tablePart"}:
                    self._sheet_element(tag, node, data, rels, position=position)
                if tag in {"c", "row", "mergeCell", "hyperlink", "tablePart"}:
                    stack[-2].remove(node)
                stack.pop()
        self._sheet_cache[sheet.name] = data
        return data

    def _sheet_element(
        self,
        tag: str,
        node: ElementTree.Element,
        data: _SheetData,
        rels: dict[str, OoxmlRelationship],
        *,
        position: tuple[int, int] | None = None,
    ) -> None:
        """Collect a completed worksheet element before its XML is released."""
        if tag == "c":
            coordinate = node.attrib["r"]
            if position is None:
                position = coordinate_to_tuple(coordinate)
            data.values[position] = self._cell_value(node)
            if "s" in node.attrib:
                data.styles[position] = int(node.attrib["s"])
            formula = node.find(f"{_XML_PREFIX}f")
            if formula is not None:
                data.formulas.append(
                    _Formula(
                        coordinate,
                        formula.text,
                        formula.get("t", "normal"),
                        formula.get("si"),
                    )
                )
        elif tag == "mergeCell":
            data.merged.append(node.attrib["ref"])
        elif tag == "hyperlink":
            rel = rels.get(node.get(f"{{{_REL}}}id", ""))
            if (
                rel is not None
                and rel.external
                and rel.relationship_type == f"{_REL}/hyperlink"
            ):
                data.links.append((node.attrib["ref"], rel.target))
        elif tag == "tablePart":
            data.table_ids.append(node.attrib[f"{{{_REL}}}id"])

    def _worksheets(self) -> list[OoxmlSheet]:
        """Filter workbook-order metadata to worksheet relationships."""
        return [sheet for sheet in self.sheets() if sheet.kind == "worksheet"]

    def extract_cells(self, *, include_links: bool = False) -> dict[str, list[CellRow]]:
        """Read cached values with the current CellRow normalization contract."""
        return {
            sheet.name: self._cell_rows(
                self._read_sheet(sheet), include_links=include_links
            )
            for sheet in self._worksheets()
        }

    def _cell_rows(self, data: _SheetData, *, include_links: bool) -> list[CellRow]:
        """Normalize sparse values and restrict links to emitted rows in range."""
        rows: dict[int, dict[str, int | float | str]] = {}
        covered = _covered_followers(data)
        for (row, col), raw in sorted(data.values.items()):
            if (row, col) in covered:
                continue
            value = _normalize_cell_value(raw)
            if value is not None:
                rows.setdefault(row, {})[str(col - 1)] = value
        links: dict[int, dict[str, str]] = {}
        if include_links:
            row_keys = list(rows)
            for ref, target in data.links:
                c1, r1, c2, r2 = _bounds(ref)
                start = bisect_left(row_keys, r1)
                stop = bisect_right(row_keys, r2)
                for row in row_keys[start:stop]:
                    links.setdefault(row, {}).update(
                        {str(col - 1): target for col in range(c1, c2 + 1)}
                    )
        return [
            CellRow(r=row, c=values, links=links.get(row))
            for row, values in rows.items()
        ]

    def extract_formulas_map(self) -> WorkbookFormulasMap:
        """Normalize ordinary/array formulas and translate shared followers."""
        return WorkbookFormulasMap(
            sheets={
                sheet.name: self._formulas_map(sheet.name, self._read_sheet(sheet))
                for sheet in self._worksheets()
            }
        )

    def _formulas_map(self, name: str, data: _SheetData) -> SheetFormulasMap:
        """Group normalized formula text, translating known shared anchors."""
        shared = {
            f.shared_id: f for f in data.formulas if f.kind == "shared" and f.text
        }
        result: dict[str, list[tuple[int, int]]] = {}
        for formula in sorted(
            data.formulas, key=lambda f: coordinate_to_tuple(f.coordinate)
        ):
            text = _normalize_formula_value(formula.text)
            if text is None and formula.kind == "shared":
                anchor = shared.get(formula.shared_id)
                if anchor is not None:
                    try:
                        anchor_text = _normalize_formula_value(anchor.text)
                        if anchor_text is not None:
                            text = translate_formula(
                                anchor_text,
                                origin=anchor.coordinate,
                                target=formula.coordinate,
                            )
                    except ValueError:
                        logger.warning(
                            "Skipping unsupported shared formula at %s!%s",
                            name,
                            formula.coordinate,
                        )
            if text is not None:
                row, col = coordinate_to_tuple(formula.coordinate)
                result.setdefault(text, []).append((row, col - 1))
        return SheetFormulasMap(sheet_name=name, formulas_map=result)

    def extract_merged_cells(self) -> dict[str, list[MergedCellRange]]:
        """Read merged extents and unnormalized cached anchor display values."""
        result: dict[str, list[MergedCellRange]] = {}
        for sheet in self._worksheets():
            data = self._read_sheet(sheet)
            ranges: list[MergedCellRange] = []
            # openpyxl's MultiCellRange stores CellRange objects in a set.  Its
            # hash key is (min_row, min_col, max_row, max_col), so retain that
            # public iteration order without depending on worksheet models.
            initial_bounds = {
                (r1, c1, r2, c2)
                for ref in data.merged
                for c1, r1, c2, r2 in [_bounds(ref)]
            }
            # MultiCellRange's UniqueSequence descriptor rebuilds the first set
            # through an iterator; copying a set directly has different ordering.
            ordered_bounds = set(bounds for bounds in initial_bounds)
            for r1, c1, r2, c2 in ordered_bounds:
                raw = data.values.get((r1, c1))
                ranges.append(
                    MergedCellRange(
                        r1,
                        c1 - 1,
                        r2,
                        c2 - 1,
                        (str(raw) if raw is not None else "") or " ",
                    )
                )
            result[sheet.name] = ranges
        return result

    def extract_print_areas(self) -> dict[str, list[PrintArea]]:
        """Read scoped print areas without splitting commas in quoted names."""
        sheets = self.sheets()
        worksheet_names = {sheet.name for sheet in sheets if sheet.kind == "worksheet"}
        result: dict[str, list[PrintArea]] = {}
        for name in self.defined_names():
            if name.name != "_xlnm.Print_Area":
                continue
            for sheet_name, ref in _defined_name_destinations(name.text):
                if sheet_name not in worksheet_names:
                    continue
                if name.local_sheet_id is not None and (
                    not 0 <= name.local_sheet_id < len(sheets)
                    or sheets[name.local_sheet_id].name != sheet_name
                ):
                    continue
                try:
                    c1, r1, c2, r2 = _bounds(ref)
                except ValueError:
                    logger.warning("Skipping unsupported OOXML print area %s", ref)
                    continue
                result.setdefault(sheet_name, []).append(
                    PrintArea(r1=r1, c1=c1 - 1, r2=r2, c2=c2 - 1)
                )
        return result

    def extract_explicit_tables(self) -> dict[str, list[OoxmlTable]]:
        """Read table parts referenced by worksheets; no border inference."""
        result: dict[str, list[OoxmlTable]] = {}
        for sheet in self._worksheets():
            result[sheet.name] = self._explicit_tables_for_sheet(sheet)
        return result

    def _explicit_tables_for_sheet(self, sheet: OoxmlSheet) -> list[OoxmlTable]:
        """Read and cache table metadata for one worksheet only."""
        if sheet.name not in self._table_cache:
            tables: list[OoxmlTable] = []
            rels = self.relationships(sheet.part_path)
            for rel_id in self._read_sheet(sheet).table_ids:
                rel = rels.get(rel_id)
                if (
                    rel is None
                    or rel.external
                    or rel.relationship_type != f"{_REL}/table"
                ):
                    continue
                root = ElementTree.fromstring(self.archive().read(rel.target))
                tables.append(
                    OoxmlTable(
                        root.get("name", ""),
                        root.get("displayName", ""),
                        root.attrib["ref"],
                        tuple(
                            node.get("name", "")
                            for node in root.findall(
                                "s:tableColumns/s:tableColumn", _NS
                            )
                        ),
                    )
                )
            self._table_cache[sheet.name] = tables
        return self._table_cache[sheet.name]

    def detect_tables(
        self,
        sheet_name: str,
        mode: Literal["light", "libreoffice", "standard", "verbose"] = "light",
    ) -> list[str]:
        """Detect explicit tables and border candidates from streamed OOXML."""
        self._ensure_open()
        sheet = next(
            (item for item in self._worksheets() if item.name == sheet_name), None
        )
        if sheet is None:
            raise KeyError(sheet_name)
        data = self._read_sheet(sheet)
        if data.has_row_styles or data.has_column_styles:
            raise UnsupportedOoxmlError(
                "Worksheet row/column styles require compatibility extraction"
            )
        explicit = self._explicit_tables_for_sheet(sheet)
        view = _OoxmlWorksheetView(data, self._border_styles(), explicit)
        from .cells import detect_tables_openpyxl_ws

        return detect_tables_openpyxl_ws(view, mode=mode, cluster_backend="python")

    def read_drawings(self) -> dict[str, SheetDrawingData]:
        """Reuse the same ZIP for the existing best-effort drawing parser."""
        from .ooxml_drawing import read_sheet_drawings_from_archive

        self._ensure_open()
        if self._drawings is None:
            self._workbook()
            self._drawings = read_sheet_drawings_from_archive(
                self.archive(), workbook_path=self._workbook_path
            )
        return self._drawings

    def close(self) -> None:
        """Release the archive and caches, including after extraction errors."""
        if self._closed:
            return
        self._closed = True
        if self._archive is not None:
            self._archive.close()
        self._relationships.clear()
        self._sheet_cache.clear()
        self._table_cache.clear()
        self._strings = None
        self._formats = None
        self._date_styles.clear()
        self._duration_styles.clear()
        self._styles_root = None
        self._borders = None
        self._workbook_root = None
        self._drawings = None

    def __exit__(self, *exc_info: object) -> None:
        """Release owned resources on normal exit and on extraction errors."""
        self.close()
