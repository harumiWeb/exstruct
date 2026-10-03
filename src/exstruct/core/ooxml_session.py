"""Extraction-scoped, streaming access to core OOXML workbook parts (#149)."""

from __future__ import annotations

from dataclasses import dataclass, field
import logging
from pathlib import Path
from typing import TYPE_CHECKING
from zipfile import ZipFile

from defusedxml import ElementTree
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format, is_timedelta_format
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter, range_boundaries
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel, from_ISO8601

from ..models import CellRow, PrintArea
from .cell_types import MergedCellRange, SheetFormulasMap, WorkbookFormulasMap
from .cells import _normalize_cell_value, _normalize_formula_value
from .ooxml_package import OoxmlRelationship, read_relationships

if TYPE_CHECKING:
    from .ooxml_drawing import SheetDrawingData

logger = logging.getLogger(__name__)
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS = {"s": _MAIN}


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
    coordinate: str
    text: str | None
    kind: str
    shared_id: str | None


@dataclass
class _SheetData:
    values: dict[tuple[int, int], object] = field(default_factory=dict)
    formulas: list[_Formula] = field(default_factory=list)
    merged: list[str] = field(default_factory=list)
    links: list[tuple[str, str]] = field(default_factory=list)
    table_ids: list[str] = field(default_factory=list)


def _string_text(node: ElementTree.Element) -> str:
    """Join plain/rich text runs without phonetic annotations."""
    return "".join(t.text or "" for t in node.findall("s:t", _NS)) + "".join(
        t.text or "" for t in node.findall("s:r/s:t", _NS)
    )


def _bounds(ref: str) -> tuple[int, int, int, int]:
    """Validate finite cell bounds rather than expanding whole-row references."""
    c1, r1, c2, r2 = range_boundaries(ref)
    if None in (c1, r1, c2, r2):
        raise ValueError(f"Unsupported cell range: {ref}")
    if not (1 <= c1 <= c2 <= 16384 and 1 <= r1 <= r2 <= 1048576):
        raise ValueError(f"Invalid cell range: {ref}")
    return c1, r1, c2, r2


class OoxmlExtractionSession:
    """Reuse one ZIP and parsed metadata; never calculate or edit a workbook.

    This internal alternative is not selected by the extraction pipeline yet.
    Worksheet XML and shared strings are streamed; returned sparse data is cached
    until close. Scalar format and formula utilities reuse openpyxl, but no
    openpyxl workbook or worksheet object is constructed.
    """

    def __init__(self, file_path: Path) -> None:
        self.file_path = file_path
        self._archive: ZipFile | None = None
        self._closed = False
        self._workbook_root: ElementTree.Element | None = None
        self._workbook_path = "xl/workbook.xml"
        self._relationships: dict[str, dict[str, OoxmlRelationship]] = {}
        self._sheet_cache: dict[str, _SheetData] = {}
        self._strings: list[str] | None = None
        self._formats: list[str] | None = None
        self._drawings: dict[str, SheetDrawingData] | None = None

    def __enter__(self) -> OoxmlExtractionSession:
        self._ensure_open()
        return self

    def _ensure_open(self) -> None:
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
        if self._formats is None:
            formats: list[str] = []
            part = self._related_part("styles")
            if part is not None:
                root = ElementTree.fromstring(self.archive().read(part))
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
        return self._formats

    def _cell_value(self, node: ElementTree.Element) -> object:
        kind = node.get("t", "n")
        text = node.findtext("s:v", namespaces=_NS)
        if kind == "inlineStr":
            inline = node.find("s:is", _NS)
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
            return from_ISO8601(text)
        if kind != "n":
            return text
        value = float(text) if any(c in text for c in ".eE") else int(text)
        style = int(node.get("s", "0"))
        formats = self._number_formats()
        fmt = formats[style] if 0 <= style < len(formats) else "General"
        if is_date_format(fmt):
            props = self._workbook().find("s:workbookPr", _NS)
            epoch = (
                MAC_EPOCH
                if props is not None and props.get("date1904") in {"1", "true"}
                else WINDOWS_EPOCH
            )
            try:
                return from_excel(
                    value, epoch=epoch, timedelta=is_timedelta_format(fmt)
                )
            except (OverflowError, ValueError):
                logger.warning("Skipping out-of-range OOXML date at %s", node.get("r"))
                return None
        return int(value) if int(value) == value else value

    def _read_sheet(self, sheet: OoxmlSheet) -> _SheetData:
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
                        number = float(node.get("r", str(row_number + 1)))
                        if not number.is_integer():
                            raise ValueError("Invalid OOXML row number")
                        row_number = int(number)
                        column_number = 0
                    continue
                tag = node.tag.removeprefix(f"{{{_MAIN}}}")
                if tag == "c":
                    coordinate = node.get("r")
                    if coordinate:
                        _, column_number = coordinate_to_tuple(coordinate)
                    else:
                        column_number += 1
                        node.set("r", f"{get_column_letter(column_number)}{row_number}")
                self._sheet_element(tag, node, data, rels)
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
    ) -> None:
        if tag == "c":
            coordinate = node.attrib["r"]
            data.values[coordinate_to_tuple(coordinate)] = self._cell_value(node)
            formula = node.find("s:f", _NS)
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
        rows: dict[int, dict[str, int | float | str]] = {}
        merged = [_bounds(ref) for ref in data.merged]
        for (row, col), raw in sorted(data.values.items()):
            if any(
                c1 <= col <= c2 and r1 <= row <= r2 and (row, col) != (r1, c1)
                for c1, r1, c2, r2 in merged
            ):
                continue
            value = _normalize_cell_value(raw)
            if value is not None:
                rows.setdefault(row, {})[str(col - 1)] = value
        links: dict[int, dict[str, str]] = {}
        if include_links:
            for ref, target in data.links:
                c1, r1, c2, r2 = _bounds(ref)
                for row in rows:
                    if r1 <= row <= r2:
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
                        text = Translator(
                            _normalize_formula_value(anchor.text),
                            origin=anchor.coordinate,
                        ).translate_formula(formula.coordinate)
                    except (TranslatorError, ValueError):
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
            for ref in data.merged:
                c1, r1, c2, r2 = _bounds(ref)
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
        from openpyxl.workbook.defined_name import DefinedName

        sheets = self.sheets()
        worksheet_names = {sheet.name for sheet in sheets if sheet.kind == "worksheet"}
        result: dict[str, list[PrintArea]] = {}
        for name in self.defined_names():
            if name.name != "_xlnm.Print_Area":
                continue
            defined = DefinedName(name=name.name, attr_text=name.text)
            for sheet_name, ref in defined.destinations:
                sheet_name = sheet_name.replace("''", "'")
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
            result[sheet.name] = tables
        return result

    def read_drawings(self) -> dict[str, SheetDrawingData]:
        """Reuse the same ZIP for the existing best-effort drawing parser."""
        from .ooxml_drawing import read_sheet_drawings_from_archive

        self._ensure_open()
        if self._drawings is None:
            self._drawings = read_sheet_drawings_from_archive(self.archive())
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
        self._strings = None
        self._formats = None
        self._workbook_root = None
        self._drawings = None

    def __exit__(self, *exc_info: object) -> None:
        self.close()
