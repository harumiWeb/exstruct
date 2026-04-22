from pathlib import Path
from zipfile import BadZipFile, ZipFile

from _pytest.monkeypatch import MonkeyPatch
from defusedxml import ElementTree

from exstruct.core.ooxml_drawing import (
    SheetDrawingData,
    SheetDrawingMetrics,
    _column_width_to_points,
    _marker_to_points,
    _merge_anchor_geometry,
    _parse_anchor_geometry,
    _parse_sheet_metrics,
    read_sheet_drawings,
)


def test_read_sheet_drawings_prefers_transform_geometry_for_flowchart_shapes() -> None:
    """Verify shape/connector positions use absolute ``xfrm`` geometry when present."""

    drawings = read_sheet_drawings(Path("sample/flowchart/sample-shape-connector.xlsx"))
    sheet = drawings["要件チェック"]

    first_shape = sheet.shapes[0]
    first_connector = sheet.connectors[0]

    assert (first_shape.ref.left, first_shape.ref.top) == (80, 46)
    assert (first_shape.ref.width, first_shape.ref.height) == (42, 42)
    assert (first_connector.ref.left, first_connector.ref.top) == (102, 88)
    assert (first_connector.ref.width, first_connector.ref.height) == (33, 80)


def test_read_sheet_drawings_skips_only_malformed_sheets(tmp_path: Path) -> None:
    """Verify malformed drawing XML only drops the affected worksheet."""

    book = tmp_path / "partial-drawings.xlsx"
    with ZipFile(book, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            """
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1" />
                <sheet name="Sheet2" sheetId="2" r:id="rId2" />
              </sheets>
            </workbook>
            """,
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            """
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetFormatPr defaultRowHeight="15" />
            </worksheet>
            """,
        )
        archive.writestr(
            "xl/worksheets/sheet2.xml",
            """
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetFormatPr defaultRowHeight="15" />
            </worksheet>
            """,
        )
        archive.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/worksheets/_rels/sheet2.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/drawings/drawing1.xml",
            """
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <xdr:oneCellAnchor>
                <xdr:from>
                  <xdr:col>0</xdr:col>
                  <xdr:colOff>0</xdr:colOff>
                  <xdr:row>0</xdr:row>
                  <xdr:rowOff>0</xdr:rowOff>
                </xdr:from>
                <xdr:ext cx="127000" cy="127000" />
                <xdr:sp>
                  <xdr:nvSpPr>
                    <xdr:cNvPr id="1" name="Shape 1" />
                    <xdr:cNvSpPr />
                  </xdr:nvSpPr>
                  <xdr:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0" />
                      <a:ext cx="127000" cy="127000" />
                    </a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst /></a:prstGeom>
                  </xdr:spPr>
                </xdr:sp>
              </xdr:oneCellAnchor>
            </xdr:wsDr>
            """,
        )
        archive.writestr("xl/drawings/drawing2.xml", "<xdr:wsDr")

    drawings = read_sheet_drawings(book)

    assert set(drawings) == {"Sheet1"}
    assert len(drawings["Sheet1"].shapes) == 1


def test_read_sheet_drawings_skips_only_badzip_sheets(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    """Verify a BadZipFile raised for one sheet does not clear sibling sheets."""

    book = tmp_path / "partial-badzip.xlsx"
    with ZipFile(book, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            """
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1" />
                <sheet name="Sheet2" sheetId="2" r:id="rId2" />
              </sheets>
            </workbook>
            """,
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            """
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetFormatPr defaultRowHeight="15" />
            </worksheet>
            """,
        )
        archive.writestr(
            "xl/worksheets/sheet2.xml",
            """
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetFormatPr defaultRowHeight="15" />
            </worksheet>
            """,
        )
        archive.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/worksheets/_rels/sheet2.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/drawings/drawing1.xml",
            """
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <xdr:oneCellAnchor>
                <xdr:from>
                  <xdr:col>0</xdr:col>
                  <xdr:colOff>0</xdr:colOff>
                  <xdr:row>0</xdr:row>
                  <xdr:rowOff>0</xdr:rowOff>
                </xdr:from>
                <xdr:ext cx="127000" cy="127000" />
                <xdr:sp>
                  <xdr:nvSpPr>
                    <xdr:cNvPr id="1" name="Shape 1" />
                    <xdr:cNvSpPr />
                  </xdr:nvSpPr>
                  <xdr:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0" />
                      <a:ext cx="127000" cy="127000" />
                    </a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst /></a:prstGeom>
                  </xdr:spPr>
                </xdr:sp>
              </xdr:oneCellAnchor>
            </xdr:wsDr>
            """,
        )
        archive.writestr("xl/drawings/drawing2.xml", "<xdr:wsDr />")

    from exstruct.core import ooxml_drawing

    original_parse = ooxml_drawing._parse_sheet_drawing

    def _patched_parse_sheet_drawing(
        archive: ZipFile,
        drawing_path: str,
        sheet_metrics: SheetDrawingMetrics,
    ) -> SheetDrawingData:
        if drawing_path == "xl/drawings/drawing2.xml":
            raise BadZipFile("bad member")
        return original_parse(archive, drawing_path, sheet_metrics)

    monkeypatch.setattr(
        "exstruct.core.ooxml_drawing._parse_sheet_drawing",
        _patched_parse_sheet_drawing,
    )

    drawings = read_sheet_drawings(book)

    assert set(drawings) == {"Sheet1"}
    assert len(drawings["Sheet1"].shapes) == 1


def test_marker_to_points_uses_custom_sheet_metrics() -> None:
    """Verify marker placement honors worksheet column and row sizing."""

    sheet_root = ElementTree.fromstring(
        """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetFormatPr defaultRowHeight="18" />
          <cols>
            <col min="1" max="2" width="3.58203125" customWidth="1" />
          </cols>
          <sheetData>
            <row r="2" ht="30" customHeight="1" />
          </sheetData>
        </worksheet>
        """
    )
    marker = ElementTree.fromstring(
        """
        <xdr:from xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <xdr:col>2</xdr:col>
          <xdr:colOff>12700</xdr:colOff>
          <xdr:row>1</xdr:row>
          <xdr:rowOff>25400</xdr:rowOff>
        </xdr:from>
        """
    )

    metrics = _parse_sheet_metrics(sheet_root)
    left, top = _marker_to_points(marker, metrics)

    expected_left = int(round(2 * _column_width_to_points(3.58203125) + 1))
    expected_top = 20
    assert (left, top) == (expected_left, expected_top)


def test_parse_anchor_geometry_uses_custom_metrics_for_two_cell_anchor() -> None:
    """Verify two-cell anchors use explicit worksheet sizes instead of fixed defaults."""

    anchor = ElementTree.fromstring(
        """
        <xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <xdr:from>
            <xdr:col>1</xdr:col>
            <xdr:colOff>12700</xdr:colOff>
            <xdr:row>2</xdr:row>
            <xdr:rowOff>12700</xdr:rowOff>
          </xdr:from>
          <xdr:to>
            <xdr:col>3</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>4</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
          </xdr:to>
        </xdr:twoCellAnchor>
        """
    )
    sheet_root = ElementTree.fromstring(
        """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetFormatPr defaultRowHeight="18" />
          <cols>
            <col min="1" max="3" width="3.58203125" customWidth="1" />
          </cols>
          <sheetData>
            <row r="1" ht="20" customHeight="1" />
            <row r="2" ht="30" customHeight="1" />
            <row r="3" ht="25" customHeight="1" />
            <row r="4" ht="22" customHeight="1" />
          </sheetData>
        </worksheet>
        """
    )

    metrics = _parse_sheet_metrics(sheet_root)
    geometry = _parse_anchor_geometry(anchor, metrics)
    col_points = _column_width_to_points(3.58203125)
    expected_left = int(round(col_points + 1))
    expected_top = int(round(20 + 30 + 1))
    expected_width = int(round(2 * col_points - 1))
    expected_height = int(round((25 + 22) - 1))

    assert geometry == (expected_left, expected_top, expected_width, expected_height)


def test_sheet_drawing_metrics_offset_cache_supports_out_of_order_reads() -> None:
    """Verify cached row/column offsets remain correct across repeated lookups."""

    metrics = SheetDrawingMetrics(
        default_column_width_points=10.0,
        default_row_height_points=5.0,
        column_width_points={1: 20.0, 3: 15.0},
        row_height_points={0: 7.0, 2: 8.0},
    )

    assert metrics.column_offset_points(4) == 55.0
    assert metrics.column_offset_points(2) == 30.0
    assert metrics.column_offset_points(5) == 65.0
    assert metrics.row_offset_points(3) == 20.0
    assert metrics.row_offset_points(1) == 7.0
    assert metrics.row_offset_points(4) == 25.0


def test_merge_anchor_geometry_prefers_transform_position_when_sized() -> None:
    """Verify absolute child geometry can override coarse anchor placement."""

    anchor = ElementTree.fromstring(
        """
        <xdr:oneCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <xdr:from>
            <xdr:col>2</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>3</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
          </xdr:from>
          <xdr:ext cx="127000" cy="254000" />
        </xdr:oneCellAnchor>
        """
    )

    left, top, width, height = _merge_anchor_geometry(
        anchor,
        left=80,
        top=46,
        width=42,
        height=42,
        prefer_transform_position_when_sized=True,
    )

    assert (left, top, width, height) == (80, 46, 42, 42)
