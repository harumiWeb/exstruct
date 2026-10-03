from pathlib import Path
from zipfile import ZipFile

from exstruct.core.ooxml_drawing import read_sheet_drawings_from_archive
from exstruct.core.ooxml_package import (
    OoxmlRelationship,
    normalize_part_path,
    read_relationships,
    relationship_part_path,
)


def test_relationship_paths_use_posix_package_paths() -> None:
    """Resolve relative and package-absolute internal relationship targets."""

    assert relationship_part_path("xl/workbook.xml") == "xl/_rels/workbook.xml.rels"
    assert relationship_part_path("") == "_rels/.rels"
    assert normalize_part_path("xl/worksheets", "../drawings/drawing1.xml") == (
        "xl/drawings/drawing1.xml"
    )
    assert normalize_part_path("xl/worksheets", "/xl/styles.xml") == "xl/styles.xml"


def test_read_relationships_resolves_internal_and_preserves_external_targets(
    tmp_path: Path,
) -> None:
    """Keep external URLs intact while resolving internal targets from the source."""

    book = tmp_path / "relationships.xlsx"
    with ZipFile(book, "w") as archive:
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="relative" Type="worksheet" Target="worksheets/../worksheets/sheet1.xml" />
              <Relationship Id="absolute" Type="styles" Target="/xl/styles.xml" />
              <Relationship Id="external" Type="hyperlink" Target="https://example.test/a?x=1&amp;y=two#part" TargetMode="External" />
            </Relationships>
            """,
        )

    with ZipFile(book) as archive:
        relationships = read_relationships(archive, "xl/workbook.xml")

    assert relationships == {
        "relative": OoxmlRelationship("xl/worksheets/sheet1.xml", "worksheet"),
        "absolute": OoxmlRelationship("xl/styles.xml", "styles"),
        "external": OoxmlRelationship(
            "https://example.test/a?x=1&y=two#part", "hyperlink", external=True
        ),
    }


def test_read_sheet_drawings_from_open_archive_ignores_external_relationships(
    tmp_path: Path,
) -> None:
    """Skip external worksheet/drawing/chart targets without closing the archive."""

    book = tmp_path / "external-relationships.xlsx"
    with ZipFile(book, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            """
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <sheets>
                <sheet name="ExternalWorksheet" sheetId="1" r:id="rId1" />
                <sheet name="ExternalDrawing" sheetId="2" r:id="rId2" />
                <sheet name="ExternalChart" sheetId="3" r:id="rId3" />
              </sheets>
            </workbook>
            """,
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="https://example.test/sheet.xml" TargetMode="External" />
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml" />
              <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/worksheets/sheet2.xml",
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />',
        )
        archive.writestr(
            "xl/worksheets/sheet3.xml",
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />',
        )
        archive.writestr(
            "xl/worksheets/_rels/sheet2.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="https://example.test/drawing.xml" TargetMode="External" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/worksheets/_rels/sheet3.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml" />
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/drawings/drawing1.xml",
            """
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <xdr:oneCellAnchor>
                <xdr:graphicFrame>
                  <xdr:nvGraphicFramePr>
                    <xdr:cNvPr id="1" name="Chart 1" />
                    <xdr:cNvGraphicFramePr />
                  </xdr:nvGraphicFramePr>
                  <xdr:xfrm />
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <c:chart r:id="rIdChart" />
                    </a:graphicData>
                  </a:graphic>
                </xdr:graphicFrame>
              </xdr:oneCellAnchor>
            </xdr:wsDr>
            """,
        )
        archive.writestr(
            "xl/drawings/_rels/drawing1.xml.rels",
            """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="https://example.test/chart.xml" TargetMode="External" />
            </Relationships>
            """,
        )

    with ZipFile(book) as archive:
        drawings = read_sheet_drawings_from_archive(archive)

        assert archive.fp is not None

    assert set(drawings) == {"ExternalChart"}
    assert drawings["ExternalChart"].charts == []
