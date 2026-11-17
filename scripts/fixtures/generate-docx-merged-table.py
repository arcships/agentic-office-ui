#!/usr/bin/env python3
"""Generate test-data/corpus/docx-merged-table.docx deterministically.

The fixture contains three tables:

1. A 3-column table whose first cell spans two rows (`w:vMerge` restart plus
   one continuation row), followed by an ordinary row.
2. A 3-column table whose first row starts with a `w:gridSpan` of 2.
3. A 2x2 outer table whose first cell contains a nested 2x2 table.

Only Python's standard library is used so the fixture can be regenerated in
CI without python-docx or other office tooling.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / "test-data" / "corpus" / "docx-merged-table.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

PACKAGE_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""


def paragraph(text: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def table_cell(content: str, properties: str = "") -> str:
    tc_pr = f"<w:tcPr>{properties}</w:tcPr>" if properties else ""
    return f"<w:tc>{tc_pr}{content}</w:tc>"


def table_row(cells: list[str]) -> str:
    return f"<w:tr>{''.join(cells)}</w:tr>"


def table(rows: list[str], widths: tuple[int, ...]) -> str:
    grid = "".join(f'<w:gridCol w:w="{width}"/>' for width in widths)
    borders = "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    return (
        "<w:tbl><w:tblPr>"
        '<w:tblW w:w="0" w:type="auto"/>'
        f"<w:tblBorders>{borders}</w:tblBorders>"
        "</w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>"
        f"{''.join(rows)}"
        "</w:tbl>"
    )


VMERGE_RESTART = '<w:vMerge w:val="restart"/>'
VMERGE_CONTINUE = "<w:vMerge/>"

vertical_merge_table = table(
    [
        table_row(
            [
                table_cell(paragraph("Merged A1"), VMERGE_RESTART),
                table_cell(paragraph("B1")),
                table_cell(paragraph("C1")),
            ]
        ),
        table_row(
            [
                table_cell("<w:p/>", VMERGE_CONTINUE),
                table_cell(paragraph("B2")),
                table_cell(paragraph("C2")),
            ]
        ),
        table_row(
            [
                table_cell(paragraph("A3")),
                table_cell(paragraph("B3")),
                table_cell(paragraph("C3")),
            ]
        ),
    ],
    (2500, 2500, 2500),
)

horizontal_merge_table = table(
    [
        table_row(
            [
                table_cell(paragraph("Wide header"), '<w:gridSpan w:val="2"/>'),
                table_cell(paragraph("Header C")),
            ]
        ),
        table_row(
            [
                table_cell(paragraph("D1")),
                table_cell(paragraph("E1")),
                table_cell(paragraph("F1")),
            ]
        ),
    ],
    (2500, 2500, 2500),
)

nested_inner_table = table(
    [
        table_row([table_cell(paragraph("Inner A1")), table_cell(paragraph("Inner B1"))]),
        table_row([table_cell(paragraph("Inner A2")), table_cell(paragraph("Inner B2"))]),
    ],
    (2000, 2000),
)

nested_outer_table = table(
    [
        table_row(
            [
                table_cell(
                    paragraph("Outer A1")
                    + nested_inner_table
                    + paragraph("Outer A1 tail")
                ),
                table_cell(paragraph("Outer B1")),
            ]
        ),
        table_row(
            [
                table_cell(paragraph("Outer A2")),
                table_cell(paragraph("Outer B2")),
            ]
        ),
    ],
    (3000, 3000),
)

DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}">
  <w:body>
    {paragraph("DOCX merged table fixture")}
    {vertical_merge_table}
    {paragraph("Spacer after vertical merge table")}
    {horizontal_merge_table}
    {paragraph("Spacer after horizontal merge table")}
    {nested_outer_table}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def write_fixture() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", PACKAGE_RELS),
            ("word/document.xml", DOCUMENT),
        ):
            info = zipfile.ZipInfo(name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o600 & 0xFFFF) << 16
            archive.writestr(info, text)


if __name__ == "__main__":
    write_fixture()
    print(f"wrote {OUTPUT}")
