"""Verify the generated office fixtures and their document-level features.

The manifest only pins file names, bytes and hashes. This script opens every
fixture with a real parser and checks the behaviours the black-box suites rely
on: DOCX paragraph/table/section counts, section page sizes and headers, PDF
page counts, `/Rotate` metadata, image-only pages, encryption and searchable
text (including CJK). It is executed by the PDF CI suite, so a regenerated
fixture that silently loses rotation or gains a text layer fails the build, and
by the DOCX suites before they assert page, furniture or export content.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from PIL import Image
from pypdf import PdfReader
from pypdf.errors import FileNotDecryptedError

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "apps" / "demo" / "public" / "samples"

valid_docx = [
    "demo.docx",
    "legal-contract.docx",
    "invoice-table.docx",
    "report-with-image.docx",
    "chinese-mixed.docx",
    "review-comments.docx",
    "review-threads.docx",
    "table-structures.docx",
    "sections-columns.docx",
    "page-sizes.docx",
    "docx-embedded-objects.docx",
    "docx-fields-controls.docx",
]
valid_pdf = [
    "sample.pdf",
    "scanned-invoice.pdf",
    "rotated-pages.pdf",
    "large-contract.pdf",
    "cjk-report.pdf",
]
valid_xlsx = [
    "financial-model.xlsx",
    "sales-table.xlsx",
    "sales-table-validated.xlsx",
    "charts-images.xlsx",
    "large-grid.xlsx",
    "xlsx-chart-families.xlsx",
    "xlsx-negative-databar-icons.xlsx",
    "xlsx-protected-1904.xlsx",
]
valid_pptx = ["deck-structure.pptx"]
valid_png = ["invoice.png", "contract-page.png"]
valid_json = ["field-citations.json", "ocr-layout.json", "manifest.json"]
corrupted = ["corrupted.docx", "corrupted.pdf", "corrupted.xlsx"]

# Page counts, per-page /Rotate values and searchable markers the PDF black-box
# suites assert in the browser. Keep this table in sync with tests/blackbox.
pdf_expectations: dict[str, dict[str, object]] = {
    "sample.pdf": {
        "pages": 4,
        "rotations": [0, 0, 0, 0],
        "page_markers": [f"ALPHA-{page}" for page in range(1, 5)],
    },
    "scanned-invoice.pdf": {
        "pages": 3,
        "rotations": [0, 0, 0],
        "image_only": True,
    },
    "rotated-pages.pdf": {
        "pages": 2,
        "rotations": [0, 90],
        "page_markers": ["Portrait Page", "Rotated Page"],
    },
    "large-contract.pdf": {
        "pages": 240,
        "minimum_pages": 200,
        "rotations": [0],
    },
    "cjk-report.pdf": {
        "pages": 2,
        "rotations": [0, 0],
        "page_markers": ["CJK-ALPHA-1", "CJK-ALPHA-2", "合同"],
    },
}
encrypted_pdf = {"file": "encrypted.pdf", "pages": 2, "password": "arcships", "wrong_password": "not-the-password"}


def pdf_rotation(page) -> int:
    return int(page.get("/Rotate") or 0)


def pdf_text(page) -> str:
    return page.extract_text() or ""


report: dict[str, dict[str, object]] = {
    "docx": {},
    "pdf": {},
    "xlsx": {},
    "pptx": {},
    "png": {},
    "json": {},
    "corrupted": {},
}
failures: list[str] = []


def expect(condition: bool, message: str) -> None:
    if not condition:
        failures.append(message)


# Exact body shape of every valid DOCX fixture. These numbers are part of the
# fixture contract: a regeneration that drops paragraphs or a table has to fail
# material verification instead of quietly narrowing what the browser suites
# can observe.
docx_expectations: dict[str, dict[str, int]] = {
    "demo.docx": {"paragraphs": 125, "tables": 1, "sections": 1},
    "legal-contract.docx": {"paragraphs": 125, "tables": 1, "sections": 1},
    "invoice-table.docx": {"paragraphs": 4, "tables": 1, "sections": 1},
    "report-with-image.docx": {"paragraphs": 9, "tables": 0, "sections": 1},
    "chinese-mixed.docx": {"paragraphs": 8, "tables": 0, "sections": 1},
    "review-comments.docx": {"paragraphs": 4, "tables": 0, "sections": 1},
    "review-threads.docx": {"paragraphs": 9, "tables": 1, "sections": 1},
    "table-structures.docx": {"paragraphs": 3, "tables": 2, "sections": 1},
    "sections-columns.docx": {"paragraphs": 41, "tables": 0, "sections": 3},
    "page-sizes.docx": {"paragraphs": 65, "tables": 0, "sections": 3},
    "docx-embedded-objects.docx": {"paragraphs": 8, "tables": 0, "sections": 1},
    "docx-fields-controls.docx": {"paragraphs": 9, "tables": 1, "sections": 1},
}


for name in valid_docx:
    path = SAMPLES / name
    doc = Document(path)
    report["docx"][name] = {
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "sections": len(doc.sections),
        "headerPreview": [
            " ".join(paragraph.text for paragraph in section.header.paragraphs).strip()
            for section in doc.sections
        ],
        "footerPreview": [
            " ".join(paragraph.text for paragraph in section.footer.paragraphs).strip()
            for section in doc.sections
        ],
        "bytes": path.stat().st_size,
    }
    expect(len(doc.paragraphs) > 0, f"{name}: expected at least one paragraph")
    # The black-box DOCX suites assert exact page, table and furniture content.
    # Recording the counts without asserting them let a regenerated fixture lose
    # half of its body while material verification still reported PASS.
    expected = docx_expectations[name]
    expect(
        len(doc.paragraphs) == expected["paragraphs"],
        f"{name}: expected {expected['paragraphs']} paragraphs, got {len(doc.paragraphs)}",
    )
    expect(
        len(doc.tables) == expected["tables"],
        f"{name}: expected {expected['tables']} tables, got {len(doc.tables)}",
    )
    expect(
        len(doc.sections) == expected["sections"],
        f"{name}: expected {expected['sections']} sections, got {len(doc.sections)}",
    )
    headers = report["docx"][name]["headerPreview"]
    footers = report["docx"][name]["footerPreview"]
    expect(
        all(headers) and all(footers),
        f"{name}: every section must keep a non-empty header and footer, got {headers} / {footers}",
    )

# Paper geometry is read from `w:pgSz`, so a fixture that silently switches back
# to Letter cannot satisfy the A4/landscape/mixed-size acceptance case.
page_size_expectations: dict[str, list[tuple[int, int, str]]] = {
    "legal-contract.docx": [(12240, 15840, "PORTRAIT")],
    "sections-columns.docx": [
        (12240, 15840, "PORTRAIT"),
        (15840, 12240, "LANDSCAPE"),
        (12240, 15840, "PORTRAIT"),
    ],
    "page-sizes.docx": [
        (11906, 16838, "PORTRAIT"),
        (16838, 11906, "LANDSCAPE"),
        (12240, 15840, "PORTRAIT"),
    ],
}
page_size_headers = {
    "sections-columns.docx": [
        "Section One Header",
        "Section Two Landscape Header",
        "Section Three Two Column Header",
    ],
    "page-sizes.docx": [
        "A4 Portrait Header",
        "A4 Landscape Header",
        "Letter Portrait Header",
    ],
}
for name, sizes in page_size_expectations.items():
    doc = Document(SAMPLES / name)
    actual = [
        (section.page_width.twips, section.page_height.twips, str(section.orientation).split(" ")[0])
        for section in doc.sections
    ]
    expect(actual == sizes, f"{name}: expected page sizes {sizes}, got {actual}")
    if name in page_size_headers:
        actual_headers = [
            " ".join(paragraph.text for paragraph in section.header.paragraphs).strip()
            for section in doc.sections
        ]
        expect(
            actual_headers == page_size_headers[name],
            f"{name}: expected section headers {page_size_headers[name]}, got {actual_headers}",
        )


# ---------------------------------------------------------------------------
# DOCX feature fixtures (issue #53)
#
# The manifest pins bytes and hashes, and python-docx proves the package opens.
# Neither detects a fixture that lost the feature it exists to test: a comment
# fixture with a single comment, a table fixture without a vMerge continuation,
# a section fixture with one section. These checks read the raw parts and fail
# when a regenerated fixture silently loses its review structure.
# ---------------------------------------------------------------------------

def docx_part(name: str, part: str) -> str:
    with zipfile.ZipFile(SAMPLES / name) as archive:
        try:
            return archive.read(part).decode("utf-8")
        except KeyError:
            return ""


review_threads = "review-threads.docx"
review_document = docx_part(review_threads, "word/document.xml")
review_comments = docx_part(review_threads, "word/comments.xml")
review_extended = docx_part(review_threads, "word/commentsExtended.xml")
review_header = docx_part(review_threads, "word/header1.xml")
review_report = {
    "comments": review_comments.count("<w:comment w:id="),
    "replies": review_extended.count("w15:paraIdParent"),
    "resolvedThreads": review_extended.count('w15:done="1"'),
    "bodyRanges": review_document.count("<w:commentRangeStart"),
    "headerRanges": review_header.count("<w:commentRangeStart"),
    "deletedRows": review_document.count("<w:trPr><w:del ") + review_document.count("<w:trPr><w:del/>"),
    "rowHeightRevisions": review_document.count("<w:trPrChange"),
}
report["docx"][f"{review_threads}#features"] = review_report
expect(review_report["comments"] == 11, f"{review_threads}: expected 11 comments, got {review_report['comments']}")
expect(review_report["replies"] == 3, f"{review_threads}: expected 3 threaded replies, got {review_report['replies']}")
expect(
    review_report["resolvedThreads"] == 1,
    f"{review_threads}: expected one resolved thread, got {review_report['resolvedThreads']}",
)
expect(
    review_document.count('<w:commentRangeStart w:id="3"/>') == 1
    and review_document.count('<w:commentRangeStart w:id="4"/>') == 1,
    f"{review_threads}: expected two comment ranges on the same paragraph",
)
# Comment 5 starts in one paragraph and ends in a later one.
start_index = review_document.find('<w:commentRangeStart w:id="5"/>')
end_index = review_document.find('<w:commentRangeEnd w:id="5"/>')
expect(
    start_index != -1
    and end_index != -1
    and review_document[start_index:end_index].count("</w:p>") >= 2,
    f"{review_threads}: comment 5 must span more than one paragraph",
)
expect(review_report["headerRanges"] == 1, f"{review_threads}: expected a header comment anchor")
expect(review_report["deletedRows"] == 1, f"{review_threads}: expected one deleted table row")
expect(
    review_report["rowHeightRevisions"] == 1,
    f"{review_threads}: expected one row-height revision",
)
expect(
    docx_part(review_threads, "word/footnotes.xml").count("<w:footnote w:id=") == 2,
    f"{review_threads}: expected two footnotes",
)
expect(
    docx_part(review_threads, "word/endnotes.xml").count("<w:endnote w:id=") == 2,
    f"{review_threads}: expected two endnotes",
)

table_structures = "table-structures.docx"
table_document = docx_part(table_structures, "word/document.xml")
table_header = docx_part(table_structures, "word/header1.xml")
table_report = {
    "tables": table_document.count("<w:tbl>"),
    "verticalMerges": table_document.count('<w:vMerge w:val="restart"/>'),
    "verticalContinuations": len(
        [m for m in table_document.split("<w:vMerge/>")[1:]]
    ),
    "gridSpans": table_document.count("<w:gridSpan "),
    "headerRows": table_document.count("<w:tblHeader/>"),
    "headerTables": table_header.count("<w:tbl>"),
    "headerTableRows": table_header.count("<w:tr>"),
}
report["docx"][f"{table_structures}#features"] = table_report
expect(table_report["tables"] == 3, f"{table_structures}: expected 3 tables including the nested one, got {table_report['tables']}")
expect(
    table_report["verticalMerges"] == 1,
    f"{table_structures}: expected one vMerge restart cell, got {table_report['verticalMerges']}",
)
expect(
    table_report["verticalContinuations"] == 2,
    f"{table_structures}: expected two vMerge continuation cells, got {table_report['verticalContinuations']}",
)
expect(table_report["gridSpans"] == 1, f"{table_structures}: expected one gridSpan cell")
expect(table_report["headerRows"] == 1, f"{table_structures}: expected one tblHeader row")
# The running head carries a table of its own; without it the "table inside a
# page header" case cannot be exercised at all.
expect(
    table_report["headerTables"] == 1,
    f"{table_structures}: expected one table inside the page header, got {table_report['headerTables']}",
)
expect(
    table_report["headerTableRows"] == 2,
    f"{table_structures}: expected the header table to have two rows, got {table_report['headerTableRows']}",
)
expect(
    "Table Register" in table_header and "Register" in table_header,
    f"{table_structures}: the header table lost its cell text",
)

sections_columns = "sections-columns.docx"
sections_document = docx_part(sections_columns, "word/document.xml")
section_header = docx_part(sections_columns, "word/header1.xml")
sections_report = {
    "sections": sections_document.count("<w:sectPr"),
    "landscape": sections_document.count('w:orient="landscape"'),
    "twoColumnSections": sections_document.count('<w:cols w:num="2"'),
    "textBoxes": sections_document.count("<w:txbxContent>"),
    "watermarks": section_header.count('behindDoc="1"'),
}
report["docx"][f"{sections_columns}#features"] = sections_report
expect(sections_report["sections"] == 3, f"{sections_columns}: expected 3 sections, got {sections_report['sections']}")
expect(sections_report["landscape"] == 1, f"{sections_columns}: expected one landscape section")
expect(
    sections_report["twoColumnSections"] == 1,
    f"{sections_columns}: expected one two-column section",
)
expect(sections_report["textBoxes"] == 1, f"{sections_columns}: expected one text box")
expect(sections_report["watermarks"] == 1, f"{sections_columns}: expected one behind-text watermark")
with zipfile.ZipFile(SAMPLES / sections_columns) as archive:
    expect(
        "word/media/watermark.png" in archive.namelist(),
        f"{sections_columns}: watermark image part is missing",
    )

# ---------------------------------------------------------------------------
# DOCX embedded objects, fields and tracked moves (issue #72)
# ---------------------------------------------------------------------------
embedded = "docx-embedded-objects.docx"
embedded_document = docx_part(embedded, "word/document.xml")
embedded_rels = docx_part(embedded, "word/_rels/document.xml.rels")
embedded_report = {
    "textBoxes": embedded_document.count("<w:txbxContent>"),
    "wordprocessingShapes": embedded_document.count("<wps:wsp"),
    "smartArtFrames": embedded_document.count("<dgm:relIds"),
    "chartFrames": embedded_document.count("<c:chart "),
    "oleObjects": embedded_document.count("<o:OLEObject"),
    "ommlEquations": embedded_document.count("<m:oMathPara>"),
}
report["docx"][f"{embedded}#features"] = embedded_report
for part in [
    "word/charts/chart1.xml",
    "word/diagrams/data1.xml",
    "word/diagrams/layout1.xml",
    "word/diagrams/quickStyle1.xml",
    "word/diagrams/colors1.xml",
]:
    expect(docx_part(embedded, part) != "", f"{embedded}: missing part {part}")
with zipfile.ZipFile(SAMPLES / embedded) as archive:
    embedded_parts = set(archive.namelist())
    ole_payload = archive.read("word/embeddings/oleObject1.bin")
for part in ["word/embeddings/oleObject1.bin", "word/media/ole-preview.png"]:
    expect(part in embedded_parts, f"{embedded}: missing binary part {part}")
expect(
    ole_payload[:8].hex() == "d0cf11e0a1b11ae1",
    f"{embedded}: OLE embedding is not a compound file",
)
expect(embedded_report["textBoxes"] == 1, f"{embedded}: expected one text box")
expect(
    embedded_report["wordprocessingShapes"] == 2,
    f"{embedded}: expected a text box and a separate shape, got {embedded_report['wordprocessingShapes']}",
)
expect(embedded_report["smartArtFrames"] == 1, f"{embedded}: expected one SmartArt frame")
expect(embedded_report["chartFrames"] == 1, f"{embedded}: expected one chart frame")
expect(embedded_report["oleObjects"] == 1, f"{embedded}: expected one OLE object")
expect(embedded_report["ommlEquations"] == 1, f"{embedded}: expected one display equation")
for token in ["charts/chart1.xml", "diagrams/data1.xml", "embeddings/oleObject1.bin"]:
    expect(token in embedded_rels, f"{embedded}: relationship target {token} is missing")
expect(
    "<c:barChart>" in docx_part(embedded, "word/charts/chart1.xml"),
    f"{embedded}: chart part lost its barChart payload",
)

fields = "docx-fields-controls.docx"
fields_document = docx_part(fields, "word/document.xml")
fields_report = {
    "contentControls": len(re.findall(r"<w:sdt\b", fields_document)),
    "moveFrom": fields_document.count("<w:moveFrom "),
    "moveTo": fields_document.count("<w:moveTo "),
    "pageRoman": " PAGE \\* roman " in fields_document,
    "numPagesAlpha": " NUMPAGES \\* ALPHA " in fields_document,
    "styleRefRoman": ' STYLEREF "Heading 1" \\* ROMAN ' in fields_document,
    "footnotes": docx_part(fields, "word/footnotes.xml").count("<w:footnote w:id="),
}
table_start = fields_document.find("<w:tbl>")
table_end = fields_document.find("</w:tbl>", table_start)
footnote_anchor = fields_document.find('<w:footnoteReference w:id="2"/>')
fields_report["footnoteInTable"] = (
    table_start != -1 and table_end != -1 and footnote_anchor != -1 and table_start < footnote_anchor < table_end
)
report["docx"][f"{fields}#features"] = fields_report
expect(fields_report["contentControls"] == 2, f"{fields}: expected block and inline content controls")
expect(fields_report["moveFrom"] == 1, f"{fields}: expected one moveFrom revision")
expect(fields_report["moveTo"] == 1, f"{fields}: expected one moveTo revision")
expect(fields_report["pageRoman"], f"{fields}: PAGE field lost its roman switch")
expect(fields_report["numPagesAlpha"], f"{fields}: NUMPAGES field lost its letter switch")
expect(fields_report["styleRefRoman"], f"{fields}: STYLEREF field lost its roman switch")
expect(fields_report["footnotes"] == 2, f"{fields}: expected two footnotes")
expect(fields_report["footnoteInTable"], f"{fields}: footnote reference must live inside the table")

for name in valid_pdf:
    path = SAMPLES / name
    expected = pdf_expectations[name]
    reader = PdfReader(str(path))
    expect(not reader.is_encrypted, f"{name}: fixture must not be encrypted")
    pages = len(reader.pages)
    texts = [pdf_text(page) for page in reader.pages]
    rotations = [pdf_rotation(page) for page in reader.pages]
    report["pdf"][name] = {
        "pages": pages,
        "rotations": rotations,
        "textLengths": [len(text) for text in texts],
        "bytes": path.stat().st_size,
    }
    expect(pages == int(expected["pages"]), f"{name}: expected {expected['pages']} pages, got {pages}")
    if "minimum_pages" in expected:
        expect(
            pages >= int(expected["minimum_pages"]),
            f"{name}: expected at least {expected['minimum_pages']} pages, got {pages}",
        )
    if "rotations" in expected:
        fixture_rotations = expected["rotations"]
        assert isinstance(fixture_rotations, list)
        if len(fixture_rotations) == 1:
            mismatched = [rotation for rotation in rotations if rotation != fixture_rotations[0]]
            expect(
                not mismatched,
                f"{name}: every page must be /Rotate {fixture_rotations[0]}, got {rotations}",
            )
        else:
            expect(
                rotations == fixture_rotations,
                f"{name}: expected page rotations {fixture_rotations}, got {rotations}",
            )
    if expected.get("image_only"):
        expect(
            all(not text.strip() for text in texts),
            f"{name}: image-only fixture must not expose a native text layer, got {texts}",
        )
        expect(
            any("/XObject" in (page.get("/Resources") or {}) for page in reader.pages),
            f"{name}: image-only fixture must draw at least one image XObject",
        )
    for marker in expected.get("page_markers", []):
        expect(
            any(marker in text for text in texts),
            f"{name}: extracted text is missing marker {marker!r}",
        )

encrypted_path = SAMPLES / encrypted_pdf["file"]
encrypted_reader = PdfReader(str(encrypted_path))
expect(encrypted_reader.is_encrypted, f"{encrypted_pdf['file']}: fixture must declare /Encrypt")
if encrypted_reader.is_encrypted:
    expect(
        encrypted_reader.decrypt(encrypted_pdf["wrong_password"]) == 0,
        f"{encrypted_pdf['file']}: a wrong password must not unlock the fixture",
    )
    locked = PdfReader(str(encrypted_path))
    expect(locked.is_encrypted, f"{encrypted_pdf['file']}: a fresh reader must still be locked")
    try:
        len(locked.pages)
        locked_error = ""
    except FileNotDecryptedError as error:
        locked_error = str(error)
    expect(locked_error != "", f"{encrypted_pdf['file']}: locked pages must raise FileNotDecryptedError")
    decrypted = PdfReader(str(encrypted_path))
    expect(
        decrypted.decrypt(encrypted_pdf["password"]) == 1,
        f"{encrypted_pdf['file']}: documented password must unlock the fixture",
    )
    decrypted_texts = [pdf_text(page) for page in decrypted.pages]
    report["pdf"][encrypted_pdf["file"]] = {
        "pages": len(decrypted.pages),
        "encrypted": True,
        "textLengths": [len(text) for text in decrypted_texts],
        "bytes": encrypted_path.stat().st_size,
    }
    expect(
        len(decrypted.pages) == int(encrypted_pdf["pages"]),
        f"{encrypted_pdf['file']}: expected {encrypted_pdf['pages']} pages after decryption",
    )
    expect(
        all("Encrypted" in text or "Payment" in text for text in decrypted_texts),
        f"{encrypted_pdf['file']}: decrypted pages lost their text layer",
    )

acro_pdf = "pdf-acroform-annotations.pdf"
acro_reader = PdfReader(str(SAMPLES / acro_pdf))
expect(not acro_reader.is_encrypted, f"{acro_pdf}: fixture must not be encrypted")
acro_root = acro_reader.trailer["/Root"]
acro_form = acro_root["/AcroForm"]
acro_fields = [field.get_object() for field in acro_form["/Fields"]]
page_zero_annots = [annotation.get_object() for annotation in acro_reader.pages[0]["/Annots"]]
acro_report = {
    "pages": len(acro_reader.pages),
    "fieldTypes": sorted({str(field.get("/FT")) for field in acro_fields}),
    "annotationTypes": sorted({str(annotation.get("/Subtype")) for annotation in page_zero_annots}),
    "attachments": sorted(acro_reader.attachments),
    "ocg": "/OCProperties" in acro_root,
    "ocgContent": b"/OC /MC0 BDC" in acro_reader.pages[1].get_contents().get_data(),
    "permissions": "/DocMDP" in acro_root.get("/Perms", {}),
    "sigFlags": int(acro_form.get("/SigFlags", 0)),
    "bytes": (SAMPLES / acro_pdf).stat().st_size,
}
report["pdf"][f"{acro_pdf}#features"] = acro_report
expect(acro_report["pages"] == 2, f"{acro_pdf}: expected two pages")
expect(
    {"/Tx", "/Btn", "/Ch", "/Sig"}.issubset(set(acro_report["fieldTypes"])),
    f"{acro_pdf}: expected text, check, choice and signature fields, got {acro_report['fieldTypes']}",
)
expect(
    {"/Highlight", "/Text", "/FreeText", "/Link"}.issubset(set(acro_report["annotationTypes"])),
    f"{acro_pdf}: expected highlight, text, free-text and link annotations, got {acro_report['annotationTypes']}",
)
expect(
    acro_report["attachments"] == ["fixture-meta.json", "fixture-readme.txt"],
    f"{acro_pdf}: expected both embedded attachments, got {acro_report['attachments']}",
)
expect(acro_report["ocg"], f"{acro_pdf}: catalog lost /OCProperties")
expect(acro_report["ocgContent"], f"{acro_pdf}: page content lost its marked OCG block")
expect(acro_report["permissions"], f"{acro_pdf}: catalog lost its DocMDP permission metadata")
expect(acro_report["sigFlags"] == 3, f"{acro_pdf}: expected /SigFlags 3")

large_pdf = "pdf-1000-pages.pdf"
large_reader = PdfReader(str(SAMPLES / large_pdf))
expect(not large_reader.is_encrypted, f"{large_pdf}: fixture must not be encrypted")
large_texts = [pdf_text(large_reader.pages[0]), pdf_text(large_reader.pages[-1])]
report["pdf"][f"{large_pdf}#features"] = {
    "pages": len(large_reader.pages),
    "firstText": large_texts[0][:80],
    "lastText": large_texts[1][:80],
    "bytes": (SAMPLES / large_pdf).stat().st_size,
}
expect(len(large_reader.pages) == 1000, f"{large_pdf}: expected 1000 pages")
expect("PAGE-0001" in large_texts[0], f"{large_pdf}: first page marker is missing")
expect("PAGE-1000" in large_texts[1], f"{large_pdf}: last page marker is missing")
expect(
    (SAMPLES / large_pdf).stat().st_size <= 2 * 1024 * 1024,
    f"{large_pdf}: fixture must stay below 2 MiB",
)

# XLSX fixture contract read through openpyxl. The browser and component
# suites assert these exact features, so a regeneration that drops the merge,
# the freeze panes, the validation rules or the two-letter column widths must
# fail material verification instead of leaving an empty fixture behind.
xlsx_expectations: dict[str, dict[str, object]] = {
    "financial-model.xlsx": {
        "sheets": ["Assumptions", "P&L", "Notes"],
        "charts": 1,
        "merged": [],
        "freeze": None,
        "validations": 0,
    },
    "sales-table.xlsx": {
        "sheets": ["Sales Data"],
        "charts": 0,
        "merged": [],
        "freeze": "A2",
        "validations": 0,
        "tables": ["SalesRecords"],
        "maxRow": 181,
        "maxColumn": 8,
    },
    "sales-table-validated.xlsx": {
        "sheets": ["Sales Data"],
        "charts": 0,
        "merged": [],
        "freeze": "C3",
        "validations": 1,
        "tables": ["SalesRecords"],
        "columnWidths": {"AA": 26.0, "AB": 16.0},
        "maxRow": 181,
        "maxColumn": 28,
    },
    "charts-images.xlsx": {
        "sheets": ["Dashboard", "Revenue Chart"],
        "charts": 1,
        "images": 1,
        "merged": ["A7:C8"],
        "freeze": "C3",
        "validations": 1,
        "conditionalFormats": 3,
        "maxRow": 8,
        "maxColumn": 4,
    },
    "large-grid.xlsx": {
        "sheets": ["Large Grid"],
        "charts": 0,
        "merged": [],
        "freeze": None,
        "validations": 0,
        "maxRow": 551,
        "maxColumn": 60,
    },
    "xlsx-chart-families.xlsx": {
        "sheets": ["Data", "Families", "Charts"],
        "charts": 12,
        "merged": [],
        "freeze": None,
        "validations": 0,
        "maxRow": 5,
        "maxColumn": 3,
    },
    "xlsx-negative-databar-icons.xlsx": {
        "sheets": ["Negative Data Bars", "Icon Sets"],
        "charts": 0,
        "merged": [],
        "freeze": None,
        "validations": 0,
        "conditionalFormats": 1,
        "maxRow": 8,
        "maxColumn": 2,
    },
    "xlsx-protected-1904.xlsx": {
        "sheets": ["Dates", "Protected"],
        "charts": 0,
        "merged": [],
        "freeze": None,
        "validations": 0,
        "maxRow": 3,
        "maxColumn": 2,
    },
}

for name in valid_xlsx:
    path = SAMPLES / name
    wb = load_workbook(path, data_only=False, read_only=False)
    sheet = wb.worksheets[0] if name != "charts-images.xlsx" else wb["Dashboard"]
    expected = xlsx_expectations[name]
    validations = sum(len(ws.data_validations.dataValidation) for ws in wb.worksheets)
    charts = sum(len(getattr(ws, "_charts", [])) for ws in wb.worksheets)
    images = sum(len(getattr(ws, "_images", [])) for ws in wb.worksheets)
    conditional_formats = len(list(sheet.conditional_formatting._cf_rules))
    report["xlsx"][name] = {
        "sheets": wb.sheetnames,
        "bytes": path.stat().st_size,
        "freeze": sheet.freeze_panes,
        "merged": [str(range_) for range_ in sheet.merged_cells.ranges],
        "validations": validations,
        "charts": charts,
        "images": images,
        "conditionalFormats": conditional_formats,
        "maxRow": sheet.max_row,
        "maxColumn": sheet.max_column,
    }
    expect(wb.sheetnames == expected["sheets"], f"{name}: expected sheets {expected['sheets']}, got {wb.sheetnames}")
    expect(len(wb.sheetnames) > 0, f"{name}: workbook must contain at least one sheet")
    expect(
        [str(range_) for range_ in sheet.merged_cells.ranges] == expected["merged"],
        f"{name}: expected merged regions {expected['merged']}, got {[str(range_) for range_ in sheet.merged_cells.ranges]}",
    )
    expect(
        sheet.freeze_panes == expected["freeze"],
        f"{name}: expected freeze panes {expected['freeze']}, got {sheet.freeze_panes}",
    )
    expect(
        validations == expected["validations"],
        f"{name}: expected {expected['validations']} data validation rules, got {validations}",
    )
    expect(
        charts == expected["charts"],
        f"{name}: expected {expected['charts']} worksheet charts, got {charts}",
    )
    expect(
        images == expected.get("images", 0),
        f"{name}: expected {expected.get('images', 0)} embedded images, got {images} (Pillow is required to read them)",
    )
    expect(
        conditional_formats == expected.get("conditionalFormats", 0),
        f"{name}: expected {expected.get('conditionalFormats', 0)} conditional-format ranges, got {conditional_formats}",
    )
    expect(
        sheet.max_row == expected.get("maxRow", sheet.max_row),
        f"{name}: expected {expected.get('maxRow')} rows, got {sheet.max_row}",
    )
    expect(
        sheet.max_column == expected.get("maxColumn", sheet.max_column),
        f"{name}: expected {expected.get('maxColumn')} columns, got {sheet.max_column}",
    )
    for column, width in expected.get("columnWidths", {}).items():
        actual = sheet.column_dimensions[column].width
        expect(actual == width, f"{name}: expected column {column} width {width}, got {actual}")
    for table_name in expected.get("tables", []):
        expect(
            table_name in sheet.tables,
            f"{name}: table {table_name} is missing from {list(sheet.tables)}",
        )
    wb.close()


def xlsx_part(name: str, part: str) -> str:
    with zipfile.ZipFile(SAMPLES / name) as archive:
        try:
            return archive.read(part).decode("utf-8")
        except KeyError:
            return ""


# Chart families are read from the chart parts, not from openpyxl's count: a
# regeneration that replaces a chart with a picture still reports one chart but
# loses the family the browser must render.
chart_families = "xlsx-chart-families.xlsx"
with zipfile.ZipFile(SAMPLES / chart_families) as archive:
    chart_parts = sorted(
        name for name in archive.namelist() if re.fullmatch(r"xl/charts/chart\d+\.xml", name)
    )
    chart_xml = "".join(archive.read(name).decode("utf-8") for name in chart_parts)
chart_report = {
    "chartParts": len(chart_parts),
    "plotFamilies": sorted(
        set(
            re.findall(
                r"<(?:c:)?(barChart|bar3DChart|lineChart|pieChart|doughnutChart|scatterChart|radarChart|areaChart)\b",
                chart_xml,
            )
        )
    ),
    "twoAxisCombos": sum(
        1
        for name in chart_parts
        if "barChart" in (part := xlsx_part(chart_families, name))
        and "lineChart" in part
        and part.count("<valAx>") >= 2
    ),
    "oneAxisCombos": sum(
        1
        for name in chart_parts
        if "barChart" in (part := xlsx_part(chart_families, name))
        and "lineChart" in part
        and part.count("<valAx>") == 1
    ),
}
report["xlsx"][f"{chart_families}#features"] = chart_report
expect(len(chart_parts) == 12, f"{chart_families}: expected 12 chart parts, got {len(chart_parts)}")
for family in [
    "barChart",
    "bar3DChart",
    "lineChart",
    "pieChart",
    "doughnutChart",
    "scatterChart",
    "radarChart",
    "areaChart",
]:
    expect(family in chart_xml, f"{chart_families}: missing {family} family")
expect(chart_report["oneAxisCombos"] == 1, f"{chart_families}: expected one single-axis combo chart")
expect(chart_report["twoAxisCombos"] == 1, f"{chart_families}: expected one secondary-axis combo chart")

bars = "xlsx-negative-databar-icons.xlsx"
bars_sheet = xlsx_part(bars, "xl/worksheets/sheet1.xml")
bars_report = {
    "dataBars": bars_sheet.count('type="dataBar"'),
    "negativeCells": len(re.findall(r"<v>-\d+</v>", bars_sheet)),
    "numCfvo": bars_sheet.count('cfvo type="num"'),
    "percentileCfvo": bars_sheet.count('cfvo type="percentile"'),
    "formulaCfvo": bars_sheet.count('type="expression"'),
    "iconSets": xlsx_part(bars, "xl/worksheets/sheet2.xml").count("iconSet="),
    "ratingIcons": xlsx_part(bars, "xl/worksheets/sheet2.xml").count('iconSet="5Rating"'),
    "arrowIcons": xlsx_part(bars, "xl/worksheets/sheet2.xml").count('iconSet="4Arrows"'),
}
report["xlsx"][f"{bars}#features"] = bars_report
expect(bars_report["dataBars"] == 1, f"{bars}: expected one data-bar rule")
expect(bars_report["negativeCells"] >= 3, f"{bars}: data-bar range must contain negative values")
expect(bars_report["numCfvo"] >= 2, f"{bars}: data-bar rule must use numeric cfvo bounds")
expect(bars_report["percentileCfvo"] == 3, f"{bars}: expected a percentile colour scale")
expect(bars_report["formulaCfvo"] == 1, f"{bars}: expected a formula conditional-format rule")
expect(bars_report["ratingIcons"] == 1, f"{bars}: expected the 5-icon rating set")
expect(bars_report["arrowIcons"] == 1, f"{bars}: expected the 4-icon arrow set")

protected = "xlsx-protected-1904.xlsx"
protected_workbook = xlsx_part(protected, "xl/workbook.xml")
protected_sheet = xlsx_part(protected, "xl/worksheets/sheet2.xml")
protected_book = load_workbook(SAMPLES / protected, data_only=False)
protected_report = {
    "date1904": 'date1904="1"' in protected_workbook,
    "workbookPassword": 'workbookPassword="9721"' in protected_workbook,
    "lockStructure": 'lockStructure="1"' in protected_workbook,
    "sheetPassword": 'password="9721"' in protected_sheet,
    "sheetProtected": 'sheet="1"' in protected_sheet,
    "formatCellsLocked": 'formatCells="0"' in protected_sheet,
    "epoch": str(protected_book.epoch),
}
report["xlsx"][f"{protected}#features"] = protected_report
expect(protected_report["date1904"], f"{protected}: workbook must use the 1904 date system")
expect(protected_report["workbookPassword"], f"{protected}: workbook structure protection password changed")
expect(protected_report["lockStructure"], f"{protected}: workbook structure must be locked")
expect(protected_report["sheetPassword"], f"{protected}: worksheet protection password changed")
expect(protected_report["sheetProtected"], f"{protected}: worksheet protection flag is missing")
expect(protected_report["formatCellsLocked"], f"{protected}: worksheet must keep formatCells disabled")
expect("1904-01-01" in protected_report["epoch"], f"{protected}: openpyxl did not read the 1904 epoch")
protected_book.close()

macro = "xlsx-macros.xlsm"
macro_archive = zipfile.ZipFile(SAMPLES / macro)
macro_names = macro_archive.namelist()
macro_types = macro_archive.read("[Content_Types].xml").decode("utf-8")
macro_rels = macro_archive.read("xl/_rels/workbook.xml.rels").decode("utf-8")
macro_bytes = macro_archive.read("xl/vbaProject.bin")
macro_report = {
    "vbaPart": "xl/vbaProject.bin" in macro_names,
    "macroContentType": "application/vnd.ms-excel.sheet.macroEnabled.main+xml" in macro_types,
    "vbaContentType": "application/vnd.ms-office.vbaProject" in macro_types,
    "vbaRelationship": "vbaProject.bin" in macro_rels,
    "cfbSignature": macro_bytes[:8].hex() == "d0cf11e0a1b11ae1",
    "bytes": len(macro_bytes),
}
macro_archive.close()
report["xlsx"][f"{macro}#features"] = macro_report
expect(macro_report["vbaPart"], f"{macro}: vbaProject.bin is missing")
expect(macro_report["macroContentType"], f"{macro}: workbook content type is not macro-enabled")
expect(macro_report["vbaContentType"], f"{macro}: vbaProject content type is missing")
expect(macro_report["vbaRelationship"], f"{macro}: vbaProject relationship is missing")
expect(macro_report["cfbSignature"], f"{macro}: vbaProject.bin is not an OLE compound file")
macro_book = load_workbook(SAMPLES / macro, keep_vba=True)
report["xlsx"][f"{macro}#vbaArchive"] = macro_book.vba_archive.namelist()
expect("xl/vbaProject.bin" in macro_book.vba_archive.namelist(), f"{macro}: openpyxl lost the VBA archive")
macro_book.close()

large = "xlsx-large-scale.xlsx"
large_book = load_workbook(SAMPLES / large, read_only=True)
large_report = {
    "sheets": large_book.sheetnames,
    "rowRange": large_book["Rows100k"].calculate_dimension(force=True),
    "columnRange": large_book["Columns10k"].calculate_dimension(force=True),
    "bytes": (SAMPLES / large).stat().st_size,
}
report["xlsx"][f"{large}#features"] = large_report
expect(large_report["sheets"] == ["Rows100k", "Columns10k"], f"{large}: expected the row and column sheets")
expect(large_report["rowRange"] == "A1:B100001", f"{large}: expected 100k data rows, got {large_report['rowRange']}")
expect(large_report["columnRange"] == "A1:NTP2", f"{large}: expected a 10k-column range, got {large_report['columnRange']}")
rows = large_book["Rows100k"].iter_rows(min_row=100001, max_row=100001, values_only=True)
expect(next(rows)[:2] == (100000, 200000), f"{large}: the 100,000th row lost its values")
large_book.close()
expect(large_report["bytes"] <= 4 * 1024 * 1024, f"{large}: fixture must stay below 4 MiB, got {large_report['bytes']}")

# PPTX parts are checked without a renderer: the fixture exists to pin the
# package structure that the preview/playback parser depends on.
for name in valid_pptx:
    with zipfile.ZipFile(SAMPLES / name) as archive:
        names = set(archive.namelist())
        master_xml = archive.read("ppt/slideMasters/slideMaster1.xml").decode("utf-8")
        slide2_xml = archive.read("ppt/slides/slide2.xml").decode("utf-8")
        slide3_xml = archive.read("ppt/slides/slide3.xml").decode("utf-8")
        layout_xml = "".join(
            archive.read(part).decode("utf-8")
            for part in sorted(names)
            if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", part)
        )
    pptx_report = {
        "masters": len([n for n in names if re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", n)]),
        "layouts": len([n for n in names if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", n)]),
        "slides": len([n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]),
        "placeholders": master_xml.count("<p:ph ") + layout_xml.count("<p:ph ") + slide2_xml.count("<p:ph ") + slide3_xml.count("<p:ph "),
        "smartArt": slide2_xml.count("<dgm:relIds"),
        "smartArtDrawing": "ppt/diagrams/drawing1.xml" in names,
        "video": slide3_xml.count("<a:videoFile") + slide3_xml.count("<p14:media"),
        "cjkRuns": slide2_xml.count("zh-CN"),
        "eaFont": slide2_xml.count('typeface="Microsoft YaHei"'),
    }
    report["pptx"][name] = pptx_report
    expect(pptx_report["masters"] == 1, f"{name}: expected one slide master")
    expect(pptx_report["layouts"] == 3, f"{name}: expected three slide layouts")
    expect(pptx_report["slides"] == 3, f"{name}: expected three slides")
    expect(pptx_report["placeholders"] >= 7, f"{name}: expected master/layout/slide placeholders")
    expect(pptx_report["smartArt"] == 1, f"{name}: expected a SmartArt relationship block")
    expect(pptx_report["smartArtDrawing"], f"{name}: SmartArt drawing fallback is missing")
    expect(pptx_report["video"] == 2, f"{name}: expected videoFile and p14:media markers")
    expect("ppt/media/video1.mp4" in names, f"{name}: video media part is missing")
    expect(pptx_report["cjkRuns"] >= 2, f"{name}: expected CJK language runs")
    expect(pptx_report["eaFont"] >= 2, f"{name}: expected East Asian font declarations")

for name in valid_png:
    path = SAMPLES / name
    with Image.open(path) as img:
        report["png"][name] = {
            "size": img.size,
            "mode": img.mode,
            "bytes": path.stat().st_size,
        }
        expect(img.size[0] > 0 and img.size[1] > 0, f"{name}: image must have a positive size")

for name in valid_json:
    path = SAMPLES / name
    data = json.loads(path.read_text(encoding="utf-8"))
    report["json"][name] = {
        "type": type(data).__name__,
        "bytes": path.stat().st_size,
    }

for name in corrupted:
    path = SAMPLES / name
    is_zip = zipfile.is_zipfile(path)
    report["corrupted"][name] = {
        "is_zip": is_zip,
        "bytes": path.stat().st_size,
    }
    expect(not is_zip, f"{name}: negative corrupted fixture must not be a zip")

report["result"] = "FAIL" if failures else "PASS"
print(json.dumps(report, ensure_ascii=False, indent=2))
if failures:
    raise SystemExit("material verification failed:\n" + "\n".join(f"- {failure}" for failure in failures))
