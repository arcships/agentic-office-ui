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
]
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
