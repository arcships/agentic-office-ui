from __future__ import annotations

import argparse
import json
import random
import re
import zipfile
from pathlib import Path
from textwrap import wrap
from typing import Callable

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Mm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.drawing.image import Image as XlsxImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.comments import Comment
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, IconSetRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table as XlsxTable, TableStyleInfo
from PIL import Image, ImageDraw, ImageFont
from reportlab import rl_config
from reportlab.lib import colors, pdfencrypt
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Image as PdfImage
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "apps" / "demo" / "public" / "samples"
SAMPLES.mkdir(parents=True, exist_ok=True)
random.seed(42)
rl_config.invariant = 1

FIXED_OOXML_TIMESTAMP = "2020-01-01T00:00:00Z"
FIXED_ZIP_TIMESTAMP = (2020, 1, 1, 0, 0, 0)


def normalize_ooxml_zip(path: Path):
    """Make OOXML bytes reproducible across repeated local and CI runs."""
    temporary = path.with_suffix(f"{path.suffix}.deterministic")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(
        temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
    ) as target:
        for original in sorted(source.infolist(), key=lambda item: item.filename):
            data = source.read(original.filename)
            if original.filename == "docProps/core.xml":
                text = data.decode("utf-8")
                text = re.sub(
                    r"(<dcterms:(?:created|modified)[^>]*>).*?(</dcterms:(?:created|modified)>)",
                    rf"\g<1>{FIXED_OOXML_TIMESTAMP}\g<2>",
                    text,
                )
                data = text.encode("utf-8")
            entry = zipfile.ZipInfo(original.filename, FIXED_ZIP_TIMESTAMP)
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.create_system = 3
            entry.external_attr = (0o600 & 0xFFFF) << 16
            target.writestr(entry, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    temporary.replace(path)


def font_path() -> str | None:
    candidates = [
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
    ]
    for p in candidates:
        if p.exists():
            return str(p)
    return None


def pil_font(size: int):
    fp = font_path()
    if fp:
        try:
            return ImageFont.truetype(fp, size)
        except Exception:
            pass
    return ImageFont.load_default()


def set_doc_defaults(doc: Document):
    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(10.5)
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)


def add_header_footer(doc: Document, header: str):
    for section in doc.sections:
        section.header.paragraphs[0].text = header
        section.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        section.footer.paragraphs[0].text = "Confidential • Agentic Office UI browser verification"
        section.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER


def make_docx_legal_contract(path: Path):
    doc = Document()
    set_doc_defaults(doc)
    add_header_footer(doc, "Master Services Agreement — Browser Verification Fixture")
    doc.add_heading("MASTER SERVICES AGREEMENT", 0)
    intro = doc.add_paragraph()
    intro.add_run("This Master Services Agreement ").bold = True
    intro.add_run("is entered into by and between Extend Labs Ltd. and Acme Manufacturing Group for a production-grade document rendering validation scenario.")

    clauses = [
        ("1. Scope of Services", "Provider will deliver document intelligence, extraction review, workflow automation, quality assurance dashboards, and support services for the customer’s finance and operations teams."),
        ("2. Service Levels", "Provider will maintain commercially reasonable availability targets, response windows, and escalation paths. Severity one incidents require acknowledgement within one business hour."),
        ("3. Data Processing", "Each party will process personal information in accordance with applicable law. Confidential data includes contracts, invoices, purchase orders, and employee records."),
        ("4. Fees and Payment", "Customer will pay monthly subscription fees, implementation charges, and usage-based overages according to the order form. Late payments may accrue interest."),
        ("5. Audit Rights", "Customer may audit security controls once per calendar year with reasonable notice. Provider may redact unrelated customer information from audit evidence."),
        ("6. Limitation of Liability", "Neither party will be liable for indirect, incidental, punitive, or consequential damages, subject to exclusions for confidentiality and data protection obligations."),
        ("7. Term and Termination", "The initial term is twenty-four months. Either party may terminate for uncured material breach after thirty days written notice."),
        ("8. Governing Law", "This agreement is governed by the laws of the State of New York, excluding conflict-of-law principles."),
    ]
    for i in range(3):
        for title, body in clauses:
            doc.add_heading(title, level=1)
            p = doc.add_paragraph(body)
            p.paragraph_format.space_after = Pt(6)
            for j in range(1, 4):
                doc.add_paragraph(f"{title}.{j} Operational note {j}: {body} This sentence is repeated to create realistic multi-page pagination and line wrapping behavior.", style="List Number")
        if i < 2:
            doc.add_page_break()

    doc.add_heading("Authorized Signatures", level=1)
    table = doc.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "Extend Labs Ltd."
    table.cell(0, 1).text = "Acme Manufacturing Group"
    table.cell(1, 0).text = "Name: Jordan Lee\nTitle: VP Operations\nDate: 2026-07-05"
    table.cell(1, 1).text = "Name: Morgan Chen\nTitle: CFO\nDate: 2026-07-05"
    table.cell(2, 0).text = "Signature: ____________________"
    table.cell(2, 1).text = "Signature: ____________________"
    doc.save(path)


def make_docx_invoice_table(path: Path):
    doc = Document()
    set_doc_defaults(doc)
    add_header_footer(doc, "Invoice Table Fixture")
    doc.add_heading("INVOICE INV-2026-0705", 0)
    doc.add_paragraph("Vendor: Northwind Industrial Supplies    Customer: Extend Labs Ltd.")
    doc.add_paragraph("Tax ID: US-98-7654321    Due Date: 2026-08-04")
    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    headers = ["Item", "Description", "Qty", "Unit Price", "Tax", "Line Total"]
    for cell, text in zip(table.rows[0].cells, headers):
        cell.text = text
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True
    rows = [
        ("A-100", "OCR document page processing", 1200, 0.18, 0.07),
        ("B-210", "Human-in-the-loop verification seats", 12, 89.00, 0.07),
        ("C-310", "Secure retention storage GB", 500, 0.12, 0.07),
        ("D-410", "Workflow automation connector", 3, 450.00, 0.07),
    ]
    subtotal = 0
    for item, desc, qty, price, tax in rows:
        line = qty * price
        subtotal += line
        cells = table.add_row().cells
        values = [item, desc, str(qty), f"${price:,.2f}", f"{tax:.0%}", f"${line:,.2f}"]
        for c, v in zip(cells, values):
            c.text = v
    merged = table.add_row().cells
    merged[0].merge(merged[4]).text = "Subtotal"
    merged[5].text = f"${subtotal:,.2f}"
    total = subtotal * 1.07
    merged2 = table.add_row().cells
    merged2[0].merge(merged2[4]).text = "Total Due"
    merged2[5].text = f"${total:,.2f}"
    doc.add_paragraph("Payment terms: Net 30. Please include invoice number on wire remittance.")
    doc.save(path)


def make_docx_report_with_image(path: Path, image_path: Path):
    doc = Document()
    set_doc_defaults(doc)
    add_header_footer(doc, "Quarterly Operations Report Fixture")
    doc.add_heading("Quarterly Operations Report", 0)
    p = doc.add_paragraph()
    p.add_run("Executive summary: ").bold = True
    p.add_run("Document automation throughput improved ").italic = True
    p.add_run("18.4% quarter over quarter").underline = True
    p.add_run(" while maintaining review accuracy above 99.2%.")
    r = doc.add_paragraph().add_run("Risk indicator: manual exception volume increased in the APAC queue.")
    r.font.color.rgb = RGBColor(180, 40, 40)
    doc.add_picture(str(image_path), width=Inches(5.2))
    doc.add_paragraph("Figure 1. Synthetic invoice image used to verify embedded image rendering.")
    doc.add_heading("Operational Highlights", level=1)
    for text in [
        "Average extraction latency fell from 4.2 seconds to 2.9 seconds.",
        "Reviewer acceptance rate stayed within the control band for three consecutive months.",
        "The top exception categories were missing purchase order, tax mismatch, and low-confidence vendor name.",
    ]:
        doc.add_paragraph(text, style="List Bullet")
    doc.save(path)


def make_docx_chinese_mixed(path: Path):
    doc = Document()
    set_doc_defaults(doc)
    add_header_footer(doc, "中文混排测试文档")
    doc.add_heading("中文、English、数字与标点混排测试", 0)
    paragraphs = [
        "本文件用于验证 Vue DOCX Viewer 在中文、English words、1234567890、全角标点，半角 punctuation，以及长句换行场景下的渲染效果。",
        "合同编号：CN-EXT-2026-0705；供应商：上海凌云科技有限公司；金额：人民币 1,234,567.89 元；税率：6%。",
        "Mixed sentence: AI document processing 在真实业务中通常包含 invoices、contracts、receipts、forms 与 handwritten signatures。",
    ]
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.add_heading("编号列表", level=1)
    for i, text in enumerate(["准备测试物料", "启动浏览器验证", "记录样式与功能缺陷"], 1):
        doc.add_paragraph(f"{i}. {text}")
    doc.save(path)


def make_docx_review_comments(path: Path):
    """Create a deterministic review fixture with one comment and one replacement."""
    doc = Document()
    set_doc_defaults(doc)
    add_header_footer(doc, "Document Review Fixture")
    doc.add_heading("Supplier Review Notes", 0)
    doc.add_paragraph("COMMENT_TARGET")
    doc.add_paragraph("REVISION_TARGET")
    doc.add_paragraph("Use the 修订 and 批注 switches to compare the reviewed document state.")
    doc.save(path)

    with zipfile.ZipFile(path, "r") as source:
        entries = {name: source.read(name) for name in source.namelist()}

    document_xml = entries["word/document.xml"].decode("utf-8")
    comment_paragraph = (
        '<w:p><w:commentRangeStart w:id="0"/>'
        '<w:r><w:t>Monthly operational report</w:t></w:r>'
        '<w:commentRangeEnd w:id="0"/>'
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        '<w:commentReference w:id="0"/></w:r></w:p>'
    )
    revision_paragraph = (
        '<w:p><w:r><w:t xml:space="preserve">Response time changed from </w:t></w:r>'
        '<w:del w:id="1" w:author="Alex Reviewer" w:date="2026-07-11T00:00:00Z">'
        '<w:r><w:delText>five business days</w:delText></w:r></w:del>'
        '<w:r><w:t xml:space="preserve"> to </w:t></w:r>'
        '<w:ins w:id="2" w:author="Morgan Editor" w:date="2026-07-11T00:05:00Z">'
        '<w:r><w:t>two business days</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>'
    )
    comment_paragraph = comment_paragraph.replace(
        "</w:p>",
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>',
    )
    revision_paragraph = revision_paragraph.replace(
        "</w:p>",
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r></w:p>',
    )
    document_xml, comment_count = re.subn(
        r"<w:p\b[^>]*>(?:(?!</w:p>).)*?<w:t>COMMENT_TARGET</w:t>(?:(?!</w:p>).)*?</w:p>",
        comment_paragraph,
        document_xml,
        count=1,
    )
    document_xml, revision_count = re.subn(
        r"<w:p\b[^>]*>(?:(?!</w:p>).)*?<w:t>REVISION_TARGET</w:t>(?:(?!</w:p>).)*?</w:p>",
        revision_paragraph,
        document_xml,
        count=1,
    )
    if comment_count != 1 or revision_count != 1:
        raise RuntimeError("Unable to inject DOCX review fixture markup")
    entries["word/document.xml"] = document_xml.encode("utf-8")

    comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:comment w:id="0" w:author="Priya Reviewer" w:initials="PR" '
        'w:date="2026-07-11T00:10:00Z"><w:p><w:r>'
        '<w:t>Please confirm this report includes the service-level summary.</w:t>'
        '</w:r></w:p></w:comment></w:comments>'
    )
    entries["word/comments.xml"] = comments_xml.encode("utf-8")
    entries["word/footnotes.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:footnote w:id="1"><w:p><w:r><w:t>脚注正文：服务级别摘要已经由审核人确认。</w:t></w:r></w:p></w:footnote>'
        '</w:footnotes>'
    ).encode("utf-8")
    entries["word/endnotes.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:endnote w:id="1"><w:p><w:r><w:t>尾注正文：响应时间变更将在下一版合同中生效。</w:t></w:r></w:p></w:endnote>'
        '</w:endnotes>'
    ).encode("utf-8")

    content_types = entries["[Content_Types].xml"].decode("utf-8")
    if "/word/comments.xml" not in content_types:
        content_types = content_types.replace(
            "</Types>",
            '<Override PartName="/word/comments.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
            "</Types>",
        )
    content_types = content_types.replace(
        "</Types>",
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        "</Types>",
    )
    entries["[Content_Types].xml"] = content_types.encode("utf-8")

    relationships = entries["word/_rels/document.xml.rels"].decode("utf-8")
    if "relationships/comments" not in relationships:
        relationships = relationships.replace(
            "</Relationships>",
            '<Relationship Id="rIdReviewComments" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
            'Target="comments.xml"/></Relationships>',
        )
    relationships = relationships.replace(
        "</Relationships>",
        '<Relationship Id="rIdReviewFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
        '<Relationship Id="rIdReviewEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>'
        "</Relationships>",
    )
    entries["word/_rels/document.xml.rels"] = relationships.encode("utf-8")

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as target:
        for name in sorted(entries):
            target.writestr(name, entries[name])


# ---------------------------------------------------------------------------
# DOCX feature fixtures (issue #53)
#
# A DOCX fixture can only catch a rendering defect when it actually contains the
# feature that regressed. The review fixture above deliberately keeps exactly
# one comment and one revision, so the fixtures below carry the harder cases:
# comment threads and replies, several comments on one paragraph, ranges that
# cross paragraphs, comments anchored in a table cell and in a header, a
# resolved thread, vertically merged and nested tables, repeating header rows,
# mixed portrait/landscape sections, two-column sections, a text box and a
# behind-text watermark.
#
# The assertions for these files live in `tests/unit/docx-fixture-fidelity.test.mjs`
# (document model) and `tests/component/docx-fixture-fidelity.test.mjs`
# (rendered pages).
# ---------------------------------------------------------------------------

WORD_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
WORD14_NS = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
WORD15_NS = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPE_COMMENTS = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
CONTENT_TYPE_COMMENTS_EXTENDED = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
CONTENT_TYPE_FOOTNOTES = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)
CONTENT_TYPE_ENDNOTES = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
)


def read_docx_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as source:
        return {name: source.read(name) for name in source.namelist()}


def write_docx_parts(path: Path, entries: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as target:
        for name in sorted(entries):
            target.writestr(name, entries[name])


def replace_docx_paragraph(document_xml: str, marker: str, replacement: str) -> str:
    """Replace the single paragraph whose text is exactly `marker`."""
    pattern = (
        r"<w:p\b[^>]*>(?:(?!</w:p>).)*?<w:t[^>]*>"
        + re.escape(marker)
        + r"</w:t>(?:(?!</w:p>).)*?</w:p>"
    )
    updated, count = re.subn(pattern, replacement, document_xml, count=1)
    if count != 1:
        raise RuntimeError(f"Unable to replace DOCX fixture paragraph {marker!r}")
    return updated


def add_content_type_overrides(content_types: str, overrides: dict[str, str]) -> str:
    for part_name, content_type in overrides.items():
        if part_name in content_types:
            continue
        content_types = content_types.replace(
            "</Types>",
            f'<Override PartName="{part_name}" ContentType="{content_type}"/></Types>',
        )
    return content_types


def add_relationships(relationships: str, entries: list[tuple[str, str, str]]) -> str:
    for rel_id, rel_type, target in entries:
        if f'Id="{rel_id}"' in relationships:
            continue
        relationships = relationships.replace(
            "</Relationships>",
            f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/></Relationships>',
        )
    return relationships


def escape_xml(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def comment_markup(comment_id: int, text: str) -> tuple[str, str]:
    """Return (range start/text/end fragment, reference run) for a comment id."""
    body = (
        f'<w:commentRangeStart w:id="{comment_id}"/>'
        f"<w:r><w:t>{escape_xml(text)}</w:t></w:r>"
        f'<w:commentRangeEnd w:id="{comment_id}"/>'
    )
    reference = (
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{comment_id}"/></w:r>'
    )
    return body, reference


def find_table_row(document_xml: str, marker: str) -> tuple[int, str]:
    """Return (row index, raw `<w:tr>` XML) for the row containing `marker`."""
    cursor = 0
    row_index = -1
    while True:
        start = document_xml.find("<w:tr>", cursor)
        alternate = document_xml.find("<w:tr ", cursor)
        if start == -1 or (alternate != -1 and alternate < start):
            start = alternate
        if start == -1:
            raise RuntimeError(f"Unable to locate DOCX fixture table row {marker!r}")
        end = document_xml.find("</w:tr>", start)
        if end == -1:
            raise RuntimeError(f"Unbalanced DOCX fixture table row {marker!r}")
        row_index += 1
        row_xml = document_xml[start : end + len("</w:tr>")]
        if marker in row_xml:
            return row_index, row_xml
        cursor = end + len("</w:tr>")


def inject_row_properties(row_xml: str, properties_xml: str) -> str:
    """Insert `properties_xml` into the row's `<w:trPr>` (creating it first)."""
    if "<w:trPr>" in row_xml:
        return row_xml.replace("<w:trPr>", f"<w:trPr>{properties_xml}", 1)
    if "<w:trPr/>" in row_xml:
        return row_xml.replace("<w:trPr/>", f"<w:trPr>{properties_xml}</w:trPr>", 1)
    open_tag_end = row_xml.find(">") + 1
    return f"{row_xml[:open_tag_end]}<w:trPr>{properties_xml}</w:trPr>{row_xml[open_tag_end:]}"


def build_comments_xml(comments: list[tuple[int, str, str, str, str]]) -> str:
    """Serialize `word/comments.xml`; tuples are (id, author, initials, paraId, text)."""
    body = []
    for comment_id, author, initials, para_id, text in comments:
        body.append(
            f'<w:comment w:id="{comment_id}" w:author="{author}" w:initials="{initials}" '
            'w:date="2026-07-11T00:10:00Z">'
            f'<w:p w14:paraId="{para_id}"><w:r><w:t>{escape_xml(text)}</w:t></w:r></w:p>'
            "</w:comment>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:comments {WORD_NS} {WORD14_NS}>{''.join(body)}</w:comments>"
    )


def build_comments_extended_xml(
    entries: list[tuple[str, str | None, bool]]
) -> str:
    """Serialize `word/commentsExtended.xml`; tuples are (paraId, parent, done)."""
    body = []
    for para_id, parent, done in entries:
        attributes = [f'w15:paraId="{para_id}"']
        if parent is not None:
            attributes.append(f'w15:paraIdParent="{parent}"')
        if done:
            attributes.append('w15:done="1"')
        body.append(f"<w15:commentEx {' '.join(attributes)}/>")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w15:commentsEx {WORD15_NS}>{''.join(body)}</w15:commentsEx>"
    )


def build_notes_xml(tag: str, notes: list[tuple[int, str]]) -> str:
    """Serialize `word/footnotes.xml` or `word/endnotes.xml`; ids are not labels."""
    body = "".join(
        f'<w:{tag[:-1]} w:id="{note_id}"><w:p><w:r><w:t>{escape_xml(text)}</w:t></w:r></w:p>'
        f"</w:{tag[:-1]}>"
        for note_id, text in notes
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:{tag} {WORD_NS}>{body}</w:{tag}>"
    )


def make_docx_review_threads(path: Path):
    """Comment threads, reply/resolve state, cross-paragraph ranges and table revisions.

    Covers the review cases the single-comment fixture cannot express:

    * comment 0 with replies 1 and 2 (`commentsExtended` paraIdParent), where
      reply 2 has no inline markup and must inherit its parent's anchor,
    * comments 3 and 4 on two runs of the same paragraph,
    * comment 5 whose range starts in one paragraph and ends three paragraphs
      later,
    * comment 6 anchored in a table cell,
    * comment 7 anchored in the page header,
    * comment 8 marked done with reply 9 (a resolved thread),
    * comment 10 on an ordinary paragraph,
    * a deleted table row and a `trPrChange` row-height revision,
    * two footnotes and two endnotes whose markers must match the notes list.
    """
    doc = Document()
    set_doc_defaults(doc)
    add_header_footer(doc, "HEADER_COMMENT_TARGET")
    doc.add_heading("Review Threads", 0)
    doc.add_paragraph("THREAD_PARAGRAPH")
    doc.add_paragraph("TWO_COMMENT_PARAGRAPH")
    doc.add_paragraph("CROSS_START_PARAGRAPH")
    doc.add_paragraph("CROSS_MIDDLE_PARAGRAPH")
    doc.add_paragraph("CROSS_END_PARAGRAPH")
    doc.add_paragraph("RESOLVED_PARAGRAPH")
    doc.add_paragraph("SIMPLE_PARAGRAPH")

    table = doc.add_table(rows=4, cols=3)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, ["Vendor", "Amount", "Status"]):
        cell.text = text
    table.cell(1, 0).text = "Northwind Industrial"
    table.cell(1, 1).text = "CELL_COMMENT"
    table.cell(1, 2).text = "Approved"
    table.cell(2, 0).text = "DELETED_ROW_TARGET"
    table.cell(2, 1).text = "880.00"
    table.cell(2, 2).text = "Rejected"
    table.cell(3, 0).text = "ROW_HEIGHT_TARGET"
    table.cell(3, 1).text = "2,400.00"
    table.cell(3, 2).text = "Approved"

    doc.add_paragraph("CLOSING_PARAGRAPH")
    doc.save(path)

    entries = read_docx_parts(path)
    document_xml = entries["word/document.xml"].decode("utf-8")

    thread_body, thread_ref0 = comment_markup(0, "Threaded comment anchor")
    thread_ref1 = (
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        '<w:commentReference w:id="1"/></w:r>'
    )
    document_xml = replace_docx_paragraph(
        document_xml,
        "THREAD_PARAGRAPH",
        f"<w:p>{thread_body}{thread_ref0}{thread_ref1}</w:p>",
    )

    left_body, left_ref = comment_markup(3, "LEFT_SEGMENT")
    right_body, right_ref = comment_markup(4, "RIGHT_SEGMENT")
    document_xml = replace_docx_paragraph(
        document_xml,
        "TWO_COMMENT_PARAGRAPH",
        f'<w:p>{left_body}{left_ref}<w:r><w:t xml:space="preserve"> and </w:t></w:r>'
        f"{right_body}{right_ref}</w:p>",
    )

    start_body = '<w:commentRangeStart w:id="5"/>'
    document_xml = replace_docx_paragraph(
        document_xml,
        "CROSS_START_PARAGRAPH",
        '<w:p><w:r><w:t xml:space="preserve">Cross paragraph range begins </w:t></w:r>'
        f"{start_body}<w:r><w:t>here</w:t></w:r></w:p>",
    )
    document_xml = replace_docx_paragraph(
        document_xml,
        "CROSS_MIDDLE_PARAGRAPH",
        '<w:p><w:r><w:t>the range continues through this middle paragraph</w:t></w:r></w:p>',
    )
    end_ref = (
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        '<w:commentReference w:id="5"/></w:r>'
    )
    document_xml = replace_docx_paragraph(
        document_xml,
        "CROSS_END_PARAGRAPH",
        '<w:p><w:r><w:t>and closes at this point</w:t></w:r>'
        f'<w:commentRangeEnd w:id="5"/>{end_ref}</w:p>',
    )

    resolved_body, resolved_ref8 = comment_markup(8, "Resolved thread anchor paragraph")
    resolved_ref9 = (
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        '<w:commentReference w:id="9"/></w:r>'
    )
    document_xml = replace_docx_paragraph(
        document_xml,
        "RESOLVED_PARAGRAPH",
        f"<w:p>{resolved_body}{resolved_ref8}{resolved_ref9}</w:p>",
    )

    simple_body, simple_ref = comment_markup(10, "simple comment anchor paragraph")
    footnote_marker = (
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
        '<w:footnoteReference w:id="1"/></w:r>'
    )
    document_xml = replace_docx_paragraph(
        document_xml,
        "SIMPLE_PARAGRAPH",
        f"<w:p>{simple_body}{footnote_marker}{simple_ref}"
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
        '<w:endnoteReference w:id="2"/></w:r></w:p>',
    )

    cell_body, cell_ref = comment_markup(6, "1,200.00")
    document_xml = replace_docx_paragraph(
        document_xml, "CELL_COMMENT", f"<w:p>{cell_body}{cell_ref}</w:p>"
    )

    document_xml = replace_docx_paragraph(
        document_xml,
        "CLOSING_PARAGRAPH",
        '<w:p><w:r><w:t>Closing paragraph with two note markers</w:t></w:r>'
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
        '<w:footnoteReference w:id="2"/></w:r>'
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
        '<w:endnoteReference w:id="1"/></w:r></w:p>',
    )

    # --- review-threads-tail ---
    _, deleted_row_xml = find_table_row(document_xml, "DELETED_ROW_TARGET")
    deleted_row_index, _ = find_table_row(document_xml, "DELETED_ROW_TARGET")
    height_row_index, height_row_xml = find_table_row(document_xml, "ROW_HEIGHT_TARGET")
    if deleted_row_index >= height_row_index:
        raise RuntimeError("review fixture rows are out of order")
    document_xml = document_xml.replace(
        deleted_row_xml,
        inject_row_properties(
            deleted_row_xml,
            '<w:del w:id="20" w:author="Alex Reviewer" w:date="2026-07-11T00:20:00Z"/>',
        ),
        1,
    )
    document_xml = document_xml.replace(
        height_row_xml,
        inject_row_properties(
            height_row_xml,
            '<w:trHeight w:val="720" w:hRule="atLeast"/>'
            '<w:trPrChange w:id="22" w:author="Morgan Editor" w:date="2026-07-11T00:25:00Z">'
            '<w:trPr><w:trHeight w:val="360" w:hRule="atLeast"/></w:trPr>'
            "</w:trPrChange>",
        ),
        1,
    )
    entries["word/document.xml"] = document_xml.encode("utf-8")

    entries["word/comments.xml"] = build_comments_xml(
        [
            (0, "Priya Reviewer", "PR", "1A2B3C01", "Root comment: confirm the service-level summary is current."),
            (1, "Alex Reviewer", "AR", "1A2B3C02", "Reply one: the summary was refreshed this morning."),
            (2, "Morgan Editor", "ME", "1A2B3C03", "Reply two: no reference run, so the anchor is inherited."),
            (3, "Priya Reviewer", "PR", "1A2B3C04", "Left half of the same paragraph."),
            (4, "Alex Reviewer", "AR", "1A2B3C05", "Right half of the same paragraph."),
            (5, "Morgan Editor", "ME", "1A2B3C06", "Range crosses the three paragraphs above."),
            (6, "Priya Reviewer", "PR", "1A2B3C07", "Cell comment: the amount must match the purchase order."),
            (7, "Alex Reviewer", "AR", "1A2B3C08", "Header comment: the running head is out of date."),
            (8, "Morgan Editor", "ME", "1A2B3C09", "Resolved root: the wording was accepted."),
            (9, "Priya Reviewer", "PR", "1A2B3C0A", "Resolved reply: thanks, closing the thread."),
            (10, "Alex Reviewer", "AR", "1A2B3C0B", "Ordinary paragraph comment."),
        ]
    ).encode("utf-8")
    entries["word/commentsExtended.xml"] = build_comments_extended_xml(
        [
            ("1A2B3C01", None, False),
            ("1A2B3C02", "1A2B3C01", False),
            ("1A2B3C03", "1A2B3C01", False),
            ("1A2B3C04", None, False),
            ("1A2B3C05", None, False),
            ("1A2B3C06", None, False),
            ("1A2B3C07", None, False),
            ("1A2B3C08", None, False),
            ("1A2B3C09", None, True),
            ("1A2B3C0A", "1A2B3C09", False),
            ("1A2B3C0B", None, False),
        ]
    ).encode("utf-8")
    entries["word/footnotes.xml"] = build_notes_xml(
        "footnotes",
        [
            (1, "Footnote one: the service-level summary was reviewed."),
            (2, "Footnote two: the invoice total includes tax."),
        ],
    ).encode("utf-8")
    entries["word/endnotes.xml"] = build_notes_xml(
        "endnotes",
        [
            (1, "Endnote one: response times change in the next contract."),
            (2, "Endnote two: the retention policy is under legal review."),
        ],
    ).encode("utf-8")

    header_xml = entries["word/header1.xml"].decode("utf-8")
    header_body, header_ref = comment_markup(7, "Review Threads Fixture")
    header_xml = replace_docx_paragraph(
        header_xml, "HEADER_COMMENT_TARGET", f"<w:p>{header_body}{header_ref}</w:p>"
    )
    entries["word/header1.xml"] = header_xml.encode("utf-8")

    content_types = entries["[Content_Types].xml"].decode("utf-8")
    entries["[Content_Types].xml"] = add_content_type_overrides(
        content_types,
        {
            "/word/comments.xml": CONTENT_TYPE_COMMENTS,
            "/word/commentsExtended.xml": CONTENT_TYPE_COMMENTS_EXTENDED,
            "/word/footnotes.xml": CONTENT_TYPE_FOOTNOTES,
            "/word/endnotes.xml": CONTENT_TYPE_ENDNOTES,
        },
    ).encode("utf-8")

    relationships = entries["word/_rels/document.xml.rels"].decode("utf-8")
    entries["word/_rels/document.xml.rels"] = add_relationships(
        relationships,
        [
            ("rIdReviewThreadComments", f"{REL_NS}/comments", "comments.xml"),
            (
                "rIdReviewThreadCommentsExtended",
                "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                "commentsExtended.xml",
            ),
            ("rIdReviewThreadFootnotes", f"{REL_NS}/footnotes", "footnotes.xml"),
            ("rIdReviewThreadEndnotes", f"{REL_NS}/endnotes", "endnotes.xml"),
        ],
    ).encode("utf-8")

    write_docx_parts(path, entries)

def make_docx_table_structures(path: Path):
    """Vertical merges, grid spans, nested tables, multi-paragraph cells, header rows.

    `invoice-table.docx` only merges cells horizontally, so a dropped `vMerge`
    continuation cell or a lost nested table is invisible to the suite. This
    fixture carries both, plus `w:tblHeader` rows that must repeat when the long
    table in the same file splits across pages.
    """
    doc = Document()
    set_doc_defaults(doc)
    add_header_footer(doc, "Table Structures Fixture")
    doc.add_heading("Table Structures", 0)
    doc.add_paragraph("The first table merges vertically and horizontally and nests a table.")

    table = doc.add_table(rows=5, cols=4)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, ["Region", "Q1", "Q2", "Total"]):
        cell.text = text

    # Vertical merge over the three body rows of column 0.
    vertical = table.cell(1, 0).merge(table.cell(3, 0))
    vertical.text = "North region"
    vertical.add_paragraph("(all quarters)")
    # Horizontal merge in row 1 spanning the two quarter columns.
    horizontal = table.cell(1, 1).merge(table.cell(1, 2))
    horizontal.text = "Q1+Q2 combined"
    table.cell(1, 3).text = "2,650"
    # Multi-paragraph cell in a normal body row.
    multi = table.cell(2, 1)
    multi.text = "1,200"
    multi.add_paragraph("(audited)")
    multi.add_paragraph("(signed off)")
    table.cell(2, 2).text = "1,450"
    table.cell(2, 3).text = "2,650"
    # Nested table inside the last body cell.
    nested = table.cell(3, 3)
    nested.text = "Nested:"
    nested_table = nested.add_table(rows=2, cols=2)
    nested_table.style = "Table Grid"
    for row_index, values in enumerate([["N1", "N2"], ["N3", "N4"]]):
        for column_index, value in enumerate(values):
            nested_table.cell(row_index, column_index).text = value
    table.cell(4, 0).text = "Total"
    table.cell(4, 1).text = "2,650"
    table.cell(4, 2).text = "2,650"
    table.cell(4, 3).text = "5,300"

    doc.add_paragraph("The second table repeats its header row on every page it spans.")
    long_table = doc.add_table(rows=46, cols=8)
    long_table.style = "Table Grid"
    for column_index in range(8):
        long_table.cell(0, column_index).text = f"Column {column_index + 1}"
    for row_index in range(1, 46):
        for column_index in range(8):
            long_table.cell(row_index, column_index).text = (
                f"R{row_index:02d}C{column_index + 1}"
            )
    doc.save(path)

    entries = read_docx_parts(path)
    document_xml = entries["word/document.xml"].decode("utf-8")

    header_row_index, header_row_xml = find_table_row(document_xml, "Column 8")
    if header_row_index < 0 or "Column 1" not in header_row_xml:
        raise RuntimeError("the long table's header row could not be located")
    document_xml = document_xml.replace(
        header_row_xml,
        inject_row_properties(
            header_row_xml, '<w:tblHeader/><w:cantSplit/>'
        ),
        1,
    )
    entries["word/document.xml"] = document_xml.encode("utf-8")

    # A running head can contain a table (a common letterhead layout), and the
    # viewer renders it through the same table host as the body. Without this a
    # header table that is dropped, flattened or rendered outside the header band
    # is invisible to the suite.
    header_xml = entries["word/header1.xml"].decode("utf-8")
    if "<w:tbl>" in header_xml:
        raise RuntimeError("table-structures header already contains a table")
    closing = header_xml.find("</w:p>")
    if closing == -1:
        raise RuntimeError("table-structures header paragraph is missing")
    insert_at = closing + len("</w:p>")
    header_xml = (
        header_xml[:insert_at] + build_header_table_xml() + header_xml[insert_at:]
    )
    entries["word/header1.xml"] = header_xml.encode("utf-8")
    write_docx_parts(path, entries)


def build_header_table_xml() -> str:
    """A two-row table for a page header, in the fixture's table style.

    Word keeps an empty paragraph after a header table so the header can still
    hold a caret; the fixture does the same so it looks like Word output.
    """

    def cell(text: str, *, bold: bool = False) -> str:
        run_properties = "<w:rPr><w:b/></w:rPr>" if bold else ""
        return (
            '<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>'
            f"<w:p><w:r>{run_properties}<w:t>{escape_xml(text)}</w:t></w:r></w:p></w:tc>"
        )

    borders = (
        "<w:tblBorders>"
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="1C3A5E"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="1C3A5E"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="1C3A5E"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="1C3A5E"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="1C3A5E"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="1C3A5E"/>'
        "</w:tblBorders>"
    )
    grid = (
        "<w:tblGrid>"
        '<w:gridCol w:w="2880"/><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/>'
        "</w:tblGrid>"
    )
    header_row = (
        "<w:tr><w:trPr><w:tblHeader/></w:trPr>"
        + cell("Register", bold=True)
        + cell("Period", bold=True)
        + cell("Owner", bold=True)
        + "</w:tr>"
    )
    value_row = (
        "<w:tr>"
        + cell("Table Register")
        + cell("FY2026")
        + cell("O. North")
        + "</w:tr>"
    )
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        + borders
        + '<w:tblLook w:val="04A0"/></w:tblPr>'
        + grid
        + header_row
        + value_row
        + "</w:tbl><w:p/>"
    )


def make_docx_sections_columns(path: Path, image_path: Path):
    """Portrait, landscape and two-column sections, a text box and a watermark.

    Every section carries its own header/footer part so the fixture can prove
    that a page renders the header of *its* section (and that the header repeats
    on every page of that section). The first section's header also holds a
    behind-text watermark image.
    """
    doc = Document()
    set_doc_defaults(doc)

    first = doc.sections[0]
    # Capture the portrait geometry before any section break: `add_section()`
    # moves the body-level sectPr to the new last section, so the object from
    # `doc.sections[0]` stops describing the first page.
    portrait_width, portrait_height = first.page_width, first.page_height
    first.header.is_linked_to_previous = False
    first.footer.is_linked_to_previous = False
    first.header.paragraphs[0].text = "Section One Header"
    first.footer.paragraphs[0].text = "Section One Footer"
    doc.add_heading("Sections and Columns", 0)
    for index in range(1, 25):
        doc.add_paragraph(
            f"Section one portrait body paragraph {index:02d}. "
            "This paragraph exists so the section spans more than one page and "
            "the header and watermark have to repeat on every page of it."
        )

    landscape = doc.add_section(WD_SECTION.NEW_PAGE)
    landscape.orientation = WD_ORIENT.LANDSCAPE
    landscape.page_width, landscape.page_height = (
        landscape.page_height,
        landscape.page_width,
    )
    landscape.header.is_linked_to_previous = False
    landscape.footer.is_linked_to_previous = False
    landscape.header.paragraphs[0].text = "Section Two Landscape Header"
    landscape.footer.paragraphs[0].text = "Section Two Footer"
    doc.add_paragraph("TEXTBOX_TARGET_PARAGRAPH")
    doc.add_paragraph(
        "Landscape body paragraph. Wide content such as a schedule table reads "
        "better on a landscape page."
    )

    columns = doc.add_section(WD_SECTION.NEW_PAGE)
    # `add_section` copies the previous section's properties, so the two-column
    # section has to be switched back to portrait explicitly.
    columns.orientation = WD_ORIENT.PORTRAIT
    columns.page_width, columns.page_height = portrait_width, portrait_height
    columns.header.is_linked_to_previous = False
    columns.footer.is_linked_to_previous = False
    columns.header.paragraphs[0].text = "Section Three Two Column Header"
    columns.footer.paragraphs[0].text = "Section Three Footer"
    for index in range(1, 13):
        doc.add_paragraph(
            f"Two column body paragraph {index:02d}. Text in this section flows "
            "through two balanced columns before the next page break."
        )
    doc.save(path)

    entries = read_docx_parts(path)
    document_xml = entries["word/document.xml"].decode("utf-8")
    document_xml = replace_docx_paragraph(
        document_xml,
        "TEXTBOX_TARGET_PARAGRAPH",
        '<w:p><w:r><w:pict>'
        '<v:shape id="FixtureTextBox" type="#_x0000_t202" '
        'style="position:absolute;margin-left:180pt;margin-top:120pt;width:300pt;height:90pt;'
        'z-index:251659264" fillcolor="#DCE6F1" strokecolor="#1C3A5E">'
        '<v:textbox><w:txbxContent>'
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>TEXT BOX: quarterly review pack</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>Second text box paragraph</w:t></w:r></w:p>'
        "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>",
    )
    # The last section's sectPr becomes two columns. Only the body-level sectPr
    # is patched, so the single-column sections above keep their own w:cols.
    body_sectpr_start = document_xml.rfind("<w:sectPr")
    body_sectpr_end = document_xml.find("</w:sectPr>", body_sectpr_start)
    body_sectpr = document_xml[body_sectpr_start:body_sectpr_end]
    patched_sectpr = body_sectpr.replace(
        '<w:cols w:space="720"/>', '<w:cols w:num="2" w:space="720"/>'
    )
    if patched_sectpr == body_sectpr:
        raise RuntimeError("Unable to switch the fixture's last section to two columns")
    document_xml = (
        document_xml[:body_sectpr_start]
        + patched_sectpr
        + document_xml[body_sectpr_end:]
    )
    entries["word/document.xml"] = document_xml.encode("utf-8")

    watermark_bytes = image_path.read_bytes()
    entries["word/media/watermark.png"] = watermark_bytes

    header_xml = entries["word/header1.xml"].decode("utf-8")
    header_xml = header_xml.replace(
        "</w:p>",
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" '
        'simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" '
        'layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>1800000</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>2400000</wp:posOffset></wp:positionV>'
        '<wp:extent cx="3450000" cy="1200000"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '<wp:docPr id="9001" name="Watermark"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="0" name="watermark.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdWatermark"/><a:stretch><a:fillRect/></a:stretch>'
        "</pic:blipFill>"
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3450000" cy="1200000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        "</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>",
        1,
    )
    entries["word/header1.xml"] = header_xml.encode("utf-8")

    header_rels_key = "word/_rels/header1.xml.rels"
    header_rels = (
        entries[header_rels_key].decode("utf-8")
        if header_rels_key in entries
        else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    )
    entries[header_rels_key] = add_relationships(
        header_rels,
        [("rIdWatermark", f"{REL_NS}/image", "media/watermark.png")],
    ).encode("utf-8")

    content_types = entries["[Content_Types].xml"].decode("utf-8")
    if 'Extension="png"' not in content_types:
        content_types = content_types.replace(
            "<Default Extension=\"rels\"",
            '<Default Extension="png" ContentType="image/png"/><Default Extension="rels"',
        )
    entries["[Content_Types].xml"] = content_types.encode("utf-8")
    write_docx_parts(path, entries)


def make_docx_page_sizes(path: Path):
    """A4 portrait, A4 landscape and US Letter sections in one document.

    The demo must derive paper geometry from each section's `w:pgSz` instead of
    passing Letter dimensions from the host, so this fixture deliberately mixes
    millimetre and inch page sizes and gives every section its own header and
    footer. Each section spans more than one page, so the parity page has to
    compare page geometry and repeated furniture for six or more pages, and the
    landscape section is the only one whose paper is wider than it is tall.
    """
    doc = Document()
    set_doc_defaults(doc)

    def configure(
        section,
        width,
        height,
        orientation,
        header,
        footer,
        heading,
        body,
        paragraph_count,
    ):
        section.page_width = width
        section.page_height = height
        section.orientation = orientation
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        section.header.paragraphs[0].text = header
        section.header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        section.footer.paragraphs[0].text = footer
        section.footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_heading(heading, 1)
        for index in range(1, paragraph_count + 1):
            doc.add_paragraph(
                f"{heading} paragraph {index:02d}. {body} "
                "This text repeats so the section has to span several pages and "
                "its header and footer have to repeat on every page of it."
            )

    a4_width, a4_height = Mm(210), Mm(297)
    letter_width, letter_height = Inches(8.5), Inches(11)

    portrait = doc.sections[0]
    configure(
        portrait,
        a4_width,
        a4_height,
        WD_ORIENT.PORTRAIT,
        "A4 Portrait Header",
        "A4 Portrait Footer",
        "A4 Portrait Section",
        "Portrait pages use the ISO A4 (210 x 297 mm) paper size.",
        28,
    )

    landscape = doc.add_section(WD_SECTION.NEW_PAGE)
    configure(
        landscape,
        a4_height,
        a4_width,
        WD_ORIENT.LANDSCAPE,
        "A4 Landscape Header",
        "A4 Landscape Footer",
        "A4 Landscape Section",
        "Landscape pages keep A4 paper but swap its width and height.",
        14,
    )

    letter = doc.add_section(WD_SECTION.NEW_PAGE)
    configure(
        letter,
        letter_width,
        letter_height,
        WD_ORIENT.PORTRAIT,
        "Letter Portrait Header",
        "Letter Portrait Footer",
        "US Letter Section",
        "The final section switches back to US Letter (8.5 x 11 in).",
        18,
    )

    doc.save(path)


def make_invoice_png(path: Path):
    img = Image.new("RGB", (1000, 1330), "white")
    d = ImageDraw.Draw(img)
    title_font = pil_font(44)
    h_font = pil_font(26)
    font = pil_font(22)
    small = pil_font(18)
    d.rectangle([40, 40, 960, 1290], outline=(40, 60, 80), width=3)
    d.text((70, 70), "INVOICE", fill=(20, 50, 80), font=title_font)
    d.text((70, 135), "Northwind Industrial Supplies", fill=(0, 0, 0), font=h_font)
    d.text((70, 178), "Tax ID: US-98-7654321", fill=(0, 0, 0), font=font)
    d.text((650, 135), "Invoice #: INV-2026-0705", fill=(0, 0, 0), font=font)
    d.text((650, 178), "Date: 2026-07-05", fill=(0, 0, 0), font=font)
    d.text((70, 250), "Bill To: Extend Labs Ltd.", fill=(0, 0, 0), font=h_font)
    y = 340
    d.rectangle([70, y, 930, y+45], fill=(230, 237, 245), outline=(80, 100, 120))
    for x, t in [(90, "Description"), (560, "Qty"), (660, "Unit"), (800, "Total")]:
        d.text((x, y+10), t, fill=(0, 0, 0), font=small)
    y += 45
    items = [("OCR page processing", "1200", "$0.18", "$216.00"), ("Review seats", "12", "$89.00", "$1,068.00"), ("Storage GB", "500", "$0.12", "$60.00"), ("Workflow connector", "3", "$450.00", "$1,350.00")]
    for row in items:
        d.rectangle([70, y, 930, y+55], outline=(180, 190, 200))
        for x, t in [(90, row[0]), (570, row[1]), (660, row[2]), (800, row[3])]:
            d.text((x, y+15), t, fill=(0, 0, 0), font=small)
        y += 55
    d.text((650, y+45), "Subtotal", fill=(0, 0, 0), font=font)
    d.text((800, y+45), "$2,694.00", fill=(0, 0, 0), font=font)
    d.text((650, y+90), "Tax", fill=(0, 0, 0), font=font)
    d.text((800, y+90), "$188.58", fill=(0, 0, 0), font=font)
    d.text((650, y+145), "Total Due", fill=(180, 40, 40), font=h_font)
    d.text((800, y+145), "$2,882.58", fill=(180, 40, 40), font=h_font)
    d.text((70, 1180), "Payment terms: Net 30. Bank transfer preferred.", fill=(60, 60, 60), font=small)
    img.save(path)


def make_contract_page_png(path: Path):
    img = Image.new("RGB", (1000, 1330), "white")
    d = ImageDraw.Draw(img)
    title_font = pil_font(38)
    h_font = pil_font(26)
    font = pil_font(20)
    d.text((90, 70), "MASTER SERVICES AGREEMENT", fill=(20, 20, 20), font=title_font)
    y = 150
    sections = [
        ("1. Scope of Services", "Provider will deliver document intelligence, extraction review, workflow automation, and QA dashboards."),
        ("2. Service Levels", "Provider will maintain availability targets, response windows, and escalation paths for severity incidents."),
        ("3. Fees and Payment", "Customer will pay monthly subscription fees, implementation charges, and usage-based overages."),
    ]
    for heading, body in sections:
        d.text((90, y), heading, fill=(25, 65, 120), font=h_font)
        y += 46
        for line in wrap(body * 2, 82):
            d.text((90, y), line, fill=(35, 35, 35), font=font)
            y += 31
        y += 24
    d.rectangle([90, 870, 910, 1060], outline=(80, 80, 80), width=2)
    d.text((120, 900), "Authorized Signatures", fill=(20, 20, 20), font=h_font)
    d.line([120, 980, 450, 980], fill=(20, 20, 20), width=2)
    d.line([560, 980, 890, 980], fill=(20, 20, 20), width=2)
    d.text((120, 995), "Provider", fill=(35, 35, 35), font=font)
    d.text((560, 995), "Customer", fill=(35, 35, 35), font=font)
    img.save(path)


def pdf_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CenterTitle", parent=styles["Title"], alignment=TA_CENTER, fontSize=24, leading=30, spaceAfter=16))
    return styles


def make_pdf_sample(path: Path):
    doc = SimpleDocTemplate(str(path), pagesize=A4, rightMargin=18*mm, leftMargin=18*mm, topMargin=16*mm, bottomMargin=16*mm)
    styles = pdf_styles()
    story = [Paragraph("Document Intelligence Verification Report", styles["CenterTitle"]), Paragraph("This text PDF contains searchable paragraphs, tables, and multiple pages for PdfViewer validation.", styles["BodyText"]), Spacer(1, 8)]
    data = [["Metric", "Jan", "Feb", "Mar"], ["Pages processed", "42,180", "47,920", "53,110"], ["Accuracy", "99.1%", "99.2%", "99.3%"], ["Exceptions", "813", "702", "655"]]
    t = Table(data, hAlign="LEFT")
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1C3A5E")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F6FA")])]))
    story += [t, Spacer(1, 12)]
    for page in range(1, 5):
        story.append(Paragraph(f"Section {page}: Search Keyword ALPHA-{page}", styles["Heading1"]))
        for i in range(8):
            story.append(Paragraph("Agentic Office UI PdfViewer should support search, zoom, page navigation, thumbnail rendering, and scrolling with a realistic text layer. " * 2, styles["BodyText"]))
        if page < 4:
            story.append(PageBreak())
    doc.build(story)


def make_pdf_scanned(path: Path, image_path: Path):
    """Image-only pages: the fixture must have no native text layer at all."""
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    image = ImageReader(str(image_path))
    for _ in range(3):
        c.drawImage(image, 18*mm, 20*mm, width=w-36*mm, height=h-40*mm, preserveAspectRatio=True, anchor="c")
        c.showPage()
    c.save()


def make_pdf_rotated(path: Path):
    """One upright page and one page whose /Rotate entry is 90 degrees.

    ReportLab stores a rotated page at its displayed page size, so the second
    page keeps an A4 portrait MediaBox with `/Rotate 90`. Viewers must apply
    the page rotation instead of guessing the orientation from the MediaBox.
    """
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setFont("Helvetica-Bold", 22)
    c.drawString(40, h - 80, "Portrait Page")
    c.setFont("Helvetica", 12)
    for i in range(20):
        c.drawString(40, h - 120 - i*22, f"Portrait content row {i+1} for rotation and viewport validation.")
    c.showPage()
    c.setPageSize(landscape(A4))
    c.setPageRotation(90)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(40, h - 80, "Rotated Page (/Rotate 90)")
    for i in range(20):
        c.drawString(40, h - 120 - i*22, f"Rotated content row {i+1} must be drawn through the page rotation.")
    c.showPage()
    c.save()


def make_pdf_encrypted(path: Path):
    """Password-protected PDF; the documented password is ``arcships``."""
    encryption = pdfencrypt.StandardEncryption(
        "arcships",
        ownerPassword="arcships-owner",
        canPrint=1,
        canModify=1,
        canCopy=1,
        canAnnotate=1,
        strength=128,
    )
    c = canvas.Canvas(str(path), pagesize=A4, encrypt=encryption)
    w, h = A4
    for index, heading in enumerate(("Encrypted Contract", "Payment Schedule"), start=1):
        c.setFont("Helvetica-Bold", 22)
        c.drawString(40, h - 80, f"{heading} {index}")
        c.setFont("Helvetica", 12)
        for row in range(10):
            c.drawString(40, h - 120 - row*22, f"Encrypted page {index} row {row+1}: password protected content.")
        c.showPage()
    c.save()


def make_pdf_cjk_text(path: Path):
    """Chinese text pages so PDF text extraction and search cover CJK glyphs."""
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    pages = [
        [
            "中文合同摘要 / CJK-ALPHA-1",
            "合同金额：人民币壹佰贰拾万元整",
            "付款条款：验收合格后三十日内付款",
            "合同双方：甲方与乙方共同签署",
            "合同附件：技术规格与交付清单",
        ],
        [
            "合同补充条款 / CJK-ALPHA-2",
            "合同变更须经双方书面确认",
            "付款节点：预付款、进度款与尾款",
            "合同争议由合同签订地法院管辖",
        ],
    ]
    for lines in pages:
        c.setFont("STSong-Light", 20)
        c.drawString(40, h - 80, lines[0])
        c.setFont("STSong-Light", 14)
        for row, line in enumerate(lines[1:], start=1):
            c.drawString(40, h - 120 - row*28, line)
        c.showPage()
    c.save()


def make_pdf_large(path: Path):
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    for page in range(1, 241):
        c.setFont("Helvetica-Bold", 18)
        c.drawString(40, h - 60, f"Large Contract Page {page}")
        c.setFont("Helvetica", 10)
        for i in range(34):
            c.drawString(40, h - 95 - i*18, f"Clause {page}.{i+1}: This long contract page validates thumbnail buffering, scroll performance, and page counter updates.")
        c.showPage()
    c.save()


def style_header(ws):
    fill = PatternFill("solid", fgColor="1C3A5E")
    font = Font(color="FFFFFF", bold=True)
    # Only style header cells that carry a label: fixtures such as the sales
    # table put far-right helper columns on row 1, and a blanket row loop would
    # paint every empty column in between.
    for cell in ws[1]:
        if cell.value in (None, ""):
            continue
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")


def make_financial_model(path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Assumptions"
    ws.append(["Driver", "FY2026", "FY2027", "FY2028"])
    drivers = [("Starting Customers", 120, 180, 255), ("ARPU", 420, 456, 492), ("Gross Margin", 0.72, 0.74, 0.76), ("Churn", 0.08, 0.07, 0.06)]
    for row in drivers:
        ws.append(row)
    style_header(ws)
    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws2 = wb.create_sheet("P&L")
    ws2.append(["Line Item", "FY2026", "FY2027", "FY2028"])
    rows = ["Revenue", "COGS", "Gross Profit", "Sales & Marketing", "R&D", "G&A", "Operating Income"]
    for r, label in enumerate(rows, 2):
        ws2.cell(r, 1, label)
    ws2["B2"] = "=Assumptions!B2*Assumptions!B3"
    ws2["C2"] = "=Assumptions!C2*Assumptions!C3"
    ws2["D2"] = "=Assumptions!D2*Assumptions!D3"
    for col in "BCD":
        ws2[f"{col}3"] = f"={col}2*(1-Assumptions!{col}4)"
        ws2[f"{col}4"] = f"={col}2-{col}3"
        ws2[f"{col}5"] = f"={col}2*22%"
        ws2[f"{col}6"] = f"={col}2*18%"
        ws2[f"{col}7"] = f"={col}2*11%"
        ws2[f"{col}8"] = f"={col}4-SUM({col}5:{col}7)"
    style_header(ws2)
    for row in ws2.iter_rows(min_row=2, min_col=2, max_col=4):
        for cell in row:
            cell.number_format = '$#,##0'
    chart = LineChart()
    chart.title = "Revenue Forecast"
    chart.y_axis.title = "USD"
    data = Reference(ws2, min_col=2, max_col=4, min_row=1, max_row=2)
    chart.add_data(data, titles_from_data=True)
    ws2.add_chart(chart, "F2")
    wb.create_sheet("Notes")["A1"] = "Financial model fixture with formulas and multiple sheets."
    wb.save(path)


def make_sales_table_legacy(path: Path):
    random.seed(42)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Data"
    headers = ["Region", "Rep", "Product", "Month", "Units", "Unit Price", "Revenue", "Notes"]
    ws.append(headers)
    regions = ["North", "South", "East", "West", "APAC", "EMEA"]
    products = ["Doc Extract", "PDF Viewer", "XLSX Engine", "Workflow Bot"]
    for i in range(1, 181):
        units = random.randint(5, 180)
        price = random.choice([49, 89, 129, 249])
        ws.append([random.choice(regions), f"Rep {i%17+1}", random.choice(products), f"2026-{(i%12)+1:02d}", units, price, f"=E{i+1}*F{i+1}", "Long note validating overflow and cell text truncation behavior in the grid viewer."])
    style_header(ws)
    ws.freeze_panes = "A2"
    widths = [14, 12, 18, 12, 10, 12, 14, 58]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    for cell in ws[1]:
        cell.border = Border(bottom=Side(style="medium", color="1C3A5E"))
    sales_table = XlsxTable(displayName="SalesRecords", ref=f"A1:H{ws.max_row}")
    sales_table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(sales_table)
    wb.save(path)


def make_sales_table_validated(path: Path):
    random.seed(4242)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Data"
    headers = ["Region", "Rep", "Product", "Month", "Units", "Unit Price", "Revenue", "Notes"]
    ws.append(headers)
    regions = ["North", "South", "East", "West", "APAC", "EMEA"]
    products = ["Doc Extract", "PDF Viewer", "XLSX Engine", "Workflow Bot"]
    for i in range(1, 181):
        units = random.randint(5, 180)
        price = random.choice([49, 89, 129, 249])
        ws.append([random.choice(regions), f"Rep {i%17+1}", random.choice(products), f"2026-{(i%12)+1:02d}", units, price, f"=E{i+1}*F{i+1}", "Long note validating overflow and cell text truncation behavior in the grid viewer."])
    # Status is an input column outside the table: its list validation is the
    # fixture contract for the data-validation UI, not part of SalesRecords.
    statuses = ["Active", "On Hold", "Closed"]
    ws["I1"] = "Status"
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 9, statuses[row % len(statuses)])
    validation = DataValidation(
        type="list",
        formula1='"Active,On Hold,Closed"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Invalid status",
        error="Pick Active, On Hold or Closed.",
    )
    ws.add_data_validation(validation)
    validation.add(f"I2:I{ws.max_row}")
    style_header(ws)
    # Freeze rows *and* columns so a real fixture exercises the four-pane
    # clipping path (A1:B2 stays pinned while both axes scroll).
    ws.freeze_panes = "C3"
    widths = [14, 12, 18, 12, 10, 12, 14, 58, 12]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    # Two-letter column labels with explicit widths prove the parser maps
    # column letters beyond "Z" to workbook indices.
    ws["AA1"] = "Region rollup"
    ws["AB1"] = "Owner"
    for cell in ws[1]:
        cell.border = Border(bottom=Side(style="medium", color="1C3A5E"))
    for name, width in {"AA": 26, "AB": 16}.items():
        ws.column_dimensions[name].width = width
    sales_table = XlsxTable(displayName="SalesRecords", ref=f"A1:H{ws.max_row}")
    sales_table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(sales_table)
    wb.save(path)


def make_charts_images(path: Path, image_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws.append(["Quarter", "Revenue", "Exceptions"])
    for row in [("Q1", 120000, 820), ("Q2", 145000, 690), ("Q3", 168000, 540), ("Q4", 210000, 410)]:
        ws.append(row)
    style_header(ws)
    ws.merge_cells("A7:C8")
    ws["A7"] = "Worker merged region — image and merge parity"
    ws["A7"].font = Font(bold=True, color="1C3A5E")
    ws["A7"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A2"].hyperlink = "#Dashboard!A7"
    ws["A2"].style = "Hyperlink"
    ws["B2"].comment = Comment("该季度收入已通过财务复核。", "财务审核")
    ws.conditional_formatting.add("B2:B5", ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add("C2:C5", DataBarRule(start_type="min", end_type="max", color="5B9BD5", showValue=True))
    ws["D1"] = "趋势"
    ws["D1"].font = Font(color="FFFFFF", bold=True)
    ws["D3"], ws["D4"], ws["D5"] = 3, 2, 1
    ws.conditional_formatting.add("D3:D5", IconSetRule("3TrafficLights1", "num", [1, 2, 3], showValue=True))
    # A list validation on the sparkline/icon column exercises the validation
    # overlay against a real workbook instead of a synthetic part.
    trend_validation = DataValidation(
        type="list",
        formula1='"1,2,3"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Invalid trend",
        error="Pick 1, 2 or 3.",
    )
    ws.add_data_validation(trend_validation)
    trend_validation.add("D2:D5")
    # Multi-axis freeze: both rows and columns stay pinned while scrolling.
    ws.freeze_panes = "C3"
    ws.row_dimensions[7].height = 28
    ws.row_dimensions[8].height = 28
    chart = BarChart()
    chart.title = "Revenue vs Quarter"
    chart.y_axis.title = "Revenue"
    data = Reference(ws, min_col=2, max_col=2, min_row=1, max_row=5)
    cats = Reference(ws, min_col=1, min_row=2, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "E2")
    chartsheet_chart = BarChart()
    chartsheet_chart.title = "Revenue by Quarter"
    chartsheet_chart.y_axis.title = "Revenue"
    chartsheet_chart.add_data(data, titles_from_data=True)
    chartsheet_chart.set_categories(cats)
    chart_sheet = wb.create_chartsheet("Revenue Chart")
    chart_sheet.add_chart(chartsheet_chart)
    img = XlsxImage(str(image_path))
    img.width = 240
    img.height = 320
    ws.add_image(img, "E18")
    wb.save(path)
    inject_xlsx_fidelity_parts(path)


def inject_xlsx_fidelity_parts(path: Path):
    with zipfile.ZipFile(path, "r") as source:
        entries = {name: source.read(name) for name in source.namelist()}

    sheet_path = "xl/worksheets/sheet1.xml"
    sheet_xml = entries[sheet_path].decode("utf-8")
    sparkline_xml = (
        '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">'
        '<x14:sparklineGroups xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
        'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
        '<x14:sparklineGroup type="line" markers="1"><x14:colorSeries rgb="FF2563EB"/>'
        '<x14:colorMarkers rgb="FFDC2626"/><x14:sparklines><x14:sparkline>'
        '<xm:f>Dashboard!B2:B5</xm:f><xm:sqref>D2</xm:sqref>'
        '</x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups>'
        '</ext></extLst>'
    )
    control_xml = (
        '<controls><control xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" shapeId="9001" name="发布前确认" r:id="rIdFidelityControl">'
        '<controlPr><xdr:anchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to></xdr:anchor></controlPr>'
        '</control></controls>'
    )
    sheet_xml = sheet_xml.replace("</worksheet>", sparkline_xml + control_xml + "</worksheet>")
    entries[sheet_path] = sheet_xml.encode("utf-8")

    rels_path = "xl/worksheets/_rels/sheet1.xml.rels"
    rels_xml = entries[rels_path].decode("utf-8")
    rels_xml = rels_xml.replace(
        "</Relationships>",
        '<Relationship Id="rIdFidelityControl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>'
        "</Relationships>",
    )
    entries[rels_path] = rels_xml.encode("utf-8")
    entries["xl/ctrlProps/ctrlProp1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" checked="1" fmlaLink="Dashboard!$A$10"/>'
    ).encode("utf-8")

    drawing_path = "xl/drawings/drawing1.xml"
    drawing_xml = entries[drawing_path].decode("utf-8")
    shape_xml = (
        '<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        '<xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9002" name="审批提醒" descr="只读形状展示"/><xdr:cNvSpPr/></xdr:nvSpPr>'
        '<xdr:spPr><a:solidFill><a:srgbClr val="DBEAFE"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></a:ln><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></xdr:spPr>'
        '<xdr:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="1200" b="1"/><a:t>审批提醒：请核对季度收入</a:t></a:r></a:p></xdr:txBody>'
        '</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
    )
    closing_tag = "</xdr:wsDr>" if "</xdr:wsDr>" in drawing_xml else "</wsDr>"
    drawing_xml = drawing_xml.replace(closing_tag, shape_xml + closing_tag)
    entries[drawing_path] = drawing_xml.encode("utf-8")

    content_types = entries["[Content_Types].xml"].decode("utf-8")
    content_types = content_types.replace(
        "</Types>",
        '<Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>'
        "</Types>",
    )
    entries["[Content_Types].xml"] = content_types.encode("utf-8")

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as target:
        for name in sorted(entries):
            target.writestr(name, entries[name])


def make_large_grid(path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Large Grid"
    headers = [f"Col {i}" for i in range(1, 61)]
    ws.append(headers)
    for r in range(1, 551):
        ws.append([f"R{r}C{c}" if c % 7 else r * c for c in range(1, 61)])
    style_header(ws)
    wb.save(path)


def make_json_fixtures():
    fields = {
        "file": "/samples/invoice.png",
        "fields": [
            {"id": "vendor", "label": "Vendor", "page": 1, "rect": [0.07, 0.10, 0.48, 0.15], "value": "Northwind Industrial Supplies", "confidence": 0.96},
            {"id": "tax_id", "label": "Tax ID", "page": 1, "rect": [0.07, 0.135, 0.33, 0.17], "value": "US-98-7654321", "confidence": 0.83},
            {"id": "invoice_no", "label": "Invoice #", "page": 1, "rect": [0.64, 0.10, 0.91, 0.14], "value": "INV-2026-0705", "confidence": 0.74},
            {"id": "total", "label": "Total Due", "page": 1, "rect": [0.64, 0.64, 0.92, 0.70], "value": "$2,882.58", "confidence": 0.42},
            {"id": "terms", "label": "Payment Terms", "page": 1, "rect": [0.07, 0.88, 0.58, 0.91], "value": "Net 30", "confidence": 0.91},
        ],
    }
    (SAMPLES / "field-citations.json").write_text(json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")

    layout = {
        "file": "/samples/contract-page.png",
        "output": {
            "width": 1000,
            "height": 1330,
            "blocks": [
                {"id": "title", "kind": "title", "bbox": [90, 65, 640, 55], "text": "MASTER SERVICES AGREEMENT", "confidence": 0.98},
                {"id": "h1", "kind": "header", "bbox": [90, 148, 360, 42], "text": "1. Scope of Services", "confidence": 0.93},
                {"id": "body1", "kind": "text", "bbox": [90, 195, 820, 130], "text": "Provider will deliver document intelligence and workflow automation.", "confidence": 0.89},
                {"id": "h2", "kind": "header", "bbox": [90, 360, 310, 42], "text": "2. Service Levels", "confidence": 0.91},
                {"id": "body2", "kind": "text", "bbox": [90, 410, 820, 130], "text": "Provider will maintain availability targets and escalation paths.", "confidence": 0.86},
                {"id": "table", "kind": "table", "bbox": [90, 870, 820, 190], "text": "Authorized Signatures", "confidence": 0.77},
                {"id": "footer", "kind": "footer", "bbox": [90, 1195, 650, 30], "text": "Confidential", "confidence": 0.64},
            ],
        },
    }
    (SAMPLES / "ocr-layout.json").write_text(json.dumps(layout, ensure_ascii=False, indent=2), encoding="utf-8")


def make_corrupted(path: Path, text: str):
    path.write_bytes((text * 20).encode("utf-8"))


def fixture_builders() -> dict[str, Callable[[], None]]:
    """Every generated fixture, keyed by file name, in generation order."""
    invoice_png = SAMPLES / "invoice.png"
    contract_png = SAMPLES / "contract-page.png"
    return {
        "invoice.png": lambda: make_invoice_png(invoice_png),
        "contract-page.png": lambda: make_contract_page_png(contract_png),
        "legal-contract.docx": lambda: make_docx_legal_contract(SAMPLES / "legal-contract.docx"),
        "invoice-table.docx": lambda: make_docx_invoice_table(SAMPLES / "invoice-table.docx"),
        "report-with-image.docx": lambda: make_docx_report_with_image(SAMPLES / "report-with-image.docx", invoice_png),
        "chinese-mixed.docx": lambda: make_docx_chinese_mixed(SAMPLES / "chinese-mixed.docx"),
        "review-comments.docx": lambda: make_docx_review_comments(SAMPLES / "review-comments.docx"),
        "review-threads.docx": lambda: make_docx_review_threads(SAMPLES / "review-threads.docx"),
        "table-structures.docx": lambda: make_docx_table_structures(SAMPLES / "table-structures.docx"),
        "sections-columns.docx": lambda: make_docx_sections_columns(SAMPLES / "sections-columns.docx", invoice_png),
        "page-sizes.docx": lambda: make_docx_page_sizes(SAMPLES / "page-sizes.docx"),
        "demo.docx": lambda: make_docx_legal_contract(SAMPLES / "demo.docx"),
        "corrupted.docx": lambda: make_corrupted(SAMPLES / "corrupted.docx", "This is not a valid DOCX package. "),
        "sample.pdf": lambda: make_pdf_sample(SAMPLES / "sample.pdf"),
        "scanned-invoice.pdf": lambda: make_pdf_scanned(SAMPLES / "scanned-invoice.pdf", invoice_png),
        "rotated-pages.pdf": lambda: make_pdf_rotated(SAMPLES / "rotated-pages.pdf"),
        "encrypted.pdf": lambda: make_pdf_encrypted(SAMPLES / "encrypted.pdf"),
        "cjk-report.pdf": lambda: make_pdf_cjk_text(SAMPLES / "cjk-report.pdf"),
        "large-contract.pdf": lambda: make_pdf_large(SAMPLES / "large-contract.pdf"),
        "corrupted.pdf": lambda: make_corrupted(SAMPLES / "corrupted.pdf", "%PDF corrupted fixture\n"),
        "financial-model.xlsx": lambda: make_financial_model(SAMPLES / "financial-model.xlsx"),
        "sales-table.xlsx": lambda: make_sales_table_legacy(SAMPLES / "sales-table.xlsx"),
        "sales-table-validated.xlsx": lambda: make_sales_table_validated(SAMPLES / "sales-table-validated.xlsx"),
        "charts-images.xlsx": lambda: make_charts_images(SAMPLES / "charts-images.xlsx", invoice_png),
        "large-grid.xlsx": lambda: make_large_grid(SAMPLES / "large-grid.xlsx"),
        "corrupted.xlsx": lambda: make_corrupted(SAMPLES / "corrupted.xlsx", "This is not a valid XLSX package. "),
        "field-citations.json": make_json_fixtures,
        "ocr-layout.json": make_json_fixtures,
    }


def main(argv: list[str] | None = None):
    """Generate every fixture, or only the ones named with `--only`.

    `--only review-threads.docx,table-structures.docx` regenerates a subset; the
    samples manifest is still rewritten for the whole directory so it always
    describes what is on disk. Use it when adding a fixture without rebuilding
    the unrelated binaries with a different toolchain version.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--only",
        default=None,
        help="comma-separated fixture file names to regenerate",
    )
    args = parser.parse_args(argv)
    builders = fixture_builders()
    selected = list(builders)
    if args.only:
        requested = [name.strip() for name in args.only.split(",") if name.strip()]
        unknown = [name for name in requested if name not in builders]
        if unknown:
            parser.error(f"unknown fixture(s): {', '.join(unknown)}")
        selected = requested

    for name in selected:
        builders[name]()

    for name in selected:
        path = SAMPLES / name
        if path.suffix in {".docx", ".xlsx"} and not name.startswith("corrupted."):
            normalize_ooxml_zip(path)

    manifest = []
    for path in sorted(SAMPLES.iterdir()):
        if path.is_file() and path.name != "manifest.json":
            manifest.append({"name": path.name, "bytes": path.stat().st_size})
    (SAMPLES / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(manifest, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
