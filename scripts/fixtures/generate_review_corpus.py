"""Generate the review corpus used by the health-check findings.

Run with:  uv run --with python-docx --with openpyxl --with pypdf --with pillow --with python-pptx \
               python scripts/fixtures/generate_review_corpus.py

Output: test-data/corpus/ (binaries) + test-data/corpus/manifest.json
The corpus is intentionally independent from apps/demo/public/samples so the demo
manifest hashes stay stable.
"""

from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart,
    BarChart,
    LineChart,
    PieChart,
    RadarChart,
    Reference,
    ScatterChart,
    Series,
)
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, IconSetRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.worksheet.datavalidation import DataValidation
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches as PInches, Pt as PPt
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "test-data" / "corpus"
OUT.mkdir(parents=True, exist_ok=True)
SAMPLES = ROOT / "apps" / "demo" / "public" / "samples"

manifest: list[dict] = []


def record(name: str, fmt: str, features: list[str], note: str) -> None:
    path = OUT / name
    manifest.append(
        {
            "file": name,
            "format": fmt,
            "bytes": path.stat().st_size,
            "features": features,
            "note": note,
        }
    )
    print(f"  {name:28} {path.stat().st_size / 1024:8.1f} KiB  {', '.join(features)}")


# ── DOCX ──────────────────────────────────────────────────────────────────


def docx_merged_table() -> None:
    doc = Document()
    doc.add_heading("Merged table matrix", 1)
    table = doc.add_table(rows=5, cols=4)
    table.style = "Table Grid"
    # vertical merge: first column rows 1-3
    table.cell(1, 0).merge(table.cell(3, 0)).text = "vertical merge"
    # horizontal merge: header row
    table.cell(0, 0).merge(table.cell(0, 3)).text = "header spanning all columns"
    # multi-paragraph cell
    cell = table.cell(4, 1)
    cell.text = "first paragraph"
    cell.add_paragraph("second paragraph")
    cell.add_paragraph("third paragraph")
    # nested table inside a cell
    nested = table.cell(4, 2)
    sub = nested.add_table(rows=2, cols=2)
    sub.cell(0, 0).text = "nested a"
    sub.cell(0, 1).text = "nested b"
    sub.cell(1, 0).text = "nested c"
    sub.cell(1, 1).text = "nested d"
    for row in table.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
    doc.save(OUT / "docx-merged-table.docx")
    record(
        "docx-merged-table.docx",
        "docx",
        ["vmerge", "gridspan", "multi-paragraph-cell", "nested-table"],
        "vertical/horizontal merges, multi-paragraph cells, nested table",
    )


def docx_mixed_orientation() -> None:
    doc = Document()
    doc.add_heading("Portrait section", 1)
    doc.add_paragraph("Portrait page content. " * 40)
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    new_w, new_h = section.page_height, section.page_width
    section.page_width = new_w
    section.page_height = new_h
    doc.add_heading("Landscape section", 1)
    doc.add_paragraph("Landscape page content. " * 40)
    doc.save(OUT / "docx-mixed-orientation.docx")
    record(
        "docx-mixed-orientation.docx",
        "docx",
        ["sections", "landscape", "portrait"],
        "portrait section followed by a landscape section",
    )


def docx_two_columns() -> None:
    doc = Document()
    doc.add_heading("Two column section", 1)
    doc.add_paragraph("Column layout text. " * 60)
    sect_pr = doc.sections[0]._sectPr
    cols = sect_pr.xpath("./w:cols")
    if cols:
        cols[0].set(qn("w:num"), "2")
        cols[0].set(qn("w:space"), "720")
    doc.save(OUT / "docx-two-columns.docx")
    record(
        "docx-two-columns.docx",
        "docx",
        ["columns"],
        "two-column section",
    )


def docx_comments() -> None:
    doc = Document()
    doc.add_heading("Comment threads", 1)
    p1 = doc.add_paragraph("This paragraph carries two independent comments. ")
    r1 = p1.runs[0]
    doc.add_comment(runs=r1, text="First reviewer: please confirm the SLA.", author="Priya Reviewer", initials="PR")
    r2 = p1.add_run("The second sentence has its own reviewer note.")
    doc.add_comment(runs=r2, text="Second reviewer: numbers need a source.", author="Alex Reviewer", initials="AR")

    doc.add_paragraph("A table with a comment inside a cell:")
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "cell with comment"
    cell_para = table.cell(0, 0).paragraphs[0]
    doc.add_comment(
        runs=cell_para.runs[0],
        text="Cell comment: keep this row visible.",
        author="Morgan Editor",
        initials="ME",
    )
    table.cell(0, 1).text = "plain cell"

    section = doc.sections[0]
    header = section.header
    header.paragraphs[0].text = "Header text (header comments need a Word-authored fixture)"

    doc.add_paragraph("Footnote anchor follows.").add_run("\u2063")
    # add a real footnote reference via raw XML
    p = doc.paragraphs[-1]
    run = p.add_run()
    fn = OxmlElement("w:footnoteReference")
    fn.set(qn("w:id"), "1")
    run._r.append(fn)

    doc.save(OUT / "docx-comments.docx")

    # inject footnotes part if python-docx did not create one
    import zipfile
    from shutil import move

    src = OUT / "docx-comments.docx"
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        entries = {n: z.read(n) for n in names}
    if "word/footnotes.xml" not in entries:
        entries["word/footnotes.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:footnote w:id="1"><w:p><w:r><w:t>Footnote body for the comment fixture.</w:t></w:r></w:p></w:footnote>'
            "</w:footnotes>"
        ).encode()
    # Register the injected part in [Content_Types].xml and document.xml.rels;
    # without both, LibreOffice/Word reject the package as corrupt.
    content_types = entries["[Content_Types].xml"].decode()
    if "/word/footnotes.xml" not in content_types:
        content_types = content_types.replace(
            "</Types>",
            '<Override PartName="/word/footnotes.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
        )
        entries["[Content_Types].xml"] = content_types.encode()
    document_rels = entries["word/_rels/document.xml.rels"].decode()
    if "relationships/footnotes" not in document_rels:
        document_rels = document_rels.replace(
            "</Relationships>",
            '<Relationship Id="rIdFootnotes1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" '
            'Target="footnotes.xml"/></Relationships>',
        )
        entries["word/_rels/document.xml.rels"] = document_rels.encode()
    tmp = OUT / "docx-comments.tmp.docx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    move(tmp, src)

    record(
        "docx-comments.docx",
        "docx",
        ["comments", "multiple-comments", "table-comment", "header-comment", "footnote"],
        "two comments on one paragraph, table-cell comment, header comment, footnote",
    )


def docx_large_table() -> None:
    doc = Document()
    doc.add_heading("Large table", 1)
    table = doc.add_table(rows=120, cols=4)
    table.style = "Table Grid"
    for r in range(120):
        for c in range(4):
            table.cell(r, c).text = f"r{r + 1}c{c + 1}"
    doc.save(OUT / "docx-large-table.docx")
    record(
        "docx-large-table.docx",
        "docx",
        ["table", "pagination-stress"],
        "120-row table with a header row",
    )


# ── XLSX ──────────────────────────────────────────────────────────────────


def xlsx_charts() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Category", "Sales", "Cost", "Margin"])
    rows = [
        ("Jan", 120, 80, 40),
        ("Feb", 150, 95, 55),
        ("Mar", 170, 110, 60),
        ("Apr", 140, 90, 50),
        ("May", 200, 130, 70),
    ]
    for row in rows:
        ws.append(row)
    data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=6)
    cats = Reference(ws, min_col=1, min_row=2, max_row=6)

    bar = BarChart()
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.title = "Bar"
    ws.add_chart(bar, "F2")

    line = LineChart()
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    line.title = "Line"
    ws.add_chart(line, "F20")

    pie = PieChart()
    pie.add_data(Reference(ws, min_col=2, min_row=1, max_row=6), titles_from_data=True)
    pie.set_categories(cats)
    pie.title = "Pie"
    ws.add_chart(pie, "N2")

    area = AreaChart()
    area.add_data(data, titles_from_data=True)
    area.set_categories(cats)
    area.title = "Area"
    ws.add_chart(area, "N20")

    radar = RadarChart()
    radar.add_data(data, titles_from_data=True)
    radar.set_categories(cats)
    radar.title = "Radar"
    ws.add_chart(radar, "V2")

    scatter = ScatterChart()
    xs = Reference(ws, min_col=2, min_row=2, max_row=6)
    ys = Reference(ws, min_col=4, min_row=2, max_row=6)
    s = Series(ys, xs, title="Margin vs Sales")
    scatter.series.append(s)
    scatter.title = "Scatter"
    ws.add_chart(scatter, "V20")

    wb.save(OUT / "xlsx-charts.xlsx")
    record(
        "xlsx-charts.xlsx",
        "xlsx",
        ["bar", "line", "pie", "area", "radar", "scatter"],
        "six chart families on one sheet",
    )


def xlsx_conditional() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "CF"
    for i in range(1, 31):
        ws.cell(row=i, column=1, value=i * 2)
        ws.cell(row=i, column=2, value=(i % 7) - 3)
    ws.conditional_formatting.add(
        "A1:A30",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="63BE7B", showValue=True),
    )
    ws.conditional_formatting.add(
        "A1:A30",
        ColorScaleRule(
            start_type="min",
            start_color="F8696B",
            mid_type="percentile",
            mid_value=50,
            mid_color="FFEB84",
            end_type="max",
            end_color="63BE7B",
        ),
    )
    ws.conditional_formatting.add(
        "B1:B30",
        IconSetRule("4TrafficLights", "num", [0, 25, 50, 75]),
    )
    ws.conditional_formatting.add(
        "B1:B30",
        DataBarRule(start_type="num", start_value=-3, end_type="num", end_value=3, color="FF0000", showValue=True),
    )
    wb.save(OUT / "xlsx-conditional-formats.xlsx")
    record(
        "xlsx-conditional-formats.xlsx",
        "xlsx",
        ["data-bar", "color-scale", "icon-set", "cfvo-num", "cfvo-percentile", "negative-values"],
        "data bars with numeric bounds, colour scale with percentile mid, 4-icon set",
    )


def xlsx_validation() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation"
    ws["A1"] = "Fruit"
    ws["A2"] = "Apple"
    ws["A3"] = "Banana"
    ws["C1"] = "Pick"
    dv_list = DataValidation(type="list", formula1="=$A$2:$A$3", allow_blank=True, showDropDown=False)
    dv_list.error = "Pick a fruit from the list"
    dv_list.errorTitle = "Invalid fruit"
    ws.add_data_validation(dv_list)
    dv_list.add("C2:C10")

    dv_num = DataValidation(type="whole", operator="between", formula1="1", formula2="100", showErrorMessage=True)
    ws.add_data_validation(dv_num)
    dv_num.add("D2:D10")
    wb.save(OUT / "xlsx-validation.xlsx")
    record(
        "xlsx-validation.xlsx",
        "xlsx",
        ["data-validation-list", "data-validation-number"],
        "list and numeric validations with error messages",
    )


def xlsx_frozen_hidden_1904() -> None:
    wb = Workbook()
    wb.epoch = CALENDAR_MAC_1904
    ws = wb.active
    ws.title = "Frozen"
    ws["A1"] = "Merged title"
    ws.merge_cells("A1:C1")
    for r in range(1, 10):
        for c in range(1, 6):
            ws.cell(row=r + 1, column=c, value=f"r{r}c{c}")
    ws.freeze_panes = "B3"
    ws.row_dimensions[3].hidden = True
    ws.column_dimensions["D"].hidden = True
    ws["H1"] = "1904 date system"
    wb.save(OUT / "xlsx-frozen-hidden.xlsx")
    record(
        "xlsx-frozen-hidden.xlsx",
        "xlsx",
        ["freeze-rows-and-columns", "hidden-row", "hidden-column", "merge", "date1904"],
        "freeze B3, hidden row 3 and column D, merged A1:C1, 1904 epoch",
    )


def xlsx_large() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Large"
    ws.append(["id", "name", "qty", "price", "total"])
    for i in range(1, 20001):
        ws.append([i, f"item-{i}", i % 97, round((i % 1000) + 0.5, 2), f"=C{i + 1}*D{i + 1}"])
    wb.save(OUT / "xlsx-large-20k.xlsx")
    record(
        "xlsx-large-20k.xlsx",
        "xlsx",
        ["large", "formulas", "20000-rows"],
        "20k rows with formulas",
    )


def xlsx_number_formats() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Formats"
    ws["A1"] = "date"
    ws["B1"] = "currency"
    ws["C1"] = "percent"
    ws["D1"] = "text"
    ws["E1"] = "custom negative"
    import datetime as dt

    ws["A2"] = dt.datetime(2024, 1, 15)
    ws["A2"].number_format = "m/d/yyyy"
    ws["B2"] = 1234.56
    ws["B2"].number_format = '"$"#,##0.00'
    ws["C2"] = 0.256
    ws["C2"].number_format = "0.0%"
    ws["D2"] = "00123"
    ws["D2"].number_format = "@"
    ws["E2"] = -1234.5
    ws["E2"].number_format = "#,##0.00_);[Red](#,##0.00)"
    wb.save(OUT / "xlsx-number-formats.xlsx")
    record(
        "xlsx-number-formats.xlsx",
        "xlsx",
        ["date", "currency", "percent", "text", "custom-format"],
        "representative number formats including a [Red] section",
    )


# ── PDF ───────────────────────────────────────────────────────────────────


def pdf_rotated() -> None:
    reader = PdfReader(str(SAMPLES / "sample.pdf"))
    writer = PdfWriter()
    for page in reader.pages:
        page[NameObject("/Rotate")] = NumberObject(90)
        writer.add_page(page)
    with (OUT / "pdf-rotated-90.pdf").open("wb") as fh:
        writer.write(fh)
    record("pdf-rotated-90.pdf", "pdf", ["rotation-90"], "pages carry /Rotate 90")


def pdf_image_only() -> None:
    img = Image.new("RGB", (1240, 1754), "white")
    draw = ImageDraw.Draw(img)
    for i in range(20):
        draw.line((100, 150 + i * 60, 1100, 150 + i * 60), fill="black", width=8)
    img.save(OUT / "pdf-image-only.pdf", "PDF", resolution=150.0)
    record("pdf-image-only.pdf", "pdf", ["image-only", "scanned"], "no text layer at all")


def pdf_encrypted() -> None:
    reader = PdfReader(str(SAMPLES / "sample.pdf"))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt("secret")
    with (OUT / "pdf-encrypted.pdf").open("wb") as fh:
        writer.write(fh)
    record("pdf-encrypted.pdf", "pdf", ["encrypted", "password"], "password: secret")


def pdf_many_pages() -> None:
    reader = PdfReader(str(SAMPLES / "sample.pdf"))
    writer = PdfWriter()
    for _ in range(300):
        writer.add_page(reader.pages[0])
    with (OUT / "pdf-300-pages.pdf").open("wb") as fh:
        writer.write(fh)
    record("pdf-300-pages.pdf", "pdf", ["large", "300-pages"], "300 copies of page 1")


# ── PPTX ──────────────────────────────────────────────────────────────────


def pptx_deck() -> None:
    prs = Presentation()
    title = prs.slides.add_slide(prs.slide_layouts[0])
    title.shapes.title.text = "Review deck"
    title.placeholders[1].text = "Placeholders, text, table, chart, notes"

    bullets = prs.slides.add_slide(prs.slide_layouts[1])
    bullets.shapes.title.text = "Bullets"
    tf = bullets.placeholders[1].text_frame
    tf.text = "First bullet"
    for text in ["Second bullet", "Third bullet"]:
        p = tf.add_paragraph()
        p.text = text
        p.level = 1

    table_slide = prs.slides.add_slide(prs.slide_layouts[5])
    table_slide.shapes.title.text = "Table"
    rows, cols = 3, 3
    table = table_slide.shapes.add_table(rows, cols, PInches(1), PInches(2), PInches(6), PInches(1.5)).table
    for r in range(rows):
        for c in range(cols):
            table.cell(r, c).text = f"r{r}c{c}"

    chart_slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_slide.shapes.title.text = "Chart"
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("S1", (1, 3, 2))
    chart_slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, PInches(1), PInches(2), PInches(6), PInches(3), chart_data)

    notes = prs.slides.add_slide(prs.slide_layouts[5])
    notes.shapes.title.text = "Notes"
    notes.notes_slide.notes_text_frame.text = "Speaker notes for the notes slide."

    prs.save(OUT / "pptx-deck-basic.pptx")
    record(
        "pptx-deck-basic.pptx",
        "pptx",
        ["placeholders", "bullets", "table", "chart", "notes"],
        "four slides: title, bullets, table, chart + notes",
    )


def main() -> None:
    print(f"generating corpus in {OUT}")
    for fn in [
        docx_merged_table,
        docx_mixed_orientation,
        docx_two_columns,
        docx_comments,
        docx_large_table,
        xlsx_charts,
        xlsx_conditional,
        xlsx_validation,
        xlsx_frozen_hidden_1904,
        xlsx_large,
        xlsx_number_formats,
        pdf_rotated,
        pdf_image_only,
        pdf_encrypted,
        pdf_many_pages,
        pptx_deck,
    ]:
        fn()

    manifest_path = OUT / "manifest.json"
    manifest_path.write_text(
        json.dumps({"version": 1, "generator": "scripts/fixtures/generate_review_corpus.py", "entries": manifest}, indent=2)
        + "\n",
        encoding="utf-8",
    )
    print(f"wrote {manifest_path} ({len(manifest)} entries)")


if __name__ == "__main__":
    main()
