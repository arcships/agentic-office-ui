"""Build the 0.8 fixture corpus and keep the fixture manifests in sync.

Issue #72 keeps the generated sample corpus honest: every fixture has to carry
the structure a reviewer reported a defect in, and the manifest has to say which
feature a file exists for. This module owns the fixtures added for the 0.8
batch, their feature tags and the fixture -> feature -> issue -> test coverage
matrix. ``scripts/generate_test_materials.py`` imports it, so one run builds the
whole corpus and rewrites both manifests.

The builders avoid desktop Office entirely. DOCX/PPTX parts are written with
python-docx plus raw OOXML where the library cannot express a feature; XLSX
parts use openpyxl plus raw package edits for the macro stub and workbook
protection; PDFs use ReportLab and pypdf. Macro and media payloads are inert
fixture stubs, not executable content.
"""

from __future__ import annotations

import hashlib
import json
import struct
import zipfile
import zlib
from datetime import datetime
from pathlib import Path
from typing import Callable

from docx import Document
from docx.oxml import parse_xml
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart,
    BarChart,
    BarChart3D,
    DoughnutChart,
    LineChart,
    PieChart,
    RadarChart,
    Reference,
    ScatterChart,
)
from openpyxl.formatting.rule import (
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
from openpyxl.styles import PatternFill, Protection
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.workbook.protection import WorkbookProtection
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject,
    DecodedStreamObject,
    DictionaryObject,
    IndirectObject,
    NameObject,
    NumberObject,
    TextStringObject,
)
from reportlab import rl_config
from reportlab.lib.pagesizes import A4, letter
from reportlab.pdfbase.pdfdoc import PDFString
from reportlab.pdfgen import canvas as pdf_canvas

ROOT = Path(__file__).resolve().parents[1]
RL_FIXED_DATE = "D:20200101000000Z"
RL_FIXED_ID = "0123456789abcdef0123456789abcdef"


# ---------------------------------------------------------------------------
# Shared OOXML helpers
# ---------------------------------------------------------------------------

WORD_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
WORD_ALL_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
)


def read_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as source:
        return {name: source.read(name) for name in source.namelist()}


def write_parts(path: Path, entries: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as target:
        for name in sorted(entries):
            target.writestr(name, entries[name])


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


def add_docx_header_footer(doc: Document, header_text: str) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(10.5)
    for section in doc.sections:
        section.header.paragraphs[0].text = header_text
        section.footer.paragraphs[0].text = "Agentic Office UI 0.8 fixture corpus"


def append_docx_content(doc: Document, xml: str) -> None:
    """Append a raw ``w:p``/``w:sdt`` block before the final section properties."""
    doc.element.body.sectPr.addprevious(parse_xml(xml))


# ---------------------------------------------------------------------------
# DOCX: embedded drawing objects and OMML
# ---------------------------------------------------------------------------


def _tiny_png(width: int = 16, height: int = 16, rgba: bytes = b"\x25\x63\xeb\xff") -> bytes:
    raw = b"".join(b"\x00" + rgba * width for _ in range(height))

    def chunk(kind: bytes, payload: bytes) -> bytes:
        value = kind + payload
        return struct.pack(">I", len(payload)) + value + struct.pack(">I", zlib.crc32(value) & 0xFFFFFFFF)

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )


def _docx_chart_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Embedded Chart</a:t></a:r></a:p>"
        '</c:rich></c:tx><c:overlay val="0"/></c:title><c:plotArea><c:layout/><c:barChart>'
        '<c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/>'
        "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>Revenue</c:v></c:pt>"
        "</c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val=\"3\"/>"
        '<c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt>'
        "</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache>"
        '<c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>120</c:v></c:pt>'
        '<c:pt idx="1"><c:v>180</c:v></c:pt><c:pt idx="2"><c:v>240</c:v></c:pt></c:numCache></c:numRef>'
        '</c:val></c:ser><c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>'
        '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/><c:crosses val="autoZero"/>'
        '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
        '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/>'
        '<c:tickLblPos val="nextTo"/><c:crossAx val="111111111"/><c:crosses val="autoZero"/>'
        '<c:crossBetween val="between"/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>'
        '</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>'
    )


def _docx_diagram_data_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst>'
        '<dgm:pt modelId="0" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>'
        '<dgm:pt modelId="1"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:r><a:rPr lang="en-US"/><a:t>Analyse</a:t></a:r></a:p></dgm:t></dgm:pt>'
        '<dgm:pt modelId="2"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:r><a:rPr lang="en-US"/><a:t>Design</a:t></a:r></a:p></dgm:t></dgm:pt>'
        '<dgm:pt modelId="3"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:r><a:rPr lang="en-US"/><a:t>Deliver</a:t></a:r></a:p></dgm:t></dgm:pt>'
        "</dgm:ptLst><dgm:cxnLst>"
        '<dgm:cxn modelId="4" srcId="0" destId="1" srcOrd="0" destOrd="0"/>'
        '<dgm:cxn modelId="5" srcId="1" destId="2" srcOrd="0" destOrd="1"/>'
        '<dgm:cxn modelId="6" srcId="2" destId="3" srcOrd="0" destOrd="2"/>'
        "</dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>"
    )


def _docx_diagram_part_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:fixture/layout/process">'
        '<dgm:title val=""/><dgm:desc val=""/><dgm:catLst/><dgm:sampData/><dgm:styleData/>'
        '<dgm:layoutNode><dgm:alg type="lin"/><dgm:shape type="roundRect"/><dgm:presOf/>'
        '<dgm:constrLst/><dgm:forEach name="sibTrans" axis="ch"/></dgm:layoutNode></dgm:layoutDef>'
    )


def _docx_diagram_quick_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:fixture/quickstyle">'
        '<dgm:title val=""/><dgm:desc val=""/><dgm:catLst/><dgm:scene3d/><dgm:styleLbl name="node0">'
        "<dgm:scene3d/><dgm:sp3d/><dgm:txPr/><dgm:style>"
        '<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
        '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
        '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
        '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'
        "</dgm:style></dgm:styleLbl></dgm:styleDef>"
    )


def _docx_diagram_colors_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:fixture/colors">'
        '<dgm:title val=""/><dgm:desc val=""/><dgm:catLst/><dgm:styleLbl name="node0">'
        '<dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst>'
        '<dgm:linClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:linClrLst>'
        "<dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst/><dgm:txEffectClrLst/>"
        "</dgm:styleLbl></dgm:colorsDef>"
    )


def make_docx_embedded_objects(path: Path) -> None:
    """Text box, shape, SmartArt, chart, OLE object and OMML equation."""
    doc = Document()
    add_docx_header_footer(doc, "Embedded Objects Fixture")
    doc.add_heading("Embedded Objects and DrawingML", 0)
    doc.add_paragraph(
        "This fixture carries the drawing-object families the viewer has to preserve: "
        "a text box, a shape, SmartArt, an embedded chart, an OLE object and an OMML equation."
    )

    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Text box: </w:t></w:r><w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="3200400" cy="914400"/>'
        '<wp:docPr id="101" name="Text Box Fixture"/><a:graphic><a:graphicData '
        'uri="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingShape"><wps:wsp>'
        '<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3200400" cy="914400"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EFF6FF"/></a:solidFill>'
        '<a:ln><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></a:ln></wps:spPr>'
        "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Text box body: quarterly review.</w:t></w:r></w:p>"
        "</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline>"
        "</w:drawing></w:r></w:p>",
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Shape: </w:t></w:r><w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2286000" cy="685800"/>'
        '<wp:docPr id="102" name="Shape Fixture"/><a:graphic><a:graphicData '
        'uri="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingShape"><wps:wsp>'
        '<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="685800"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0F766E"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>'
        "</wp:inline></w:drawing></w:r></w:p>",
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>SmartArt: </w:t></w:r><w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="1600200"/>'
        '<wp:docPr id="103" name="SmartArt Fixture"/><a:graphic><a:graphicData '
        'uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        '<dgm:relIds r:dm="rIdDgmData" r:lo="rIdDgmLayout" r:qs="rIdDgmQuick" r:cs="rIdDgmColors"/>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>",
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Chart: </w:t></w:r><w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="2743200"/>'
        '<wp:docPr id="104" name="Chart Fixture"/><a:graphic><a:graphicData '
        'uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>",
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>OLE object: </w:t></w:r><w:r><w:object w:dxaOrig="1440" w:dyaOrig="1440">'
        '<v:shape id="OleShape" style="width:72pt;height:32pt"><v:imagedata r:id="rIdOlePreview" o:title="OLE preview"/></v:shape>'
        '<o:OLEObject Type="Embed" ProgID="Package" ShapeID="OleShape" DrawAspect="Icon" '
        'ObjectID="_1700000001" r:id="rIdOleObject"/></w:object></w:r></w:p>',
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Equation: </w:t></w:r><m:oMathPara><m:oMath>'
        "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>"
        "<m:r><m:t>+</m:t></m:r><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>"
        "</m:oMath></m:oMathPara></w:p>",
    )
    doc.save(path)

    entries = read_parts(path)
    entries["word/charts/chart1.xml"] = _docx_chart_xml().encode()
    entries["word/charts/_rels/chart1.xml.rels"] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    )
    entries["word/diagrams/data1.xml"] = _docx_diagram_data_xml().encode()
    entries["word/diagrams/layout1.xml"] = _docx_diagram_part_xml().encode()
    entries["word/diagrams/quickStyle1.xml"] = _docx_diagram_quick_xml().encode()
    entries["word/diagrams/colors1.xml"] = _docx_diagram_colors_xml().encode()
    entries["word/media/ole-preview.png"] = _tiny_png()
    entries["word/embeddings/oleObject1.bin"] = build_ole_package_stub()

    content_types = entries["[Content_Types].xml"].decode("utf-8")
    if 'Extension="png"' not in content_types:
        content_types = content_types.replace(
            '<Default Extension="jpeg" ContentType="image/jpeg"/>',
            '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/>',
        )
    content_types = add_content_type_overrides(
        content_types,
        {
            "/word/charts/chart1.xml": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            "/word/diagrams/data1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
            "/word/diagrams/layout1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
            "/word/diagrams/quickStyle1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
            "/word/diagrams/colors1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
            "/word/embeddings/oleObject1.bin": "application/vnd.openxmlformats-officedocument.oleObject",
        },
    )
    entries["[Content_Types].xml"] = content_types.encode("utf-8")
    entries["word/_rels/document.xml.rels"] = add_relationships(
        entries["word/_rels/document.xml.rels"].decode("utf-8"),
        [
            (
                "rIdChart1",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                "charts/chart1.xml",
            ),
            (
                "rIdDgmData",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
                "diagrams/data1.xml",
            ),
            (
                "rIdDgmLayout",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
                "diagrams/layout1.xml",
            ),
            (
                "rIdDgmQuick",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
                "diagrams/quickStyle1.xml",
            ),
            (
                "rIdDgmColors",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
                "diagrams/colors1.xml",
            ),
            (
                "rIdOlePreview",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "media/ole-preview.png",
            ),
            (
                "rIdOleObject",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                "embeddings/oleObject1.bin",
            ),
        ],
    ).encode("utf-8")
    write_parts(path, entries)


# ---------------------------------------------------------------------------
# DOCX: fields, content controls, tracked moves and a footnote in a table
# ---------------------------------------------------------------------------


def make_docx_fields_controls(path: Path) -> None:
    """PAGE/NUMPAGES/STYLEREF, SDT blocks, tracked moves and a table footnote."""
    doc = Document()
    add_docx_header_footer(doc, "Fields and Controls Fixture")
    doc.add_heading("Fields, Content Controls and Moves", 0)
    doc.add_paragraph(
        "This fixture carries PAGE, NUMPAGES and STYLEREF fields with roman and letter "
        "formats, block and inline content controls, a tracked move and a footnote "
        "anchored inside a table cell."
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Page field: </w:t></w:r>'
        '<w:fldSimple w:instr=" PAGE \\* roman "><w:r><w:t>iv</w:t></w:r></w:fldSimple>'
        '<w:r><w:t> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES \\* ALPHA "><w:r><w:t>D</w:t></w:r></w:fldSimple>'
        "<w:r><w:t> pages.</w:t></w:r></w:p>",
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Style reference: </w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> STYLEREF "Heading 1" \\* ROMAN </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>I</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    )
    append_docx_content(
        doc,
        f'<w:sdt {WORD_ALL_NS}><w:sdtPr><w:alias w:val="Customer Name"/>'
        '<w:tag w:val="fixture.customer.name"/><w:id w:val="101"/><w:text/></w:sdtPr>'
        "<w:sdtContent><w:p><w:r><w:t>Content control body: Northwind Industrial Supplies.</w:t></w:r></w:p>"
        "</w:sdtContent></w:sdt>",
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Inline control: </w:t></w:r><w:sdt>'
        '<w:sdtPr><w:alias w:val="Risk Rating"/><w:tag w:val="fixture.risk"/><w:id w:val="102"/>'
        '<w:dropDownList><w:listItem w:displayText="Low" w:value="Low"/>'
        '<w:listItem w:displayText="High" w:value="High"/></w:dropDownList></w:sdtPr>'
        "<w:sdtContent><w:r><w:t>Low</w:t></w:r></w:sdtContent></w:sdt>"
        "<w:r><w:t>.</w:t></w:r></w:p>",
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Moved from: </w:t></w:r>'
        '<w:moveFromRangeStart w:id="201" w:name="move-one" w:author="Fixture Reviewer" '
        'w:date="2026-07-11T00:20:00Z"/><w:moveFrom w:id="202" w:author="Fixture Reviewer" '
        'w:date="2026-07-11T00:20:00Z"><w:r><w:t>old clause location</w:t></w:r></w:moveFrom>'
        '<w:moveFromRangeEnd w:id="201"/></w:p>',
    )
    append_docx_content(
        doc,
        f'<w:p {WORD_ALL_NS}><w:r><w:t>Moved to: </w:t></w:r>'
        '<w:moveToRangeStart w:id="203" w:name="move-one" w:author="Fixture Reviewer" '
        'w:date="2026-07-11T00:20:00Z"/><w:moveTo w:id="204" w:author="Fixture Reviewer" '
        'w:date="2026-07-11T00:20:00Z"><w:r><w:t>new clause location</w:t></w:r></w:moveTo>'
        '<w:moveToRangeEnd w:id="203"/></w:p>',
    )
    doc.add_paragraph("The table below anchors a footnote inside a cell instead of the body.")
    table = doc.add_table(rows=3, cols=3)
    table.style = "Table Grid"
    for row_index, row in enumerate(table.rows):
        for column_index, cell in enumerate(row.cells):
            cell.text = f"R{row_index}C{column_index}"
    cell = table.cell(1, 1)
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.add_run("Net total")
    paragraph._p.append(
        parse_xml(
            f'<w:r {WORD_ALL_NS}><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
            '<w:footnoteReference w:id="2"/></w:r>'
        )
    )
    doc.add_paragraph("Body footnote reference after the table.")
    doc.paragraphs[-1]._p.append(
        parse_xml(
            f'<w:r {WORD_ALL_NS}><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
            '<w:footnoteReference w:id="1"/></w:r>'
        )
    )
    doc.save(path)

    entries = read_parts(path)
    entries["word/footnotes.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:footnotes {WORD_NS}>"
        '<w:footnote w:id="1"><w:p><w:r><w:t>Body footnote: PAGE and NUMPAGES fields resolve per section.</w:t>'
        '</w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Table footnote: anchored inside a '
        "table cell, not the body.</w:t></w:r></w:p></w:footnote></w:footnotes>"
    ).encode("utf-8")
    entries["[Content_Types].xml"] = add_content_type_overrides(
        entries["[Content_Types].xml"].decode("utf-8"),
        {
            "/word/footnotes.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        },
    ).encode("utf-8")
    entries["word/_rels/document.xml.rels"] = add_relationships(
        entries["word/_rels/document.xml.rels"].decode("utf-8"),
        [
            (
                "rIdFixtureFootnotes",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
                "footnotes.xml",
            ),
        ],
    ).encode("utf-8")
    write_parts(path, entries)


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------


def make_xlsx_chart_families(path: Path) -> None:
    """Column/bar, stacked/100%, line, pie, doughnut, scatter, radar, area, combo and 3D."""
    workbook = Workbook()
    data = workbook.active
    data.title = "Data"
    data.append(["Category", "Revenue", "Cost"])
    for row in [("Q1", 10, 4), ("Q2", 20, -3), ("Q3", -5, 8), ("Q4", 30, 12)]:
        data.append(list(row))
    revenue = Reference(data, min_col=2, min_row=1, max_row=5)
    cost = Reference(data, min_col=3, min_row=1, max_row=5)
    categories = Reference(data, min_col=1, min_row=2, max_row=5)

    charts: list[tuple[str, object]] = []

    column = BarChart()
    column.type = "col"
    column.title = "Column"
    column.add_data(revenue, titles_from_data=True)
    column.set_categories(categories)
    charts.append(("Column", column))

    stacked = BarChart()
    stacked.type = "bar"
    stacked.grouping = "stacked"
    stacked.overlap = 100
    stacked.title = "Stacked Bar"
    stacked.add_data(revenue, titles_from_data=True)
    stacked.set_categories(categories)
    charts.append(("Stacked Bar", stacked))

    percent = BarChart()
    percent.type = "col"
    percent.grouping = "percentStacked"
    percent.title = "100% Column"
    percent.add_data(revenue, titles_from_data=True)
    percent.set_categories(categories)
    charts.append(("100% Column", percent))

    line = LineChart()
    line.title = "Smooth Line"
    line.add_data(revenue, titles_from_data=True)
    line.set_categories(categories)
    for series in line.series:
        series.smooth = True
        series.marker.symbol = "circle"
    charts.append(("Smooth Line", line))

    pie = PieChart()
    pie.title = "Pie"
    pie.add_data(revenue, titles_from_data=True)
    charts.append(("Pie", pie))

    doughnut = DoughnutChart()
    doughnut.title = "Doughnut"
    doughnut.add_data(revenue, titles_from_data=True)
    charts.append(("Doughnut", doughnut))

    scatter = ScatterChart()
    scatter.title = "Scatter"
    scatter.add_data(Reference(data, min_col=2, min_row=2, max_row=5), titles_from_data=False)
    scatter.add_data(cost, titles_from_data=False)
    charts.append(("Scatter", scatter))

    radar = RadarChart()
    radar.title = "Radar"
    radar.add_data(revenue, titles_from_data=True)
    radar.set_categories(categories)
    charts.append(("Radar", radar))

    area = AreaChart()
    area.title = "Area"
    area.add_data(revenue, titles_from_data=True)
    area.set_categories(categories)
    charts.append(("Area", area))

    combo = BarChart()
    combo.title = "Combo One Axis"
    combo.add_data(revenue, titles_from_data=True)
    combo.set_categories(categories)
    combo_line = LineChart()
    combo_line.add_data(cost, titles_from_data=True)
    combo += combo_line
    charts.append(("Combo One Axis", combo))

    secondary = BarChart()
    secondary.title = "Combo Two Axes"
    secondary.add_data(revenue, titles_from_data=True)
    secondary.set_categories(categories)
    secondary.y_axis.crosses = "max"
    secondary_line = LineChart()
    secondary_line.add_data(cost, titles_from_data=True)
    secondary_line.y_axis.axId = 200
    secondary_line.y_axis.title = "Secondary"
    secondary_line.y_axis.crosses = "max"
    secondary += secondary_line
    charts.append(("Combo Two Axes", secondary))

    fallback = BarChart3D()
    fallback.title = "3D Fallback"
    fallback.add_data(revenue, titles_from_data=True)
    fallback.set_categories(categories)
    charts.append(("3D Fallback", fallback))

    families = workbook.create_sheet("Families")
    families.append(["Family", "Chart part"])
    for index, (name, _) in enumerate(charts, start=1):
        families.append([name, f"xl/charts/chart{index}.xml"])

    chart_sheet = workbook.create_sheet("Charts")
    chart_sheet.append(["Chart family", "Anchor"])
    for index, (name, chart) in enumerate(charts, start=1):
        chart_sheet.cell(row=2 + (index - 1) * 16, column=1, value=name)
        chart_sheet.add_chart(chart, f"B{2 + (index - 1) * 16}")
    workbook.save(path)


def make_xlsx_databars_icons(path: Path) -> None:
    """Data bars over negative values, 4- and 5-icon sets and all cfvo families."""
    workbook = Workbook()
    bars = workbook.active
    bars.title = "Negative Data Bars"
    bars.append(["Metric", "Value"])
    for index, value in enumerate([-40, -12, 3, 18, 45, -8, 27], start=2):
        bars.append([f"M{index - 1}", value])
    bars.conditional_formatting.add(
        "B2:B8",
        DataBarRule(
            start_type="num",
            start_value=-50,
            end_type="num",
            end_value=50,
            color="2563EB",
            showValue=True,
        ),
    )
    bars.conditional_formatting.add(
        "B2:B8",
        ColorScaleRule(
            start_type="percentile",
            start_value=10,
            start_color="F8696B",
            mid_type="percentile",
            mid_value=50,
            mid_color="FFEB84",
            end_type="percentile",
            end_value=90,
            end_color="63BE7B",
        ),
    )
    bars.conditional_formatting.add(
        "B2:B8",
        FormulaRule(formula=["B2<0"], fill=PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")),
    )

    icons = workbook.create_sheet("Icon Sets")
    icons.append(["Rating", "Arrows"])
    for index, value in enumerate([10, 25, 40, 60, 80, 95], start=2):
        icons.append([value, value])
    icons.conditional_formatting.add("A2:A7", IconSetRule("5Rating", "num", [0, 20, 40, 60, 80]))
    icons.conditional_formatting.add("B2:B7", IconSetRule("4Arrows", "percent", [0, 25, 50, 75]))
    workbook.save(path)


def make_xlsx_protected_1904(path: Path) -> None:
    """1904 date system plus workbook- and sheet-protection passwords."""
    workbook = Workbook()
    workbook.epoch = CALENDAR_MAC_1904
    workbook.security = WorkbookProtection(
        workbookPassword="arcships",
        lockStructure=True,
        lockWindows=True,
    )
    dates = workbook.active
    dates.title = "Dates"
    dates.append(["Protected 1904 workbook"])
    dates.append(["Date", datetime(2026, 1, 1)])
    dates["B2"].number_format = "yyyy-mm-dd"
    dates.append(["Formula", "=B2+30"])
    dates["B3"].number_format = "yyyy-mm-dd"

    protected = workbook.create_sheet("Protected")
    protected.append(["Only unlocked cells are editable"])
    protected.append(["Locked value", 100])
    protected.append(["Unlocked input", 5])
    protected["B3"].protection = Protection(locked=False)
    protected.protection.sheet = True
    protected.protection.password = "arcships"
    protected.protection.formatCells = False
    protected.protection.insertRows = False
    protected.protection.deleteRows = False
    protected.protection.sort = False
    protected.protection.autoFilter = False
    protected.protection.selectLockedCells = False
    protected.protection.selectUnlockedCells = False
    workbook.save(path)


def make_xlsx_large_scale(path: Path) -> None:
    """100k populated rows and a 10k column used range, streamed to keep memory low."""
    workbook = Workbook(write_only=True)
    rows = workbook.create_sheet("Rows100k")
    rows.append(["Row", "Amount"])
    for index in range(1, 100_001):
        rows.append([index, index * 2])

    columns = workbook.create_sheet("Columns10k")
    columns.append([f"C{index}" for index in range(1, 10_001)])
    columns.append([f"R1C{index}" if index % 5 else index for index in range(1, 10_001)])
    workbook.save(path)


# ---------------------------------------------------------------------------
# VBA project stub for the macro-enabled workbook
# ---------------------------------------------------------------------------

SECTOR_SIZE = 512
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD


def _cfb_directory_entry(
    name: str,
    entry_type: int,
    *,
    child: int = 0xFFFFFFFF,
    start: int = ENDOFCHAIN,
    size: int = 0,
    right: int = 0xFFFFFFFF,
    left: int = 0xFFFFFFFF,
) -> bytes:
    raw_name = name.encode("utf-16le") + b"\x00\x00"
    if len(raw_name) > 64:
        raise ValueError(f"CFB directory name too long: {name}")
    entry = bytearray(128)
    entry[0 : len(raw_name)] = raw_name
    struct.pack_into("<H", entry, 64, len(raw_name))
    entry[66] = entry_type
    entry[67] = 1  # black node
    struct.pack_into("<I", entry, 68, left)
    struct.pack_into("<I", entry, 72, right)
    struct.pack_into("<I", entry, 76, child)
    struct.pack_into("<I", entry, 116, start)
    struct.pack_into("<Q", entry, 120, size)
    return bytes(entry)


def build_ole_package_stub() -> bytes:
    """A small, inert OLE compound file used as ``word/embeddings/oleObject1.bin``.

    A DOCX OLE object references an embedded package. Keeping a real compound
    file (even with inert text) exercises the OLE branch instead of a plain
    byte blob that no OLE reader can open.
    """
    payload = (
        b"Inert Agentic Office UI OLE package payload. Not an executable object.\n"
    ).ljust(4200, b" ")
    count = (len(payload) + SECTOR_SIZE - 1) // SECTOR_SIZE
    chain = list(range(2, 2 + count))
    fat = [FREESECT] * 128
    fat[0] = FATSECT
    fat[1] = ENDOFCHAIN
    for index, sector in enumerate(chain):
        fat[sector] = chain[index + 1] if index + 1 < len(chain) else ENDOFCHAIN
    directory = b"".join(
        [
            _cfb_directory_entry("Root Entry", 5, child=1),
            _cfb_directory_entry("Package", 2, start=chain[0], size=len(payload)),
        ]
    ).ljust(SECTOR_SIZE, b"\x00")
    header = bytearray(SECTOR_SIZE)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", header, 24, 0x003E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<I", header, 40, 0)
    struct.pack_into("<I", header, 44, 1)
    struct.pack_into("<I", header, 48, 1)
    struct.pack_into("<I", header, 56, 4096)
    struct.pack_into("<I", header, 60, ENDOFCHAIN)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)
    for index in range(109):
        struct.pack_into("<I", header, 76 + index * 4, 0 if index == 0 else FREESECT)
    return b"".join(
        [
            bytes(header),
            b"".join(struct.pack("<I", value) for value in fat),
            directory,
            payload.ljust(len(chain) * SECTOR_SIZE, b"\x00"),
        ]
    )


def build_vba_project_stub() -> bytes:
    """A small, inert OLE compound file used as ``xl/vbaProject.bin``.

    The streams contain plain fixture text, not executable VBA. The point is to
    exercise the macro-enabled package branch (content type, relationship and
    macro warning) without shipping a runnable macro.
    """
    this_workbook = (
        b'Attribute VB_Name = "ThisWorkbook"\r\n'
        b"' Inert Agentic Office UI fixture text. This is not a runnable VBA project.\r\n"
    ).ljust(4200, b" ")
    module_one = (
        b'Attribute VB_Name = "Module1"\r\n'
        b"' Inert fixture module. No executable payload.\r\n"
    ).ljust(4200, b" ")

    chains: dict[str, tuple[list[int], bytes]] = {}
    next_sector = 2
    for name, data in (("dir", this_workbook), ("ThisWorkbook", module_one)):
        count = (len(data) + SECTOR_SIZE - 1) // SECTOR_SIZE
        chain = list(range(next_sector, next_sector + count))
        next_sector += count
        chains[name] = (chain, data)

    fat = [FREESECT] * 128
    fat[0] = FATSECT
    fat[1] = ENDOFCHAIN
    for chain, _ in chains.values():
        for index, sector in enumerate(chain):
            fat[sector] = chain[index + 1] if index + 1 < len(chain) else ENDOFCHAIN

    directory = b"".join(
        [
            _cfb_directory_entry("Root Entry", 5, child=1),
            _cfb_directory_entry("VBA", 1, child=2),
            _cfb_directory_entry(
                "dir",
                2,
                start=chains["dir"][0][0],
                size=len(chains["dir"][1]),
                right=3,
            ),
            _cfb_directory_entry(
                "ThisWorkbook",
                2,
                start=chains["ThisWorkbook"][0][0],
                size=len(chains["ThisWorkbook"][1]),
            ),
        ]
    ).ljust(SECTOR_SIZE, b"\x00")

    header = bytearray(SECTOR_SIZE)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", header, 24, 0x003E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<I", header, 40, 0)
    struct.pack_into("<I", header, 44, 1)
    struct.pack_into("<I", header, 48, 1)
    struct.pack_into("<I", header, 52, 0)
    struct.pack_into("<I", header, 56, 4096)
    struct.pack_into("<I", header, 60, ENDOFCHAIN)
    struct.pack_into("<I", header, 64, 0)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)
    struct.pack_into("<I", header, 72, 0)
    for index in range(109):
        struct.pack_into("<I", header, 76 + index * 4, 0 if index == 0 else FREESECT)

    return b"".join(
        [
            bytes(header),
            b"".join(struct.pack("<I", value) for value in fat),
            directory,
            chains["dir"][1].ljust(len(chains["dir"][0]) * SECTOR_SIZE, b"\x00"),
            chains["ThisWorkbook"][1].ljust(len(chains["ThisWorkbook"][0]) * SECTOR_SIZE, b"\x00"),
        ]
    )


def make_xlsx_macro_enabled(path: Path) -> None:
    """A deterministic ``.xlsm`` package with an inert VBA project stub."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Macro Stub"
    sheet["A1"] = "Macro-enabled workbook fixture"
    sheet["A2"] = "The VBA project is an inert stub; it contains no executable macro."
    sheet["B3"] = 42
    workbook.save(path)

    entries = read_parts(path)
    entries["xl/vbaProject.bin"] = build_vba_project_stub()
    content_types = entries["[Content_Types].xml"].decode("utf-8").replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    )
    entries["[Content_Types].xml"] = add_content_type_overrides(
        content_types,
        {"/xl/vbaProject.bin": "application/vnd.ms-office.vbaProject"},
    ).encode("utf-8")
    entries["xl/_rels/workbook.xml.rels"] = add_relationships(
        entries["xl/_rels/workbook.xml.rels"].decode("utf-8"),
        [
            (
                "rIdVbaProject",
                "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                "vbaProject.bin",
            ),
        ],
    ).encode("utf-8")
    write_parts(path, entries)


# ---------------------------------------------------------------------------
# PPTX: master, layouts, placeholders, SmartArt, video and CJK text
# ---------------------------------------------------------------------------

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
PPTX_NS = f'xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"'
PPTX_XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

PPTX_THEME = (
    f'{PPTX_XML_HEAD}<a:theme xmlns:a="{NS_A}" name="Fixture Theme"><a:themeElements>'
    '<a:clrScheme name="Fixture"><a:dk1><a:srgbClr val="1F2937"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="1E3A5F"/></a:dk2><a:lt2><a:srgbClr val="F1F5F9"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1>'
    '<a:accent2><a:srgbClr val="0F766E"/></a:accent2><a:accent3><a:srgbClr val="B45309"/></a:accent3>'
    '<a:accent4><a:srgbClr val="7C3AED"/></a:accent4><a:accent5><a:srgbClr val="BE123C"/></a:accent5>'
    '<a:accent6><a:srgbClr val="4D7C0F"/></a:accent6><a:hlink><a:srgbClr val="1D4ED8"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="7E22CE"/></a:folHlink></a:clrScheme>'
    '<a:fontScheme name="Fixture"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface="Microsoft YaHei"/>'
    '<a:cs typeface=""/><a:font script="Hans" typeface="宋体"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/>'
    '<a:ea typeface="Microsoft YaHei"/><a:cs typeface=""/><a:font script="Hans" typeface="宋体"/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Fixture"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:fillStyleLst>'
    '<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/>'
    '</a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
    '<a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>'
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill>'
    '<a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr>'
    '</a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>'
    "</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
)


def _pptx_text_shape(
    shape_id: int,
    name: str,
    placeholder: str,
    paragraphs: str,
    x: int,
    y: int,
    cx: int,
    cy: int,
) -> str:
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr><p:ph {placeholder}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        '<p:txBody><a:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:normAutofit/></a:bodyPr>'
        f"<a:lstStyle/>{paragraphs}</p:txBody></p:sp>"
    )


def _pptx_latin(text: str, size: int = 2400, bold: bool = False) -> str:
    weight = ' b="1"' if bold else ""
    return f'<a:p><a:r><a:rPr lang="en-US" sz="{size}"{weight}/><a:t>{text}</a:t></a:r></a:p>'


def _pptx_cjk(text: str) -> str:
    return (
        f'<a:p><a:pPr algn="l"/><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="2400" dirty="0">'
        '<a:latin typeface="Calibri"/><a:ea typeface="Microsoft YaHei"/></a:rPr>'
        f"<a:t>{text}</a:t></a:r><a:endParaRPr lang=\"zh-CN\" altLang=\"en-US\"/></a:p>"
    )


def _pptx_rels(entries: list[tuple[str, str, str]]) -> str:
    body = "".join(
        f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{target}"/>'
        for rel_id, rel_type, target in entries
    )
    return (
        f'{PPTX_XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{body}</Relationships>"
    )


def make_pptx_structure(path: Path) -> None:
    """One master, three layouts, placeholders, SmartArt, a video and CJK runs."""
    master = (
        f'{PPTX_XML_HEAD}<p:sldMaster {PPTX_NS}><p:cSld><p:bg><p:bgRef idx="1001">'
        '<a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{_pptx_text_shape(2, "Title Placeholder 1", "type=\"title\"", _pptx_latin("Click to edit Master title style", 3200, True), 457200, 274638, 8229600, 1143000)}'
        f'{_pptx_text_shape(3, "Text Placeholder 2", "type=\"body\" idx=\"1\"", _pptx_latin("Click to edit Master text styles") + _pptx_latin("Second level", 2000), 457200, 1600200, 8229600, 4525963)}'
        f'{_pptx_text_shape(4, "Date Placeholder 3", "type=\"dt\" sz=\"half\" idx=\"2\"", "<a:p><a:endParaRPr lang=\"en-US\"/></a:p>", 457200, 6356350, 2133600, 365125)}'
        f'{_pptx_text_shape(5, "Footer Placeholder 4", "type=\"ftr\" sz=\"quarter\" idx=\"3\"", "<a:p><a:endParaRPr lang=\"en-US\"/></a:p>", 3124200, 6356350, 3352800, 365125)}'
        f'{_pptx_text_shape(6, "Slide Number Placeholder 5", "type=\"sldNum\" sz=\"quarter\" idx=\"4\"", "<a:p><a:fld id=\"{2A1B3C4D-1111-4444-8888-1234567890AB}\" type=\"slidenum\"><a:rPr lang=\"en-US\"/><a:t>1</a:t></a:fld></a:p>", 7772400, 6356350, 1470025, 365125)}'
        "</p:spTree></p:cSld>"
        '<p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
        'accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/>'
        '<p:sldLayoutId id="2147483650" r:id="rId2"/><p:sldLayoutId id="2147483651" r:id="rId3"/>'
        "</p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn=\"l\"><a:defRPr sz=\"3200\" b=\"1\">"
        '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/>'
        '<a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="0" indent="0">'
        '<a:defRPr sz="2000"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/>'
        '<a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:defPPr>'
        '<a:defRPr lang="en-US"/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>'
    )
    layout_title = (
        f'{PPTX_XML_HEAD}<p:sldLayout {PPTX_NS} type="title" preserve="1"><p:cSld name="Title Slide"><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{_pptx_text_shape(2, "Title 1", "type=\"ctrTitle\"", _pptx_latin("Click to edit Master title style", 3200, True), 685800, 2130425, 7772400, 1470025)}'
        f'{_pptx_text_shape(3, "Subtitle 2", "type=\"subTitle\" idx=\"1\"", _pptx_latin("Click to edit Master subtitle style", 2000), 1371600, 3886200, 6400800, 1752600)}'
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
    )
    layout_content = (
        f'{PPTX_XML_HEAD}<p:sldLayout {PPTX_NS} type="obj" preserve="1"><p:cSld name="Title and Content"><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{_pptx_text_shape(2, "Title 1", "type=\"title\"", _pptx_latin("Click to edit Master title style", 3200, True), 457200, 274638, 8229600, 1143000)}'
        f'{_pptx_text_shape(3, "Content Placeholder 2", "type=\"body\" idx=\"1\"", _pptx_latin("Click to edit Master text styles"), 457200, 1600200, 8229600, 4525963)}'
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
    )
    layout_video = (
        f'{PPTX_XML_HEAD}<p:sldLayout {PPTX_NS} type="obj" preserve="1"><p:cSld name="Video with Caption"><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{_pptx_text_shape(2, "Title 1", "type=\"title\"", _pptx_latin("Click to edit Master title style", 3200, True), 457200, 274638, 8229600, 1143000)}'
        f'{_pptx_text_shape(3, "Video Placeholder 2", "type=\"pic\" idx=\"1\"", "<a:p><a:endParaRPr lang=\"en-US\"/></a:p>", 1371600, 1600200, 6400800, 3600000)}'
        f'{_pptx_text_shape(4, "Caption 3", "type=\"body\" idx=\"2\"", _pptx_latin("Click to add caption", 1800), 1371600, 5300000, 6400800, 600000)}'
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
    )
    slide_title = (
        f'{PPTX_XML_HEAD}<p:sld {PPTX_NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{_pptx_text_shape(2, "Title 1", "type=\"ctrTitle\"", _pptx_latin("Structure Fixture Deck", 3600, True), 685800, 1800000, 7772400, 1400000)}'
        f'{_pptx_text_shape(3, "Subtitle 2", "type=\"subTitle\" idx=\"1\"", _pptx_latin("Masters, layouts, placeholders, SmartArt, video and CJK metrics", 2000), 1371600, 3500000, 6400800, 1200000)}'
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )
    smartart_frame = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="SmartArt Fixture"/>'
        '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="3700000"/>'
        '<a:ext cx="7000000" cy="1200000"/></p:xfrm><a:graphic><a:graphicData '
        'uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        f'<dgm:relIds xmlns:dgm="{NS_DGM}" r:dm="rIdDgmData" r:lo="rIdDgmLayout" '
        'r:qs="rIdDgmQuick" r:cs="rIdDgmColors"/></a:graphicData></a:graphic></p:graphicFrame>'
    )
    slide_content = (
        f'{PPTX_XML_HEAD}<p:sld {PPTX_NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{_pptx_text_shape(2, "Title 1", "type=\"title\"", "<a:p><a:r><a:rPr lang=\"zh-CN\" altLang=\"en-US\" sz=\"3000\" b=\"1\"><a:latin typeface=\"Calibri\"/><a:ea typeface=\"Microsoft YaHei\"/></a:rPr><a:t>版式、SmartArt 与 CJK 文本</a:t></a:r></a:p>", 457200, 274638, 8229600, 900000)}'
        f'{_pptx_text_shape(3, "Content Placeholder 2", "type=\"body\" idx=\"1\"", _pptx_cjk("中文排版度量：宋体与微软雅黑回退") + _pptx_latin("Latin metrics sit beside CJK glyphs.", 2000) + _pptx_cjk("第二段中文用于测试换行与字距。"), 457200, 1300000, 8229600, 2100000)}'
        f"{smartart_frame}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )
    video_picture = (
        '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Video Placeholder"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/>'
        '</p:cNvPicPr><p:nvPr><p:ph type="pic" idx="1"/><a:videoFile r:link="rIdVideoFile"/><p:extLst>'
        '<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media '
        'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rIdVideoMedia"/>'
        "</p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed=\"rIdPoster\"/>"
        '<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1371600" y="1600200"/>'
        '<a:ext cx="6400800" cy="3600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "</p:spPr></p:pic>"
    )
    slide_video = (
        f'{PPTX_XML_HEAD}<p:sld {PPTX_NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{_pptx_text_shape(2, "Title 1", "type=\"title\"", _pptx_latin("Video Placeholder", 3000, True), 457200, 274638, 8229600, 900000)}'
        f"{video_picture}"
        f'{_pptx_text_shape(5, "Caption 3", "type=\"body\" idx=\"2\"", _pptx_latin("Inert media stub for layout and placeholder coverage.", 1600), 1371600, 5300000, 6400800, 600000)}'
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )
    diagram_data = (
        f'{PPTX_XML_HEAD}<dgm:dataModel xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}"><dgm:ptLst>'
        '<dgm:pt modelId="0" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>'
        '<dgm:pt modelId="1"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:r><a:rPr lang="en-US"/><a:t>Plan</a:t></a:r></a:p></dgm:t></dgm:pt>'
        '<dgm:pt modelId="2"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:r><a:rPr lang="zh-CN"><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>设计</a:t></a:r></a:p></dgm:t></dgm:pt>'
        '<dgm:pt modelId="3"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:r><a:rPr lang="en-US"/><a:t>Deliver</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst>'
        '<dgm:cxnLst><dgm:cxn modelId="4" srcId="0" destId="1" srcOrd="0" destOrd="0"/>'
        '<dgm:cxn modelId="5" srcId="1" destId="2" srcOrd="1" destOrd="1"/>'
        '<dgm:cxn modelId="6" srcId="2" destId="3" srcOrd="2" destOrd="2"/></dgm:cxnLst>'
        "<dgm:bg/><dgm:whole/></dgm:dataModel>"
    )
    diagram_layout = (
        f'{PPTX_XML_HEAD}<dgm:layoutDef xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}" '
        'uniqueId="urn:fixture/layout/linear"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst/>'
        '<dgm:sampData/><dgm:styleData/><dgm:layoutNode><dgm:alg type="lin"/><dgm:shape type="rect"/>'
        "<dgm:presOf/><dgm:constrLst/><dgm:forEach name=\"sibTrans\" axis=\"ch\"/></dgm:layoutNode></dgm:layoutDef>"
    )
    diagram_quick = (
        f'{PPTX_XML_HEAD}<dgm:styleDef xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}" '
        'uniqueId="urn:fixture/quickstyle"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst/><dgm:scene3d/>'
        '<dgm:styleLbl name="node0"><dgm:scene3d/><dgm:sp3d/><dgm:txPr/><dgm:style>'
        '<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
        '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor">'
        '<a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl></dgm:styleDef>'
    )
    diagram_colors = (
        f'{PPTX_XML_HEAD}<dgm:colorsDef xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}" '
        'uniqueId="urn:fixture/colors"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst/>'
        '<dgm:styleLbl name="node0"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst>'
        '<dgm:linClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:linClrLst><dgm:effectClrLst/>'
        "<dgm:txLinClrLst/><dgm:txFillClrLst/><dgm:txEffectClrLst/></dgm:styleLbl></dgm:colorsDef>"
    )
    diagram_drawing = (
        f'{PPTX_XML_HEAD}<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" '
        f'xmlns:a="{NS_A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="1" name=""/><dsp:cNvGrpSpPr/>'
        '<dsp:nvPr/></dsp:nvGrpSpPr><dsp:grpSpPr/><dsp:sp modelId="1"><dsp:nvSpPr><dsp:cNvPr id="2" name="Plan"/>'
        '<dsp:cNvSpPr/><dsp:nvPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="2000000" cy="800000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/>'
        '<a:p><a:r><a:rPr lang="en-US"/><a:t>Plan</a:t></a:r></a:p></dsp:txBody></dsp:sp>'
        '<dsp:sp modelId="2"><dsp:nvSpPr><dsp:cNvPr id="3" name="Design"/><dsp:cNvSpPr/><dsp:nvPr/>'
        '</dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="2200000" y="0"/><a:ext cx="2000000" cy="800000"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0F766E"/></a:solidFill>'
        '</dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"><a:ea typeface="Microsoft YaHei"/>'
        "</a:rPr><a:t>设计</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>"
    )

    content_types = (
        f'{PPTX_XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="mp4" ContentType="video/mp4"/>'
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
        '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'
        '<Override PartName="/ppt/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>'
        '<Override PartName="/ppt/diagrams/layout1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>'
        '<Override PartName="/ppt/diagrams/quickStyle1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/>'
        '<Override PartName="/ppt/diagrams/colors1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>'
        '<Override PartName="/ppt/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )
    presentation = (
        f'{PPTX_XML_HEAD}<p:presentation {PPTX_NS} saveSubsetFonts="1"><p:sldMasterIdLst>'
        '<p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>'
        '<p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/>'
        '</p:sldIdLst><p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/>'
        "</a:defPPr></p:defaultTextStyle></p:presentation>"
    )
    root_rels = _pptx_rels(
        [
            ("rId1", f"{NS_R}/officeDocument", "ppt/presentation.xml"),
            ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"),
            ("rId3", f"{NS_R}/extended-properties", "docProps/app.xml"),
        ]
    )
    presentation_rels = _pptx_rels(
        [
            ("rId1", f"{NS_R}/slideMaster", "slideMasters/slideMaster1.xml"),
            ("rId2", f"{NS_R}/slide", "slides/slide1.xml"),
            ("rId3", f"{NS_R}/slide", "slides/slide2.xml"),
            ("rId4", f"{NS_R}/slide", "slides/slide3.xml"),
            ("rId5", f"{NS_R}/theme", "theme/theme1.xml"),
            ("rId6", f"{NS_R}/presProps", "presProps.xml"),
            ("rId7", f"{NS_R}/viewProps", "viewProps.xml"),
            ("rId8", f"{NS_R}/tableStyles", "tableStyles.xml"),
        ]
    )
    master_rels = _pptx_rels(
        [
            ("rId1", f"{NS_R}/slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("rId2", f"{NS_R}/slideLayout", "../slideLayouts/slideLayout2.xml"),
            ("rId3", f"{NS_R}/slideLayout", "../slideLayouts/slideLayout3.xml"),
            ("rId4", f"{NS_R}/theme", "../theme/theme1.xml"),
        ]
    )
    layout_rels = _pptx_rels([("rId1", f"{NS_R}/slideMaster", "../slideMasters/slideMaster1.xml")])
    slide1_rels = _pptx_rels([("rId1", f"{NS_R}/slideLayout", "../slideLayouts/slideLayout1.xml")])
    slide2_rels = _pptx_rels(
        [
            ("rId1", f"{NS_R}/slideLayout", "../slideLayouts/slideLayout2.xml"),
            ("rIdDgmData", f"{NS_R}/diagramData", "../diagrams/data1.xml"),
            ("rIdDgmLayout", f"{NS_R}/diagramLayout", "../diagrams/layout1.xml"),
            ("rIdDgmQuick", f"{NS_R}/diagramQuickStyle", "../diagrams/quickStyle1.xml"),
            ("rIdDgmColors", f"{NS_R}/diagramColors", "../diagrams/colors1.xml"),
            (
                "rIdDgmDrawing",
                "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
                "../diagrams/drawing1.xml",
            ),
        ]
    )
    slide3_rels = _pptx_rels(
        [
            ("rId1", f"{NS_R}/slideLayout", "../slideLayouts/slideLayout3.xml"),
            ("rIdVideoFile", f"{NS_R}/video", "../media/video1.mp4"),
            ("rIdVideoMedia", "http://schemas.microsoft.com/office/2007/relationships/media", "../media/video1.mp4"),
            ("rIdPoster", f"{NS_R}/image", "../media/poster.png"),
        ]
    )
    core = (
        f'{PPTX_XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title>PPTX structure fixture</dc:title><dc:creator>Agentic Office UI</dc:creator>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>'
    )
    app = (
        f'{PPTX_XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        "<Application>Agentic Office UI fixture generator</Application><Slides>3</Slides>"
        "<PresentationFormat>Widescreen</PresentationFormat></Properties>"
    )
    pres_props = f"{PPTX_XML_HEAD}<p:presentationPr {PPTX_NS}/>"
    view_props = (
        f'{PPTX_XML_HEAD}<p:viewPr {PPTX_NS}><p:normalViewPr/><p:slideViewPr/><p:notesTextViewPr/>'
        '<p:gridSpacing cx="72008" cy="72008"/></p:viewPr>'
    )
    table_styles = f'{PPTX_XML_HEAD}<a:tblStyleLst xmlns:a="{NS_A}" def="{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}"/>'
    video_stub = (
        struct.pack(">I", 24)
        + b"ftypisom"
        + struct.pack(">I", 512)
        + b"isomiso2"
        + struct.pack(">I", 8)
        + b"free"
    )
    entries = {
        "[Content_Types].xml": content_types.encode(),
        "_rels/.rels": root_rels.encode(),
        "docProps/core.xml": core.encode(),
        "docProps/app.xml": app.encode(),
        "ppt/presentation.xml": presentation.encode(),
        "ppt/_rels/presentation.xml.rels": presentation_rels.encode(),
        "ppt/presProps.xml": pres_props.encode(),
        "ppt/viewProps.xml": view_props.encode(),
        "ppt/tableStyles.xml": table_styles.encode(),
        "ppt/theme/theme1.xml": PPTX_THEME.encode(),
        "ppt/slideMasters/slideMaster1.xml": master.encode(),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": master_rels.encode(),
        "ppt/slideLayouts/slideLayout1.xml": layout_title.encode(),
        "ppt/slideLayouts/slideLayout2.xml": layout_content.encode(),
        "ppt/slideLayouts/slideLayout3.xml": layout_video.encode(),
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": layout_rels.encode(),
        "ppt/slideLayouts/_rels/slideLayout2.xml.rels": layout_rels.encode(),
        "ppt/slideLayouts/_rels/slideLayout3.xml.rels": layout_rels.encode(),
        "ppt/slides/slide1.xml": slide_title.encode(),
        "ppt/slides/slide2.xml": slide_content.encode(),
        "ppt/slides/slide3.xml": slide_video.encode(),
        "ppt/slides/_rels/slide1.xml.rels": slide1_rels.encode(),
        "ppt/slides/_rels/slide2.xml.rels": slide2_rels.encode(),
        "ppt/slides/_rels/slide3.xml.rels": slide3_rels.encode(),
        "ppt/diagrams/data1.xml": diagram_data.encode(),
        "ppt/diagrams/layout1.xml": diagram_layout.encode(),
        "ppt/diagrams/quickStyle1.xml": diagram_quick.encode(),
        "ppt/diagrams/colors1.xml": diagram_colors.encode(),
        "ppt/diagrams/drawing1.xml": diagram_drawing.encode(),
        "ppt/media/poster.png": _tiny_png(64, 36),
        "ppt/media/video1.mp4": video_stub,
    }
    write_parts(path, entries)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------


def _pdf_writer(writer: PdfWriter) -> None:
    """Keep the trailer ID fixed so repeated generations hash identically."""
    writer._ID = ArrayObject(  # type: ignore[attr-defined]
        [TextStringObject(RL_FIXED_ID), TextStringObject(RL_FIXED_ID)]
    )


def _as_object(value):
    return value.get_object() if isinstance(value, IndirectObject) else value


def make_pdf_acroform_annotations(path: Path) -> None:
    """AcroForm widgets, annotations, attachments, an OCG and signature metadata."""
    rl_config.invariant = 1
    base = path.with_suffix(".base.pdf")
    canvas = pdf_canvas.Canvas(
        str(base),
        pagesize=A4,
        invariant=1,
        pageCompression=1,
        initialFontName="Helvetica",
    )
    canvas.setTitle("AcroForm and Annotation Fixture")
    canvas.setAuthor("Agentic Office UI")
    canvas.setSubject("Real AcroForm, annotations, attachments, OCG and signature metadata")
    canvas.setFont("Helvetica", 18)
    canvas.drawString(72, 760, "AcroForm and Annotation Fixture")
    canvas.setFont("Helvetica", 11)
    canvas.drawString(72, 735, "Text field, checkbox, choice, highlight, sticky note and link.")
    canvas.acroForm.textfield(
        name="customer_name",
        value="Northwind Industrial",
        x=72,
        y=680,
        width=240,
        height=24,
        fontSize=11,
    )
    canvas.acroForm.checkbox(name="approved", checked=True, x=72, y=640, size=16)
    canvas.acroForm.choice(
        name="plan",
        value="Pro",
        options=["Basic", "Pro", "Enterprise"],
        x=72,
        y=600,
        width=160,
        height=22,
        fontSize=10,
    )
    canvas.highlightAnnotation(
        "Highlighted contract clause",
        Rect=(70, 725, 500, 742),
        Color=[0.83, 0.89, 0.95],
    )
    canvas.textAnnotation("Sticky note annotation", Rect=(520, 680, 540, 700))
    canvas.freeTextAnnotation(
        "Free text annotation",
        PDFString("/Helv 10 Tf 0.1 0.1 0.1 rg"),
        Rect=(72, 540, 360, 580),
    )
    canvas.linkRect("Jump to page 2", "page2", Rect=(72, 500, 220, 518), relative=0)
    canvas.showPage()
    canvas.setFont("Helvetica", 18)
    canvas.drawString(72, 760, "Attachment, OCG and Signature Metadata")
    canvas.setFont("Helvetica", 11)
    canvas.drawString(72, 735, "Layered content and embedded files live on this page.")
    canvas.bookmarkPage("page2")
    canvas.showPage()
    canvas.save()

    writer = PdfWriter(clone_from=str(base))
    writer.add_attachment(
        "fixture-readme.txt",
        b"Inert attachment for the Agentic Office UI fixture corpus.\n",
    )
    writer.add_attachment(
        "fixture-meta.json",
        b'{"fixture":"pdf-acroform-annotations.pdf","kind":"inert"}\n',
    )
    root = writer._root_object  # type: ignore[attr-defined]
    ocg = DictionaryObject(
        {
            NameObject("/Type"): NameObject("/OCG"),
            NameObject("/Name"): TextStringObject("Fixture Layer"),
        }
    )
    ocg_ref = writer._add_object(ocg)  # type: ignore[attr-defined]
    root[NameObject("/OCProperties")] = DictionaryObject(
        {
            NameObject("/OCGs"): ArrayObject([ocg_ref]),
            NameObject("/D"): DictionaryObject(
                {
                    NameObject("/Order"): ArrayObject([ocg_ref]),
                    NameObject("/ON"): ArrayObject([ocg_ref]),
                }
            ),
        }
    )
    page = writer.pages[1]
    resources = _as_object(page["/Resources"])
    properties = _as_object(resources.get("/Properties"))
    if properties is None:
        properties = DictionaryObject()
        resources[NameObject("/Properties")] = properties
    properties[NameObject("/MC0")] = ocg_ref
    stream = DecodedStreamObject()
    stream.set_data(b"q /OC /MC0 BDC 0.1 0.4 0.8 rg 72 650 360 30 re f EMC Q\n")
    page[NameObject("/Contents")] = ArrayObject(
        [page["/Contents"], writer._add_object(stream)]  # type: ignore[attr-defined]
    )

    signature = DictionaryObject(
        {
            NameObject("/Type"): NameObject("/Sig"),
            NameObject("/Filter"): NameObject("/Adobe.PPKLite"),
            NameObject("/SubFilter"): NameObject("/adbe.pkcs7.detached"),
            NameObject("/ByteRange"): ArrayObject(
                [NumberObject(0), NumberObject(0), NumberObject(0), NumberObject(0)]
            ),
            NameObject("/Contents"): TextStringObject("0" * 64),
            NameObject("/M"): TextStringObject(RL_FIXED_DATE),
            NameObject("/Name"): TextStringObject("Fixture Signature"),
            NameObject("/Reason"): TextStringObject("Unsigned fixture metadata"),
            NameObject("/Location"): TextStringObject("Shanghai"),
        }
    )
    signature_ref = writer._add_object(signature)  # type: ignore[attr-defined]
    acro_form = _as_object(root["/AcroForm"])
    page_zero = writer.pages[0]
    signature_field = DictionaryObject(
        {
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Widget"),
            NameObject("/FT"): NameObject("/Sig"),
            NameObject("/T"): TextStringObject("Signature1"),
            NameObject("/Rect"): ArrayObject(
                [NumberObject(360), NumberObject(600), NumberObject(540), NumberObject(650)]
            ),
            NameObject("/F"): NumberObject(4),
            NameObject("/P"): page_zero.indirect_reference,
            NameObject("/V"): signature_ref,
        }
    )
    signature_field_ref = writer._add_object(signature_field)  # type: ignore[attr-defined]
    page_zero["/Annots"].append(signature_field_ref)
    acro_form["/Fields"].append(signature_field_ref)
    acro_form[NameObject("/SigFlags")] = NumberObject(3)
    root[NameObject("/Perms")] = DictionaryObject({NameObject("/DocMDP"): signature_ref})
    _pdf_writer(writer)
    with path.open("wb") as target:
        writer.write(target)
    base.unlink(missing_ok=True)


def make_pdf_1000_pages(path: Path) -> None:
    """A deterministic 1000-page PDF for pagination and virtual-scrolling tests."""
    rl_config.invariant = 1
    canvas = pdf_canvas.Canvas(
        str(path),
        pagesize=letter,
        pageCompression=1,
        invariant=1,
        initialFontName="Helvetica",
    )
    canvas.setTitle("1000-page fixture")
    canvas.setAuthor("Agentic Office UI")
    for page_number in range(1, 1001):
        canvas.setFont("Helvetica", 16)
        canvas.drawString(72, 720, f"PAGE-{page_number:04d}")
        canvas.setFont("Helvetica", 10)
        canvas.drawString(72, 690, "Large page-count fixture for pagination and virtual scrolling.")
        canvas.drawRightString(540, 60, f"{page_number} / 1000")
        canvas.showPage()
    canvas.save()


# ---------------------------------------------------------------------------
# Corpus metadata, manifest and coverage matrix
# ---------------------------------------------------------------------------

NEW_FIXTURES: dict[str, dict[str, object]] = {
    "docx-embedded-objects.docx": {
        "type": "docx",
        "category": "normal",
        "expected": (
            "parse a text box, a shape, SmartArt, an embedded chart, an OLE object and "
            "an OMML equation without dropping their relationships"
        ),
        "features": [
            "docx:textbox",
            "docx:shape",
            "docx:smartart",
            "docx:chart",
            "docx:ole-object",
            "docx:math-omml",
        ],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "docx-fields-controls.docx": {
        "type": "docx",
        "category": "normal",
        "expected": (
            "parse PAGE/NUMPAGES/STYLEREF fields with roman and letter formats, block and "
            "inline content controls, tracked moveFrom/moveTo revisions and a footnote in a table cell"
        ),
        "features": [
            "docx:field-page-roman",
            "docx:field-numpages-alpha",
            "docx:field-styleref-roman",
            "docx:content-control",
            "docx:tracked-move",
            "docx:footnote-in-table",
        ],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "xlsx-macros.xlsm": {
        "type": "xlsm",
        "category": "boundary",
        "expected": "open a macro-enabled workbook and report the inert VBA project stub without executing it",
        "features": ["xlsx:macro-enabled", "xlsx:vba-project-stub"],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "xlsx-protected-1904.xlsx": {
        "type": "xlsx",
        "category": "boundary",
        "expected": "open a date-1904 workbook whose workbook structure and worksheet carry passwords",
        "features": [
            "xlsx:date-1904",
            "xlsx:workbook-protection",
            "xlsx:sheet-protection",
        ],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "xlsx-chart-families.xlsx": {
        "type": "xlsx",
        "category": "normal",
        "expected": "parse column/bar, stacked/100%, line, pie, doughnut, scatter, radar, area, combo and 3D chart families",
        "features": [
            "xlsx:chart-family-bar",
            "xlsx:chart-family-line",
            "xlsx:chart-family-pie",
            "xlsx:chart-family-scatter",
            "xlsx:chart-family-radar",
            "xlsx:chart-family-area",
            "xlsx:chart-combo-one-axis",
            "xlsx:chart-combo-two-axis",
            "xlsx:chart-3d-fallback",
        ],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "xlsx-negative-databar-icons.xlsx": {
        "type": "xlsx",
        "category": "normal",
        "expected": "parse data bars over negative values, 4- and 5-icon sets, percentile/percent/list/formula conditional-format inputs",
        "features": [
            "xlsx:data-bar-negative",
            "xlsx:icon-set-4",
            "xlsx:icon-set-5",
            "xlsx:cfvo-num",
            "xlsx:cfvo-percent",
            "xlsx:cfvo-percentile",
            "xlsx:cfvo-formula",
        ],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "xlsx-large-scale.xlsx": {
        "type": "xlsx",
        "category": "large",
        "expected": "open and scroll a 100k-row sheet and a 10k-column used range without exhausting memory",
        "features": ["xlsx:large-rows", "xlsx:wide-columns"],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "deck-structure.pptx": {
        "type": "pptx",
        "category": "normal",
        "expected": "render one master, three layouts and their placeholders, SmartArt, a video placeholder and CJK runs",
        "features": [
            "pptx:slide-master",
            "pptx:slide-layout",
            "pptx:placeholder",
            "pptx:smartart",
            "pptx:video-placeholder",
            "pptx:cjk-text",
        ],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "pdf-acroform-annotations.pdf": {
        "type": "pdf",
        "category": "normal",
        "expected": "read AcroForm widgets, highlight/free-text/link annotations, attachments, an OCG and signature/DocMDP metadata",
        "features": [
            "pdf:acroform",
            "pdf:annotation",
            "pdf:attachment",
            "pdf:ocg",
            "pdf:signature-field",
            "pdf:docmdp",
        ],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
    "pdf-1000-pages.pdf": {
        "type": "pdf",
        "category": "large",
        "expected": "open 1000 pages, extract the first and last page markers and keep the 1000-page budget",
        "features": ["pdf:large", "pdf:text-layer"],
        "issues": [72],
        "tests": [
            "tests/unit/fixture-corpus.test.mjs",
            "scripts/verify_test_materials.py",
            "scripts/ci/verify-fixtures.mjs",
        ],
    },
}

EXISTING_FEATURES: dict[str, list[str]] = {
    "charts-images.xlsx": [
        "xlsx:chart",
        "xlsx:image",
        "xlsx:merged-cells",
        "xlsx:freeze-panes",
        "xlsx:data-validation",
        "xlsx:conditional-format",
    ],
    "chinese-mixed.docx": ["docx:cjk-text", "docx:mixed-script"],
    "cjk-report.pdf": ["pdf:cjk-text", "pdf:text-layer"],
    "contract-page.png": ["image:png"],
    "corrupted.docx": ["negative:corrupt-package"],
    "corrupted.pdf": ["negative:corrupt-package"],
    "corrupted.xlsx": ["negative:corrupt-package"],
    "demo.docx": ["docx:long-document"],
    "encrypted.pdf": ["pdf:encrypted"],
    "field-citations.json": ["ocr:field-citations"],
    "financial-model.xlsx": ["xlsx:formulas", "xlsx:number-format", "xlsx:multiple-sheets"],
    "invoice-table.docx": ["docx:table", "docx:merged-cells"],
    "invoice.png": ["image:png"],
    "large-contract.pdf": ["pdf:large"],
    "large-grid.xlsx": ["xlsx:large-grid"],
    "legal-contract.docx": ["docx:long-document"],
    "ocr-layout.json": ["ocr:layout"],
    "page-sizes.docx": ["docx:section-page-size", "docx:section-orientation"],
    "report-with-image.docx": ["docx:image"],
    "review-comments.docx": ["docx:comments", "docx:tracked-changes", "docx:footnotes", "docx:endnotes"],
    "review-threads.docx": [
        "docx:threaded-comments",
        "docx:resolved-comment",
        "docx:comment-in-table",
        "docx:comment-in-header",
        "docx:tracked-changes",
    ],
    "rotated-pages.pdf": ["pdf:rotation"],
    "sales-table-validated.xlsx": [
        "xlsx:freeze-panes",
        "xlsx:data-validation",
        "xlsx:table",
        "xlsx:column-width",
    ],
    "sales-table.xlsx": ["xlsx:table", "xlsx:formulas", "xlsx:number-format"],
    "sample.pdf": ["pdf:text-layer"],
    "scanned-invoice.pdf": ["pdf:scanned", "pdf:image-only"],
    "sections-columns.docx": [
        "docx:section-orientation",
        "docx:section-columns",
        "docx:textbox",
        "docx:watermark",
    ],
    "table-structures.docx": [
        "docx:table",
        "docx:vertical-merge",
        "docx:grid-span",
        "docx:nested-table",
        "docx:table-header-repeat",
    ],
}

EXISTING_REFERENCES: dict[str, dict[str, object]] = {
    "review-threads.docx": {"issues": [53], "tests": ["tests/unit/docx-fixture-fidelity.test.mjs"]},
    "table-structures.docx": {"issues": [53], "tests": ["tests/unit/docx-fixture-fidelity.test.mjs"]},
    "sections-columns.docx": {"issues": [53], "tests": ["tests/unit/docx-fixture-fidelity.test.mjs"]},
    "page-sizes.docx": {"issues": [58], "tests": ["tests/unit/docx-fixture-fidelity.test.mjs"]},
    "charts-images.xlsx": {"issues": [62], "tests": ["tests/component/xlsx-fixture-structure.test.mjs"]},
    "sales-table-validated.xlsx": {"issues": [62], "tests": ["tests/component/xlsx-fixture-structure.test.mjs"]},
    "corrupted.docx": {"issues": [62]},
    "corrupted.pdf": {"issues": [62]},
    "corrupted.xlsx": {"issues": [62]},
    "encrypted.pdf": {"issues": [62]},
}


def new_fixture_builders(samples: Path) -> dict[str, Callable[[], None]]:
    """Builders keyed by file name, matching the main generator's shape."""
    return {
        "docx-embedded-objects.docx": lambda: make_docx_embedded_objects(samples / "docx-embedded-objects.docx"),
        "docx-fields-controls.docx": lambda: make_docx_fields_controls(samples / "docx-fields-controls.docx"),
        "xlsx-macros.xlsm": lambda: make_xlsx_macro_enabled(samples / "xlsx-macros.xlsm"),
        "xlsx-protected-1904.xlsx": lambda: make_xlsx_protected_1904(samples / "xlsx-protected-1904.xlsx"),
        "xlsx-chart-families.xlsx": lambda: make_xlsx_chart_families(samples / "xlsx-chart-families.xlsx"),
        "xlsx-negative-databar-icons.xlsx": lambda: make_xlsx_databars_icons(samples / "xlsx-negative-databar-icons.xlsx"),
        "xlsx-large-scale.xlsx": lambda: make_xlsx_large_scale(samples / "xlsx-large-scale.xlsx"),
        "deck-structure.pptx": lambda: make_pptx_structure(samples / "deck-structure.pptx"),
        "pdf-acroform-annotations.pdf": lambda: make_pdf_acroform_annotations(samples / "pdf-acroform-annotations.pdf"),
        "pdf-1000-pages.pdf": lambda: make_pdf_1000_pages(samples / "pdf-1000-pages.pdf"),
    }


def _entry_features(name: str) -> list[str]:
    if name in NEW_FIXTURES:
        return list(NEW_FIXTURES[name]["features"])  # type: ignore[arg-type]
    if name in EXISTING_FEATURES:
        return list(EXISTING_FEATURES[name])
    raise RuntimeError(f"No feature tags declared for fixture {name!r}")


def _entry_references(name: str) -> tuple[list[int], list[str]]:
    if name in NEW_FIXTURES:
        return list(NEW_FIXTURES[name]["issues"]), list(NEW_FIXTURES[name]["tests"])  # type: ignore[arg-type]
    reference = EXISTING_REFERENCES.get(name, {})
    issues = list(reference.get("issues", []))  # type: ignore[arg-type]
    tests = ["scripts/ci/verify-fixtures.mjs", "scripts/verify_test_materials.py"]
    tests.extend(reference.get("tests", []))  # type: ignore[arg-type]
    suffix = Path(name).suffix
    if suffix in {".png", ".json"}:
        tests = ["scripts/ci/verify-fixtures.mjs"]
    return issues, list(dict.fromkeys(tests))


def update_corpus_manifests(samples: Path) -> None:
    """Refresh hashes, feature tags and the coverage matrix from what is on disk."""
    manifest_path = ROOT / "test-data" / "manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    existing = {entry["file"]: entry for entry in manifest["entries"]}
    files = sorted(
        path.name for path in samples.iterdir() if path.is_file() and path.name != "manifest.json"
    )
    unknown = [name for name in files if name not in NEW_FIXTURES and name not in EXISTING_FEATURES]
    if unknown:
        raise RuntimeError(f"Fixtures without feature metadata: {', '.join(unknown)}")

    entries: list[dict[str, object]] = []
    coverage: list[dict[str, object]] = []
    for name in files:
        entry = dict(existing.get(name, {}))
        if not entry:
            metadata = NEW_FIXTURES.get(name)
            if metadata is None:
                raise RuntimeError(f"No manifest metadata declared for fixture {name!r}")
            entry = {
                "file": name,
                "type": metadata["type"],
                "category": metadata["category"],
                "source": "synthetic-v1",
                "expected": metadata["expected"],
            }
        path = samples / name
        payload = path.read_bytes()
        entry["file"] = name
        entry["features"] = _entry_features(name)
        entry["bytes"] = len(payload)
        entry["sha256"] = hashlib.sha256(payload).hexdigest()
        entries.append(entry)
        issues, tests = _entry_references(name)
        coverage.append(
            {
                "file": name,
                "features": entry["features"],
                "issues": issues,
                "tests": tests,
            }
        )

    manifest["entries"] = entries
    manifest["coverage"] = "test-data/corpus-coverage.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    coverage_path = ROOT / "test-data" / "corpus-coverage.json"
    coverage_path.write_text(
        json.dumps(
            {
                "version": 1,
                "description": "Fixture -> feature -> issue -> test coverage matrix for the real fixture corpus.",
                "generatedBy": "scripts/generate_test_materials.py",
                "fixtures": coverage,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    doc_path = ROOT / "docs" / "fixture-corpus.md"
    rows = [
        "| Fixture | Features | Issues | Tests |",
        "| --- | --- | --- | --- |",
    ]
    for item in coverage:
        features = "<br>".join(f"`{feature}`" for feature in item["features"])  # type: ignore[arg-type]
        issues = ", ".join(f"#{issue}" for issue in item["issues"]) or "—"  # type: ignore[arg-type]
        tests = "<br>".join(f"`{test}`" for test in item["tests"])  # type: ignore[arg-type]
        rows.append(f"| `{item['file']}` | {features} | {issues} | {tests} |")
    doc_path.write_text(
        "# Fixture corpus coverage\n\n"
        "Generated by `scripts/generate_test_materials.py`. Do not edit by hand; "
        "re-run the generator after adding a fixture.\n\n"
        "The machine-readable matrix is `test-data/corpus-coverage.json`, and every "
        "manifest entry carries the same `features` list.\n\n"
        + "\n".join(rows)
        + "\n",
        encoding="utf-8",
    )
