#!/usr/bin/env python3
"""Generate test-data/corpus/docx-svg-pictures.docx deterministically.

The fixture carries the picture forms a DOCX produced by Word can contain:

- an inline SVG icon whose markup attempts script execution and external
  references (must be sanitised by the viewer, not executed);
- an inline SVG chart (vector bars) painted at an anchored size;
- a floating (``wp:anchor`` + ``wrapSquare``) SVG so the image layer path is
  covered as well as the in-flow runner;
- a Word-style dual ``a:blip`` with a raster ``r:embed`` fallback plus an
  ``asvg:svgBlip`` extension, where the SVG is the picture that must render;
- an SVG using ``foreignObject`` and an SVG using a filter, the two features
  the vector ``<img>`` pipeline cannot paint and must replace with a
  placeholder plus a diagnostic.

Only Python's standard library is used so the fixture can be regenerated in CI
without office tooling.
"""

from __future__ import annotations

import struct
import zipfile
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / "test-data" / "corpus" / "docx-svg-pictures.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
ASVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"

EMU_PER_PIXEL = 9525


def extent(width_px: int, height_px: int) -> tuple[int, int]:
    return width_px * EMU_PER_PIXEL, height_px * EMU_PER_PIXEL


def picture_graphic(relationship_id: str, name: str, width_px: int, height_px: int) -> str:
    cx, cy = extent(width_px, height_px)
    return (
        f'<a:graphic xmlns:a="{A_NS}">'
        f'<a:graphicData uri="{PIC_NS}">'
        f'<pic:pic xmlns:pic="{PIC_NS}">'
        f'<pic:nvPicPr><pic:cNvPr id="0" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{relationship_id}"/>'
        f"<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        f"</pic:pic></a:graphicData></a:graphic>"
    )


def dual_blip_picture_graphic(
    raster_relationship_id: str,
    svg_relationship_id: str,
    name: str,
    width_px: int,
    height_px: int,
) -> str:
    """Word's representation: raster fallback plus an ``asvg:svgBlip``."""
    cx, cy = extent(width_px, height_px)
    svg_extension = (
        '<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">'
        f'<asvg:svgBlip xmlns:asvg="{ASVG_NS}" r:embed="{svg_relationship_id}"/>'
        "</a:ext></a:extLst>"
    )
    return (
        f'<a:graphic xmlns:a="{A_NS}">'
        f'<a:graphicData uri="{PIC_NS}">'
        f'<pic:pic xmlns:pic="{PIC_NS}">'
        f'<pic:nvPicPr><pic:cNvPr id="0" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{raster_relationship_id}">{svg_extension}</a:blip>'
        f"<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        f"</pic:pic></a:graphicData></a:graphic>"
    )


def inline_drawing(graphic: str, name: str, width_px: int, height_px: int) -> str:
    cx, cy = extent(width_px, height_px)
    return (
        f"<w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">"
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="1" name="{name}"/>'
        f"{graphic}"
        "</wp:inline></w:drawing>"
    )


def floating_drawing(graphic: str, name: str, width_px: int, height_px: int) -> str:
    cx, cy = extent(width_px, height_px)
    return (
        "<w:drawing>"
        '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" '
        'simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" '
        'layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>720000</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapSquare wrapText="bothSides"/>'
        f'<wp:docPr id="2" name="{name}"/>'
        f"{graphic}"
        "</wp:anchor></w:drawing>"
    )


def paragraph(text: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def drawing_paragraph(drawing: str) -> str:
    return f"<w:p><w:r>{drawing}</w:r></w:p>"


ICON_SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
     width="96" height="96" viewBox="0 0 96 96">
  <defs>
    <linearGradient id="iconFill" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#2563eb"/>
      <stop offset="1" stop-color="#1e3a8a"/>
    </linearGradient>
  </defs>
  <script>window.__docxSvgPwned = true;</script>
  <rect width="96" height="96" rx="12" fill="url(#iconFill)"
        onload="window.__docxSvgPwned = true"/>
  <circle cx="48" cy="48" r="22" fill="#fbbf24"/>
  <path d="M40 48l6 6 12-14" fill="none" stroke="#1e3a8a" stroke-width="6"
        stroke-linecap="round" stroke-linejoin="round"/>
  <image xlink:href="https://example.invalid/tracker.png" width="8" height="8"/>
  <a xlink:href="javascript:alert(1)"><text x="10" y="90">icon</text></a>
</svg>
"""

CHART_SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="320" height="200" viewBox="0 0 320 200">
  <rect width="320" height="200" fill="#f8fafc"/>
  <line x1="32" y1="168" x2="304" y2="168" stroke="#94a3b8" stroke-width="2"/>
  <line x1="32" y1="16" x2="32" y2="168" stroke="#94a3b8" stroke-width="2"/>
  <rect x="56" y="108" width="40" height="60" fill="#2563eb"/>
  <rect x="112" y="76" width="40" height="92" fill="#3b82f6"/>
  <rect x="168" y="48" width="40" height="120" fill="#60a5fa"/>
  <rect x="224" y="92" width="40" height="76" fill="#93c5fd"/>
  <text x="160" y="196" text-anchor="middle" font-family="Arial" font-size="12"
        fill="#334155">Quarterly chart</text>
</svg>
"""

FLOATING_SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="200" height="120" viewBox="0 0 200 120">
  <rect width="200" height="120" rx="10" fill="#ecfdf5"/>
  <circle cx="60" cy="60" r="32" fill="#10b981"/>
  <rect x="108" y="30" width="72" height="16" rx="8" fill="#34d399"/>
  <rect x="108" y="56" width="56" height="12" rx="6" fill="#6ee7b7"/>
  <rect x="108" y="76" width="64" height="12" rx="6" fill="#a7f3d0"/>
</svg>
"""

DUAL_SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="144" height="96" viewBox="0 0 144 96">
  <rect width="144" height="96" fill="#fef3c7"/>
  <polygon points="72,16 120,80 24,80" fill="#f59e0b"/>
</svg>
"""

FOREIGN_OBJECT_SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="160" height="96" viewBox="0 0 160 96">
  <rect width="160" height="96" fill="#fee2e2"/>
  <foreignObject x="8" y="8" width="144" height="80">
    <div xmlns="http://www.w3.org/1999/xhtml">HTML content</div>
  </foreignObject>
</svg>
"""

FILTER_SVG = """<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="160" height="96" viewBox="0 0 160 96">
  <defs>
    <filter id="blur"><feGaussianBlur stdDeviation="4"/></filter>
  </defs>
  <rect width="160" height="96" fill="#e0e7ff"/>
  <circle cx="80" cy="48" r="28" fill="#6366f1" filter="url(#blur)"/>
</svg>
"""


def png_chunk(kind: bytes, payload: bytes) -> bytes:
    return (
        struct.pack(">I", len(payload))
        + kind
        + payload
        + struct.pack(">I", zlib.crc32(kind + payload) & 0xFFFFFFFF)
    )


def fallback_png(width: int = 16, height: int = 16) -> bytes:
    """A deterministic solid PNG, used as the raster half of the dual blip."""
    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    scanlines = b"".join(
        b"\x00" + bytes([248, 203, 61]) * width for _ in range(height)
    )
    return (
        b"\x89PNG\r\n\x1a\n"
        + png_chunk(b"IHDR", header)
        + png_chunk(b"IDAT", zlib.compress(scanlines, 9))
        + png_chunk(b"IEND", b"")
    )


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

PACKAGE_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOCUMENT_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId10" Type="{R_NS}/image" Target="media/icon.svg"/>
  <Relationship Id="rId11" Type="{R_NS}/image" Target="media/chart.svg"/>
  <Relationship Id="rId12" Type="{R_NS}/image" Target="media/floating.svg"/>
  <Relationship Id="rId13" Type="{R_NS}/image" Target="media/fallback.png"/>
  <Relationship Id="rId14" Type="{R_NS}/image" Target="media/dual.svg"/>
  <Relationship Id="rId15" Type="{R_NS}/image" Target="media/foreign-object.svg"/>
  <Relationship Id="rId16" Type="{R_NS}/image" Target="media/filter.svg"/>
</Relationships>
"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>
"""


def document_xml() -> str:
    parts = [
        paragraph("SVG picture fixture"),
        drawing_paragraph(
            inline_drawing(
                picture_graphic("rId10", "icon.svg", 96, 96),
                "SVG icon",
                96,
                96,
            )
        ),
        drawing_paragraph(
            inline_drawing(
                picture_graphic("rId11", "chart.svg", 320, 200),
                "SVG chart",
                320,
                200,
            )
        ),
        drawing_paragraph(
            floating_drawing(
                picture_graphic("rId12", "floating.svg", 200, 120),
                "Floating SVG",
                200,
                120,
            )
        ),
        drawing_paragraph(
            inline_drawing(
                dual_blip_picture_graphic("rId13", "rId14", "dual.svg", 144, 96),
                "Dual SVG",
                144,
                96,
            )
        ),
        drawing_paragraph(
            inline_drawing(
                picture_graphic("rId15", "foreign-object.svg", 160, 96),
                "Foreign object SVG",
                160,
                96,
            )
        ),
        drawing_paragraph(
            inline_drawing(
                picture_graphic("rId16", "filter.svg", 160, 96),
                "Filtered SVG",
                160,
                96,
            )
        ),
    ]
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:wp="{WP_NS}">'
        "<w:body>"
        + "".join(parts)
        + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )


def write_fixture() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    entries = (
        ("[Content_Types].xml", CONTENT_TYPES.encode("utf-8")),
        ("_rels/.rels", PACKAGE_RELS.encode("utf-8")),
        ("word/document.xml", document_xml().encode("utf-8")),
        ("word/_rels/document.xml.rels", DOCUMENT_RELS.encode("utf-8")),
        ("word/styles.xml", STYLES.encode("utf-8")),
        ("word/media/icon.svg", ICON_SVG.encode("utf-8")),
        ("word/media/chart.svg", CHART_SVG.encode("utf-8")),
        ("word/media/floating.svg", FLOATING_SVG.encode("utf-8")),
        ("word/media/dual.svg", DUAL_SVG.encode("utf-8")),
        ("word/media/foreign-object.svg", FOREIGN_OBJECT_SVG.encode("utf-8")),
        ("word/media/filter.svg", FILTER_SVG.encode("utf-8")),
        ("word/media/fallback.png", fallback_png()),
    )
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries:
            info = zipfile.ZipInfo(name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o600 & 0xFFFF) << 16
            archive.writestr(info, payload)


if __name__ == "__main__":
    write_fixture()
    print(f"wrote {OUTPUT}")
