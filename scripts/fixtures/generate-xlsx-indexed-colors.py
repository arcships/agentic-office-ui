#!/usr/bin/env python3
"""Generate the legacy indexed-colour XLSX fixtures deterministically.

Two workbooks exercise the two halves of ECMA-376 18.8.27:

* ``xlsx-indexed-colors.xlsx`` uses the standard 64-colour indexed palette
  (indices 0-63 plus the system foreground/background entries 64/65), without
  changing the palette in ``xl/styles.xml``.
* ``xlsx-indexed-custom-palette.xlsx`` writes a workbook-level
  ``<colors><indexedColors>`` list that overrides indices 8, 10 and 22, then
  styles cells with those indices.

Both files also carry an indexed font colour, so the fixture covers cell fills
and fonts. Only Python's standard library is used so the files can be
regenerated without office tooling.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
CORPUS = ROOT / "test-data" / "corpus"

STANDARD_INDEXED = (
    "00000000", "00FFFFFF", "00FF0000", "0000FF00", "000000FF",
    "00FFFF00", "00FF00FF", "0000FFFF", "00000000", "00FFFFFF",
    "00FF0000", "0000FF00", "000000FF", "00FFFF00", "00FF00FF",
    "0000FFFF", "00800000", "00008000", "00000080", "00808000",
    "00800080", "00008080", "00C0C0C0", "00808080", "009999FF",
    "00993366", "00FFFFCC", "00CCFFFF", "00660066", "00FF8080",
    "000066CC", "00CCCCFF", "00000080", "00FF00FF", "00FFFF00",
    "0000FFFF", "00800080", "00800000", "00008080", "000000FF",
    "0000CCFF", "00CCFFFF", "00CCFFCC", "00FFFF99", "0099CCFF",
    "00FF99CC", "00CC99FF", "00FFCC99", "003366FF", "0033CCCC",
    "0099CC00", "00FFCC00", "00FF9900", "00FF6600", "00666699",
    "00969696", "00003366", "00339966", "00003300", "00333300",
    "00993300", "00993366", "00333399", "00333333",
)

CUSTOM_INDEXED = list(STANDARD_INDEXED)
CUSTOM_INDEXED[8] = "00FF00FF"   # was black
CUSTOM_INDEXED[10] = "00112233"  # was red
CUSTOM_INDEXED[22] = "00ABCDEF"  # was silver

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""

PACKAGE_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Indexed" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""


def solid_fill(indexed: int, tint: float | None = None) -> str:
    tint_attribute = "" if tint is None else f' tint="{tint}"'
    return (
        '<fill><patternFill patternType="solid">'
        f'<fgColor indexed="{indexed}"{tint_attribute}/>'
        "</patternFill></fill>"
    )


def indexed_colors_xml(colors: tuple[str, ...] | list[str]) -> str:
    entries = "".join(f'<rgbColor rgb="{color}"/>' for color in colors)
    return f"<colors><indexedColors>{entries}</indexedColors></colors>"


def styles_xml(
    fills: list[str],
    *,
    indexed_colors: tuple[str, ...] | list[str],
    font_indexed: int | None = None,
) -> str:
    fonts = [
        '<font><name val="Calibri"/><family val="2"/><color theme="1"/><sz val="11"/><scheme val="minor"/></font>'
    ]
    if font_indexed is not None:
        fonts.append(
            '<font><name val="Calibri"/><family val="2"/>'
            f'<color indexed="{font_indexed}"/><sz val="11"/><scheme val="minor"/></font>'
        )

    fill_entries = [
        "<fill><patternFill/></fill>",
        '<fill><patternFill patternType="gray125"/></fill>',
        *fills,
    ]
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>']
    xfs.extend(
        f'<xf numFmtId="0" fontId="0" fillId="{2 + index}" borderId="0"/>'
        for index in range(len(fills))
    )
    if font_indexed is not None:
        xfs.append('<xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>'
        f'<fills count="{len(fill_entries)}">{"".join(fill_entries)}</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        f"{indexed_colors_xml(indexed_colors)}"
        "</styleSheet>"
    )


def worksheet_xml(cells: list[tuple[str, int, str]]) -> str:
    by_row: dict[int, list[tuple[str, int, str]]] = {}
    for reference, style_id, value in cells:
        row = int("".join(character for character in reference if character.isdigit()))
        by_row.setdefault(row, []).append((reference, style_id, value))

    rows = []
    for row in sorted(by_row):
        row_cells = "".join(
            f'<c r="{reference}" s="{style_id}" t="inlineStr"><is><t>{value}</t></is></c>'
            for reference, style_id, value in by_row[row]
        )
        rows.append(f'<row r="{row}">{row_cells}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(rows)}</sheetData>'
        "</worksheet>"
    )


def write_fixture(
    name: str,
    *,
    styles: str,
    cells: list[tuple[str, int, str]],
) -> None:
    path = CORPUS / name
    path.parent.mkdir(parents=True, exist_ok=True)
    entries = (
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", PACKAGE_RELS),
        ("xl/workbook.xml", WORKBOOK),
        ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
        ("xl/styles.xml", styles),
        ("xl/worksheets/sheet1.xml", worksheet_xml(cells)),
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for part_name, text in entries:
            info = zipfile.ZipInfo(part_name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (0o600 & 0xFFFF) << 16
            archive.writestr(info, text)
    print(f"wrote {path}")


def write_standard_fixture() -> None:
    fills = [
        solid_fill(10),
        solid_fill(22),
        solid_fill(55),
        solid_fill(63),
        solid_fill(9),
        solid_fill(64),
        solid_fill(65),
        solid_fill(0),
        solid_fill(10, 0.5),
        solid_fill(10, -0.5),
    ]
    cells = [
        ("A1", 1, "index 10"),
        ("A2", 2, "index 22"),
        ("A3", 3, "index 55"),
        ("A4", 4, "index 63"),
        ("A5", 5, "index 9"),
        ("A6", 6, "index 64"),
        ("A7", 7, "index 65"),
        ("A8", 8, "index 0"),
        ("B1", 9, "index 10 tint 0.5"),
        ("B2", 10, "index 10 tint -0.5"),
        ("C1", 11, "font index 10"),
    ]
    write_fixture(
        "xlsx-indexed-colors.xlsx",
        styles=styles_xml(fills, indexed_colors=STANDARD_INDEXED, font_indexed=10),
        cells=cells,
    )


def write_custom_fixture() -> None:
    fills = [
        solid_fill(8),
        solid_fill(10),
        solid_fill(9),
    ]
    cells = [
        ("A1", 1, "custom index 8"),
        ("A2", 2, "custom index 10"),
        ("A3", 3, "standard index 9"),
        ("B1", 4, "font custom index 10"),
    ]
    write_fixture(
        "xlsx-indexed-custom-palette.xlsx",
        styles=styles_xml(fills, indexed_colors=CUSTOM_INDEXED, font_indexed=10),
        cells=cells,
    )


if __name__ == "__main__":
    write_standard_fixture()
    write_custom_fixture()
