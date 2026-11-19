#!/usr/bin/env python3
"""Generate the pivot-table and slicer XLSX fixture.

The workbook is intentionally small (one worksheet, six data rows, one pivot
table and one slicer) so the metadata parser, the drawing layer and the
export/reload path can be asserted without a large binary fixture. It uses only
Python's standard library so it can be regenerated without office tooling.

The pivot table's output is already written into the worksheet cells (that is
what Excel stores for a pivot table, and what the viewer paints); the pivot
parts describe where it came from. The slicer's visual is a DrawingML shape
inside an ``mc:AlternateContent`` block, exactly as Excel writes it.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
CORPUS = ROOT / "test-data" / "corpus"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

CONTENT_TYPES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>
"""

ROOT_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

WORKBOOK = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="SalesData">Data!$A$2:$C$6</definedName>
    <definedName name="CurrentProduct" localSheetId="0">Data!$B$5</definedName>
    <definedName name="RevenueTotal">SUM(Data!$C$2:$C$6)</definedName>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$C$6</definedName>
  </definedNames>
  <pivotCaches>
    <pivotCache cacheId="1" r:id="rId3"/>
  </pivotCaches>
  <extLst>
    <ext uri="{{BBE1A952-AA13-448e-AADC-164F8A28A991}}" xmlns:x14="{X14_NS}">
      <x14:slicerCaches>
        <x14:slicerCache r:id="rId4"/>
      </x14:slicerCaches>
    </ext>
  </extLst>
</workbook>
"""

WORKBOOK_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{REL_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="{REL_NS}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
</Relationships>
"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{MAIN_NS}">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"""

SHEET_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{REL_NS}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
  <Relationship Id="rId3" Type="{REL_NS}/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
"""

SHEET = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{MAIN_NS}" xmlns:r="{REL_NS}" xmlns:mc="{MC_NS}" xmlns:x14="{X14_NS}" mc:Ignorable="x14">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Region</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Product</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Revenue</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>East</t></is></c>
      <c r="B2" t="inlineStr"><is><t>Widget</t></is></c>
      <c r="C2"><v>100</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>West</t></is></c>
      <c r="B3" t="inlineStr"><is><t>Widget</t></is></c>
      <c r="C3"><v>150</v></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>North</t></is></c>
      <c r="B4" t="inlineStr"><is><t>Gadget</t></is></c>
      <c r="C4"><v>200</v></c>
    </row>
    <row r="5">
      <c r="A5" t="inlineStr"><is><t>South</t></is></c>
      <c r="B5" t="inlineStr"><is><t>Gadget</t></is></c>
      <c r="C5"><v>250</v></c>
    </row>
    <row r="6">
      <c r="A6" t="inlineStr"><is><t>East</t></is></c>
      <c r="B6" t="inlineStr"><is><t>Gadget</t></is></c>
      <c r="C6"><v>300</v></c>
    </row>
    <row r="10">
      <c r="J10" t="inlineStr"><is><t>Sum of Revenue</t></is></c>
      <c r="K10" t="inlineStr"><is><t>Column Labels</t></is></c>
    </row>
    <row r="11">
      <c r="J11" t="inlineStr"><is><t>Row Labels</t></is></c>
      <c r="K11" t="inlineStr"><is><t>Gadget</t></is></c>
      <c r="L11" t="inlineStr"><is><t>Widget</t></is></c>
      <c r="M11" t="inlineStr"><is><t>Grand Total</t></is></c>
    </row>
    <row r="12">
      <c r="J12" t="inlineStr"><is><t>East</t></is></c>
      <c r="K12"><v>300</v></c>
      <c r="L12"><v>100</v></c>
      <c r="M12"><v>400</v></c>
    </row>
    <row r="13">
      <c r="J13" t="inlineStr"><is><t>North</t></is></c>
      <c r="K13"><v>200</v></c>
      <c r="M13"><v>200</v></c>
    </row>
    <row r="14">
      <c r="J14" t="inlineStr"><is><t>Grand Total</t></is></c>
      <c r="K14"><v>500</v></c>
      <c r="L14"><v>100</v></c>
      <c r="M14"><v>600</v></c>
    </row>
  </sheetData>
  <drawing r:id="rId3"/>
  <pivotTableParts count="1">
    <pivotTablePart r:id="rId1"/>
  </pivotTableParts>
  <extLst>
    <ext uri="{{A8765BA9-456A-4dab-B4F3-ACF838C121DE}}" xmlns:x14="{X14_NS}">
      <x14:slicerList>
        <x14:slicer r:id="rId2"/>
      </x14:slicerList>
    </ext>
  </extLst>
</worksheet>
"""

PIVOT_TABLE = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="{MAIN_NS}" name="PivotTable1" cacheId="1" dataCaption="Values"
  updatedVersion="6" minRefreshableVersion="3" createdVersion="6"
  compact="0" outline="1" outlineData="1" gridDropZones="1"
  rowHeaderCaption="Region" colHeaderCaption="Product">
  <location ref="J10:M14" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="3">
    <pivotField axis="axisRow" showAll="0">
      <items count="5">
        <item x="0"/><item x="1"/><item x="2"/><item x="3"/><item t="default"/>
      </items>
    </pivotField>
    <pivotField axis="axisCol" showAll="0">
      <items count="4">
        <item x="0"/><item x="1"/><item t="default"/>
      </items>
    </pivotField>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <colFields count="1"><field x="1"/></colFields>
  <dataFields count="1">
    <dataField name="Sum of Revenue" fld="2" baseField="0" baseItem="0" numFmtId="0" subtotal="sum" showDataAs="normal"/>
  </dataFields>
  <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>
"""

PIVOT_CACHE = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="{MAIN_NS}" xmlns:r="{REL_NS}" r:id="rId1"
  refreshOnLoad="1" refreshedBy="Fixture" createdVersion="6" refreshedVersion="6"
  minRefreshableVersion="3" recordCount="5">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:C6" sheet="Data"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region" numFmtId="0">
      <sharedItems count="4"><s v="East"/><s v="West"/><s v="North"/><s v="South"/></sharedItems>
    </cacheField>
    <cacheField name="Product" numFmtId="0">
      <sharedItems count="2"><s v="Widget"/><s v="Gadget"/></sharedItems>
    </cacheField>
    <cacheField name="Revenue" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1">
        <n v="100"/><n v="150"/><n v="200"/><n v="250"/><n v="300"/>
      </sharedItems>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>
"""

PIVOT_CACHE_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="{REL_NS}/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>
"""

PIVOT_CACHE_RECORDS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="{MAIN_NS}" count="5">
  <r><x v="0"/><x v="0"/><n v="100"/></r>
  <r><x v="1"/><x v="0"/><n v="150"/></r>
  <r><x v="2"/><x v="1"/><n v="200"/></r>
  <r><x v="3"/><x v="1"/><n v="250"/></r>
  <r><x v="0"/><x v="1"/><n v="300"/></r>
</pivotCacheRecords>
"""

SLICER = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="{X14_NS}">
  <slicer name="Region" cache="Slicer_Region" caption="Region" startItem="0" columnCount="1" style="SlicerStyleLight1"/>
</slicers>
"""

SLICER_CACHE = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="{X14_NS}" name="Slicer_Region" sourceName="Region">
  <pivotTables>
    <pivotTable tabId="1" name="PivotTable1"/>
  </pivotTables>
  <data>
    <tabular pivotCacheId="1">
      <items count="4"><i x="0"/><i x="1" s="1"/><i x="2"/><i x="3"/></items>
    </tabular>
  </data>
</slicerCacheDefinition>
"""

DRAWING = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{XDR_NS}" xmlns:a="{A_NS}" xmlns:mc="{MC_NS}" xmlns:x14="{X14_NS}" mc:Ignorable="x14">
  <mc:AlternateContent>
    <mc:Choice Requires="x14">
      <xdr:twoCellAnchor editAs="oneCell">
        <xdr:from><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>16</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp macro="" textlink="">
          <xdr:nvSpPr>
            <xdr:cNvPr id="2" name="Slicer_Region"/>
            <xdr:cNvSpPr/>
          </xdr:nvSpPr>
          <xdr:spPr>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            <a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>
          </xdr:spPr>
          <xdr:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p><a:r><a:rPr lang="en-US" sz="1100" b="1"/><a:t>Region</a:t></a:r></a:p>
          </xdr:txBody>
        </xdr:sp>
        <xdr:clientData/>
      </xdr:twoCellAnchor>
    </mc:Choice>
  </mc:AlternateContent>
</xdr:wsDr>
"""

FILES = {
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": ROOT_RELS,
    "xl/workbook.xml": WORKBOOK,
    "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
    "xl/styles.xml": STYLES,
    "xl/worksheets/sheet1.xml": SHEET,
    "xl/worksheets/_rels/sheet1.xml.rels": SHEET_RELS,
    "xl/pivotTables/pivotTable1.xml": PIVOT_TABLE,
    "xl/pivotCache/pivotCacheDefinition1.xml": PIVOT_CACHE,
    "xl/pivotCache/pivotCacheRecords1.xml": PIVOT_CACHE_RECORDS,
    "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels": PIVOT_CACHE_RELS,
    "xl/slicers/slicer1.xml": SLICER,
    "xl/slicerCaches/slicerCache1.xml": SLICER_CACHE,
    "xl/drawings/drawing1.xml": DRAWING,
}


def main() -> None:
    CORPUS.mkdir(parents=True, exist_ok=True)
    target = CORPUS / "xlsx-pivot-slicers.xlsx"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in FILES.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, content)
    print(f"wrote {target}")


if __name__ == "__main__":
    main()
