import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

import { buildChartMediaWorkbookBytes } from "../fixtures/chart-embedded-media.mjs";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Node: XmlNode,
  XMLSerializer: XmlSerializer,
} = requireFromXlsxCore("@xmldom/xmldom");
globalThis.DOMParser ??= XmlDomParser;
globalThis.Node ??= XmlNode;
globalThis.XMLSerializer ??= XmlSerializer;

const { strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

const OFFICE_THEME = {
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "1F497D",
  lt2: "EEECE1",
  accent1: "4F81BD",
  accent2: "C0504D",
  accent3: "9BBB59",
  accent4: "8064A2",
  accent5: "4BACC6",
  accent6: "F79646",
  hlink: "0000FF",
  folHlink: "800080",
};

const GRAYSCALE_THEME = {
  ...OFFICE_THEME,
  accent1: "404040",
  accent2: "595959",
  accent3: "737373",
  accent4: "8C8C8C",
  accent5: "A6A6A6",
  accent6: "BFBFBF",
};

function themeXml(theme) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="${theme.dk1}"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="${theme.lt1}"/></a:lt1>
      <a:dk2><a:srgbClr val="${theme.dk2}"/></a:dk2>
      <a:lt2><a:srgbClr val="${theme.lt2}"/></a:lt2>
      <a:accent1><a:srgbClr val="${theme.accent1}"/></a:accent1>
      <a:accent2><a:srgbClr val="${theme.accent2}"/></a:accent2>
      <a:accent3><a:srgbClr val="${theme.accent3}"/></a:accent3>
      <a:accent4><a:srgbClr val="${theme.accent4}"/></a:accent4>
      <a:accent5><a:srgbClr val="${theme.accent5}"/></a:accent5>
      <a:accent6><a:srgbClr val="${theme.accent6}"/></a:accent6>
      <a:hlink><a:srgbClr val="${theme.hlink}"/></a:hlink>
      <a:folHlink><a:srgbClr val="${theme.folHlink}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;
}

function seriesXml(index, { color } = {}) {
  return `<c:ser>
    <c:idx val="${index}"/>
    <c:order val="${index}"/>
    <c:tx><c:v>Series ${index + 1}</c:v></c:tx>
    ${color ? `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>` : ""}
    <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
    <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>
  </c:ser>`;
}

function chartXml({
  chartType = "Bar",
  customSeriesColors = [],
  seriesCount = 1,
  styleId = 2,
} = {}) {
  const series = Array.from({ length: seriesCount }, (_, index) => seriesXml(index, {
    color: customSeriesColors[index],
  })).join("\n");
  const typeNode = chartType === "Surface"
    ? `<c:surfaceChart>
        <c:wireframe val="0"/>
        ${series}
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:surfaceChart>`
    : `<c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        ${series}
        <c:gapWidth val="150"/>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>`;
  return `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:style val="${styleId}"/>
  <c:chart>
    <c:plotArea>
      ${typeNode}
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
        <c:majorGridlines/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;
}

function colorStyleXml(colors) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">
  ${colors.map((color) => `<a:srgbClr val="${color.replace(/^#/, "").toUpperCase()}"/>`).join("")}
</cs:colorStyle>`;
}

function workbookBytes({
  chartType,
  customColorStyle,
  customSeriesColors,
  seriesCount,
  styleId,
  theme = OFFICE_THEME,
} = {}) {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
  const sheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <drawing r:id="rId1"/>
</worksheet>`;
  const sheetRels = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
  const drawingXml = `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="5000000" cy="3000000"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`;
  const drawingRels = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  ${customColorStyle ? `<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>` : ""}
</Relationships>`;
  const archive = {
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
    "xl/worksheets/sheet1.xml": strToU8(sheetXml),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(sheetRels),
    "xl/drawings/drawing1.xml": strToU8(drawingXml),
    "xl/drawings/_rels/drawing1.xml.rels": strToU8(drawingRels),
    "xl/charts/chart1.xml": strToU8(chartXml({
      chartType,
      customSeriesColors,
      seriesCount,
      styleId,
    })),
    "xl/theme/theme1.xml": strToU8(themeXml(theme)),
  };
  if (customColorStyle) {
    archive["xl/charts/_rels/chart1.xml.rels"] = strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
</Relationships>`);
    archive["xl/charts/colors1.xml"] = strToU8(colorStyleXml(customColorStyle));
  }
  return zipSync(archive);
}

function cell(value) {
  return {
    is_empty: false,
    free() {},
    asNumber() {
      return typeof value === "number" ? value : Number(value);
    },
    asText() {
      return String(value);
    },
    toString() {
      return String(value);
    },
  };
}

function createSheet() {
  const values = {
    "1,0": "Jan",
    "1,1": 10,
    "2,0": "Feb",
    "2,1": 20,
    "3,0": "Mar",
    "3,1": 30,
  };
  const read = (row, col) => values[`${row},${col}`] ?? "";
  return {
    charts: [],
    chartsEx: [],
    free() {},
    getCalculatedValueAt(row, col) {
      return cell(read(row, col));
    },
    getFormattedValueAt(row, col) {
      return cell(read(row, col));
    },
    name: "Sheet1",
    visibility: "visible",
  };
}

function createWorkbook(rawCharts = []) {
  const sheet = createSheet();
  sheet.charts = rawCharts;
  return {
    chartsheets: [],
    getNamedRange() {
      return undefined;
    },
    getSheet() {
      return sheet;
    },
    getSheetByName() {
      return sheet;
    },
    sheetCount: 1,
    sheetNames: ["Sheet1"],
    sheetOrder: [{ index: 0, slotType: "worksheet" }],
  };
}

function loadChart(bytes, workbook = createWorkbook()) {
  const assets = core.parseWorkbookChartStyleAssets(bytes);
  const chartAssets = core.loadWorkbookChartAssets(
    workbook,
    assets,
    new Map([[0, 0]]),
    false,
  );
  return chartAssets.chartsByWorkbookSheetIndex[0]?.[0] ?? null;
}

test("built-in chart style 41 paints six grayscale series colours", () => {
  const chart = loadChart(workbookBytes({ seriesCount: 6, styleId: 41 }));
  assert.ok(chart);
  assert.equal(chart.chartStyleId, 41);
  assert.equal(chart.raw.hasCustomChartColorStyle, false);
  assert.deepEqual(
    chart.series.map((series) => series.color),
    ["#dddddd", "#9e9e9e", "#727272", "#d2d2d2", "#838383", "#c1c1c1"],
  );
  assert.deepEqual(chart.chartColorPalette, chart.series.map((series) => series.color));
});

test("built-in chart style 34 paints monochrome accent series and surface walls", () => {
  const chart = loadChart(workbookBytes({
    chartType: "Surface",
    seriesCount: 4,
    styleId: 34,
  }));
  assert.ok(chart);
  assert.equal(chart.chartType, "Surface");
  assert.deepEqual(
    chart.series.map((series) => series.color),
    ["#4f81bd", "#c0504d", "#9bbb59", "#8064a2"],
  );
  assert.equal(chart.chartColorPalette?.length, 6, "surface keeps the dedicated palette for the renderer");
  assert.equal(chart.floor?.fillColor, "#e7e7e7");
  assert.equal(chart.sideWall?.fillColor, "#e7e7e7");
  assert.equal(chart.backWall?.fillColor, "#e7e7e7");
  assert.equal(chart.floor?.lineColor, "#a8adb7");
  assert.equal(chart.valueAxis?.majorGridlineColor, "#000000");
  assert.equal(chart.categoryAxis?.majorGridlineColor, "#000000");
});

test("chart style 2 keeps theme accents and becomes grayscale with a grayscale theme", () => {
  const chart = loadChart(workbookBytes({
    seriesCount: 6,
    styleId: 2,
    theme: GRAYSCALE_THEME,
  }));
  assert.ok(chart);
  for (const series of chart.series) {
    assert.match(series.color ?? "", /^#[0-9a-f]{6}$/);
    const channels = [1, 3, 5].map((offset) => series.color?.slice(offset, offset + 2));
    assert.equal(new Set(channels).size, 1, `${series.color} is a neutral gray`);
  }
});

test("chart style 102 uses the built-in palette of its normalized style", () => {
  const chart = loadChart(workbookBytes({ seriesCount: 4, styleId: 102 }));
  assert.ok(chart);
  assert.equal(chart.chartStyleId, 102);
  assert.deepEqual(
    chart.series.map((series) => series.color),
    ["#4f81bd", "#c0504d", "#9bbb59", "#8064a2"],
  );
});

test("dark chart style variants keep the dark chart chrome", () => {
  const chart = loadChart(workbookBytes({ seriesCount: 4, styleId: 142 }));
  assert.ok(chart);
  assert.deepEqual(
    chart.series.map((series) => series.color),
    ["#4f81bd", "#c0504d", "#9bbb59", "#8064a2"],
  );
  assert.equal(chart.chartAreaFillColor, "#1f1f1f");
  assert.equal(chart.chartAreaBorderColor, "#1f1f1f");
  assert.equal(chart.textColor, "#f5f5f5");
  assert.equal(chart.titleColor, "#f5f5f5");
  assert.equal(chart.axisLabelColor, "#d9d9d9");
  assert.equal(chart.axisLineColor, "#8c8c8c");
});

test("a chartColorStyle part wins over the built-in palette", () => {
  const chart = loadChart(workbookBytes({
    customColorStyle: ["#ff0000", "#00ff00", "#0000ff"],
    seriesCount: 3,
    styleId: 41,
  }));
  assert.ok(chart);
  assert.equal(chart.raw.hasCustomChartColorStyle, true);
  assert.deepEqual(chart.chartColorPalette, ["#ff0000", "#00ff00", "#0000ff"]);
  assert.deepEqual(
    chart.series.map((series) => series.color),
    ["#ff0000", "#00ff00", "#0000ff"],
  );
});

test("an explicit series fill wins over the built-in palette", () => {
  const workbook = createWorkbook([{
    chartType: "BarClustered",
    series: [{
      name: "Series 1",
      shapeProperties: { solidFillHex: "FF00FF" },
      values: { formula: "Sheet1!$B$2:$B$4" },
    }],
  }]);
  const chart = loadChart(workbookBytes({ seriesCount: 1, styleId: 41 }), workbook);
  assert.ok(chart);
  assert.equal(chart.series[0]?.color, "#ff00ff");
  assert.equal(chart.chartColorPalette?.[0], "#f9f9f9");
});

test("built-in chart palettes survive an export/reload of the chart parts", () => {
  const source = workbookBytes({ seriesCount: 6, styleId: 41 });
  const first = loadChart(source);
  assert.ok(first);
  const exported = zipSync(unzipSync(source), { level: 6 });
  const reloaded = loadChart(exported);
  assert.ok(reloaded);
  assert.deepEqual(reloaded.series.map((series) => series.color), first.series.map((series) => series.color));
  assert.equal(reloaded.chartStyleId, 41);
  assert.equal(reloaded.raw.hasCustomChartColorStyle, false);
});

test("a saved workbook reloads built-in chart colours through the engine export path", async () => {
  const wasmPath = requireFromXlsxCore.resolve(
    "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
  );
  const wasmBytes = readFileSync(wasmPath);
  core.setWasmSource(
    wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
  );
  const wasm = await core.getSheetsWasmModule();
  const samplePath = new URL("../../apps/demo/public/samples/charts-images.xlsx", import.meta.url);
  const archive = unzipSync(new Uint8Array(readFileSync(samplePath)));
  const chartPath = "xl/charts/chart1.xml";
  const chartPart = new TextDecoder().decode(archive[chartPath]);
  archive[chartPath] = strToU8(chartPart.replace(
    "<chart>",
    '<style val="41"/><chart>',
  ));
  const sourceBytes = zipSync(archive, { level: 6 });

  const workbook = wasm.Workbook.fromBytes(sourceBytes);
  try {
    const assets = core.parseWorkbookImageAssets(sourceBytes);
    const chart = core.loadWorkbookChartAssets(
      workbook,
      assets,
      new Map(Array.from({ length: workbook.sheetCount }, (_, index) => [index, index])),
      false,
    ).chartsByWorkbookSheetIndex[0]?.find((entry) => entry.chartPath === chartPath);
    assert.ok(chart, "the patched chart parses");
    assert.equal(chart.chartStyleId, 41);
    assert.equal(chart.series[0]?.color, "#f9f9f9");

    const saved = workbook.saveXlsxBytes();
    const authored = core.mergeWorkbookSheetAuthoringAssets(saved, assets.archive);
    const merged = core.mergeWorkbookImageAssets(authored, assets, assets.sheetOrigins);
    const reloadedWorkbook = wasm.Workbook.fromBytes(merged);
    try {
      const reloadedAssets = core.parseWorkbookImageAssets(merged);
      const reloadedChart = core.loadWorkbookChartAssets(
        reloadedWorkbook,
        reloadedAssets,
        new Map(Array.from({ length: reloadedWorkbook.sheetCount }, (_, index) => [index, index])),
        false,
      ).chartsByWorkbookSheetIndex[0]?.find((entry) => entry.chartPath === chartPath);
      assert.ok(reloadedChart, "the exported chart reloads");
      assert.equal(reloadedChart.chartStyleId, 41);
      assert.equal(reloadedChart.series[0]?.color, "#f9f9f9");
      assert.equal(reloadedChart.raw.hasCustomChartColorStyle, false);
    } finally {
      reloadedWorkbook.free();
    }
  } finally {
    workbook.free();
  }
});

test("built-in palettes leave picture-filled series and points unpainted", () => {
  const archive = unzipSync(buildChartMediaWorkbookBytes());
  const chartPath = "xl/charts/chart1.xml";
  archive[chartPath] = strToU8(
    new TextDecoder().decode(archive[chartPath]).replace(
      "<chart>",
      '<style val="41"/><chart>',
    ),
  );
  const chart = loadChart(zipSync(archive, { level: 6 }));
  assert.ok(chart);
  assert.equal(chart.chartStyleId, 41);
  assert.equal(chart.series[0]?.color, undefined, "the picture fill is not replaced by the builtin color");
  assert.equal(chart.series[0]?.markerColor, undefined);
  assert.match(chart.series[0]?.fillImage ?? "", /^data:image\/png/);
  assert.equal(chart.series[0]?.lineColor, "#f9f9f9");
  const picturePoint = chart.series[0]?.dataPointStyles?.find((point) => point.fillImage);
  assert.ok(picturePoint, "the data-point picture fill is parsed");
  assert.equal(picturePoint.color, undefined);
  assert.deepEqual(chart.chartColorPalette, ["#f9f9f9"]);
});
