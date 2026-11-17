import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

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

const core = await import(
  new URL("../../packages/xlsx-core/dist/index.js", import.meta.url).href
);
const { strToU8, zipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);

function chartOnly(overrides = {}) {
  return {
    chartType: "Line",
    id: "chart-1",
    series: [{
      categories: [],
      categoriesRef: null,
      dataPoints: [],
      id: "series-1",
      values: [],
      valuesRef: null,
    }],
    workbookSheetIndex: 0,
    ...overrides,
  };
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

function createSheet(values, { hiddenColumns = [], hiddenRows = [] } = {}) {
  const read = (row, col) => values[`${row},${col}`] ?? "";
  return {
    free() {},
    getCalculatedValueAt(row, col) {
      return cell(read(row, col));
    },
    getFormattedValueAt(row, col) {
      return cell(read(row, col));
    },
    isColumnHidden(col) {
      return hiddenColumns.includes(col);
    },
    isRowHidden(row) {
      return hiddenRows.includes(row);
    },
    name: "Sheet1",
  };
}

function createWorkbook(sheet) {
  return {
    getNamedRange() {
      return undefined;
    },
    getSheet() {
      return sheet;
    },
    getSheetByName() {
      return sheet;
    },
  };
}

test("chart series resolution skips hidden rows and columns by default", () => {
  const values = {
    "1,0": "A",
    "1,1": 10,
    "1,2": 100,
    "2,0": "B",
    "2,1": 20,
    "2,2": 200,
    "3,0": "C",
    "3,1": 30,
    "3,2": 300,
  };
  const sheet = createSheet(values, { hiddenRows: [2] });
  const chart = chartOnly();

  const filtered = core.applyChartSeriesFormula(
    chart,
    0,
    '=SERIES("Series 1",Sheet1!$A$2:$A$4,Sheet1!$B$2:$B$4,1)',
    createWorkbook(sheet),
  );
  assert.deepEqual(filtered.series[0].categories, ["A", "C"]);
  assert.deepEqual(filtered.series[0].values, [10, 30]);

  const unfiltered = core.applyChartSeriesFormula(
    { ...chart, plotVisibleOnly: false },
    0,
    '=SERIES("Series 1",Sheet1!$A$2:$A$4,Sheet1!$B$2:$B$4,1)',
    createWorkbook(sheet),
  );
  assert.deepEqual(unfiltered.series[0].categories, ["A", "B", "C"]);
  assert.deepEqual(unfiltered.series[0].values, [10, 20, 30]);

  const hiddenColumnSheet = createSheet(values, { hiddenColumns: [2] });
  const hiddenColumnChart = core.applyChartSeriesFormula(
    chart,
    0,
    '=SERIES("Series 1",Sheet1!$A$2:$A$4,Sheet1!$B$2:$C$4,1)',
    createWorkbook(hiddenColumnSheet),
  );
  assert.deepEqual(hiddenColumnChart.series[0].values, [10, 20, 30]);
});

function multiLevelWorkbookBytes() {
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
</Relationships>`;
  const chartXml = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>Series 1</c:v></c:tx>
          <c:cat>
            <c:multiLvlStrRef>
              <c:f>Sheet1!$A$2:$B$4</c:f>
              <c:multiLvlStrCache>
                <c:ptCount val="2"/>
                <c:lvl>
                  <c:pt idx="0"><c:v>2023</c:v></c:pt>
                  <c:pt idx="1"><c:v>2023</c:v></c:pt>
                </c:lvl>
                <c:lvl>
                  <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                  <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                </c:lvl>
              </c:multiLvlStrCache>
            </c:multiLvlStrRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$C$2:$C$3</c:f>
              <c:numCache>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:dLbls>
          <c:dLblPos val="outEnd"/>
          <c:numFmt formatCode="0.0%" sourceLinked="0"/>
          <c:showVal val="1"/>
        </c:dLbls>
      </c:lineChart>
      <c:dateAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
        <c:numFmt formatCode="m/d/yyyy" sourceLinked="0"/>
        <c:majorTimeUnit val="days"/>
      </c:dateAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/><c:logBase val="10"/></c:scaling>
        <c:axPos val="l"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/_rels/workbook.xml.rels": strToU8(workbookRels),
    "xl/worksheets/sheet1.xml": strToU8(sheetXml),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(sheetRels),
    "xl/drawings/drawing1.xml": strToU8(drawingXml),
    "xl/drawings/_rels/drawing1.xml.rels": strToU8(drawingRels),
    "xl/charts/chart1.xml": strToU8(chartXml),
  });
}

test("multi-level chart categories keep every cache level", () => {
  const imageAssets = core.parseWorkbookChartStyleAssets(multiLevelWorkbookBytes());
  const worksheet = {
    charts: [],
    chartsEx: [],
    free() {},
    name: "Sheet1",
    visibility: "visible",
  };
  const workbook = {
    chartsheets: [],
    getNamedRange() {
      return undefined;
    },
    getSheet() {
      return worksheet;
    },
    getSheetByName() {
      return undefined;
    },
    sheetNames: ["Sheet1"],
    sheetOrder: [{ index: 0, slotType: "worksheet" }],
    sheetCount: 1,
  };
  const assets = core.loadWorkbookChartAssets(
    workbook,
    imageAssets,
    new Map([[0, 0]]),
  );
  const chart = assets.chartsByWorkbookSheetIndex[0]?.[0];
  assert.ok(chart, "chart is resolved from the drawing");
  assert.deepEqual(chart.series[0].categories, ["2023 / Jan", "2023 / Feb"]);
  assert.deepEqual(chart.series[0].values, [10, 20]);
  assert.equal(chart.categoryAxis?.axisType, "date");
  assert.equal(chart.valueAxis?.logBase, 10);
  assert.equal(chart.dataLabels?.position, "outEnd");
  assert.equal(chart.dataLabels?.numberFormat?.formatCode, "0.0%");
  assert.equal(chart.dataLabels?.showValue, true);
});

function createReferencedChart() {
  const series = {
    categories: ["cached-a", "cached-b"],
    categoriesRef: { formula: "Sheet1!$A$2:$A$3" },
    dataPoints: [],
    id: "series-1",
    name: "Series 1",
    values: [1, 2],
    valuesRef: { formula: "Sheet1!$B$2:$B$3" },
  };
  return chartOnly({
    series: [series],
    typeGroups: [{ chartType: "Line", series: [series] }],
  });
}

function createValueSheet(initialValues) {
  const values = { ...initialValues };
  const sheet = {
    free() {},
    getCalculatedValueAt(row, col) {
      return cell(values[`${row},${col}`] ?? "");
    },
    getFormattedValueAt(row, col) {
      return cell(values[`${row},${col}`] ?? "");
    },
    isColumnHidden() {
      return false;
    },
    isRowHidden() {
      return false;
    },
    name: "Sheet1",
  };
  return {
    set(row, col, value) {
      values[`${row},${col}`] = value;
    },
    sheet,
  };
}

test("chart data refreshes from an edited workbook model", () => {
  const chart = createReferencedChart();
  const { sheet, set } = createValueSheet({
    "1,0": "A",
    "1,1": 10,
    "2,0": "B",
    "2,1": 20,
  });
  const workbook = createWorkbook(sheet);

  const first = core.refreshChartDataFromWorkbook(workbook, [[chart]], []);
  assert.deepEqual(first.chartsByWorkbookSheetIndex[0][0].series[0].categories, ["A", "B"]);
  assert.deepEqual(first.chartsByWorkbookSheetIndex[0][0].series[0].values, [10, 20]);
  assert.deepEqual(first.chartsByWorkbookSheetIndex[0][0].typeGroups?.[0].series[0].values, [10, 20]);

  set(1, 1, 30);
  set(2, 1, 40);
  const second = core.refreshChartDataFromWorkbook(workbook, [[chart]], []);
  assert.deepEqual(second.chartsByWorkbookSheetIndex[0][0].series[0].values, [30, 40]);
  assert.deepEqual(second.chartsByWorkbookSheetIndex[0][0].typeGroups?.[0].series[0].values, [30, 40]);
});

test("chartsheet charts refresh from the edited workbook model too", () => {
  const chart = { ...createReferencedChart(), workbookSheetIndex: -1 };
  const chartsheet = {
    chartIds: [chart.id],
    charts: [chart],
    id: "chartsheet-0",
    index: 0,
    name: "Chart 1",
  };
  const { sheet, set } = createValueSheet({
    "1,0": "A",
    "1,1": 10,
    "2,0": "B",
    "2,1": 20,
  });
  const workbook = createWorkbook(sheet);

  assert.deepEqual(
    core.refreshChartDataFromWorkbook(workbook, [], [chartsheet]).chartsheets[0].charts?.[0].series[0].values,
    [10, 20],
  );
  set(2, 1, 25);
  const refreshed = core.refreshChartDataFromWorkbook(workbook, [], [chartsheet]).chartsheets[0];
  assert.deepEqual(refreshed.charts?.[0].series[0].values, [10, 25]);
  assert.deepEqual(refreshed.chartIds, [chart.id]);
});

function createRespondingWorker(result) {
  const listeners = new Map();
  const worker = {
    addEventListener(type, listener) {
      listeners.set(type, [...(listeners.get(type) ?? []), listener]);
    },
    removeEventListener(type, listener) {
      listeners.set(type, (listeners.get(type) ?? []).filter((entry) => entry !== listener));
    },
    postMessage(message) {
      this.lastMessage = message;
      queueMicrotask(() => {
        const currentResult = typeof result === "function" ? result() : result;
        for (const listener of listeners.get("message") ?? []) {
          listener({ data: { id: message.id, result: currentResult, success: true } });
        }
      });
    },
    terminate() {},
  };
  return worker;
}

test("worker client refreshes chart data without reparsing the source bytes", async () => {
  const chart = createReferencedChart();
  let values = [10, 20];
  const worker = createRespondingWorker(() => ({
    chartsByWorkbookSheetIndex: [[{
      ...chart,
      series: chart.series.map((series) => ({ ...series, values: [...values] })),
    }]],
    chartsheets: [],
    tabs: [],
  }));
  const client = new core.XlsxWorkerClient({
    createWorker: () => worker,
    useLegacyDefaultWasmSource: false,
  });
  try {
    const first = await client.refreshCharts(true);
    assert.equal(worker.lastMessage?.type, "refreshCharts");
    assert.equal(worker.lastMessage?.payload?.showHiddenSheets, true);
    assert.deepEqual(first.chartsByWorkbookSheetIndex[0][0].series[0].values, [10, 20]);

    values = [30, 40];
    const second = await client.refreshCharts();
    assert.deepEqual(second.chartsByWorkbookSheetIndex[0][0].series[0].values, [30, 40]);
  } finally {
    client.dispose();
  }
});
