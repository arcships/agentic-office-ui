import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import test from "node:test";
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

const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromXlsxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromXlsxCore.resolve("@arcships/xlsx-core")).href
);

const samples = new URL("../../apps/demo/public/samples/", import.meta.url);
const wasmPath = requireFromXlsxCore.resolve(
  "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
core.setWasmSource(
  wasmBytes.buffer.slice(wasmBytes.byteOffset, wasmBytes.byteOffset + wasmBytes.byteLength),
);
const wasm = await core.getSheetsWasmModule();

const CHART_EX_CONTENT_TYPE = "application/vnd.ms-office.chartex+xml";
const CLASSIC_CHART_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CHART_STYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const CHART_COLOR_STYLE_REL_TYPE =
  "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
const CHART_EX_REL_TYPE = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
const CHART_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

function sampleBytes(name) {
  return new Uint8Array(readFileSync(new URL(name, samples)));
}

function loadAssets(name = "charts-images.xlsx") {
  const bytes = sampleBytes(name);
  return {
    assets: core.parseWorkbookImageAssets(bytes),
    bytes,
    workbook: wasm.Workbook.fromBytes(bytes),
  };
}

function visibleMap(count) {
  return new Map(Array.from({ length: count }, (_, index) => [index, index]));
}

function chartAssetsFor(workbook, assets) {
  return core.loadWorkbookChartAssets(
    workbook,
    assets,
    visibleMap(workbook.sheetCount),
    false,
  );
}

function findChartByPath(chartAssets, chartPath) {
  for (const sheetCharts of chartAssets.chartsByWorkbookSheetIndex) {
    const match = sheetCharts.find((chart) => chart.chartPath === chartPath);
    if (match) return match;
  }
  for (const chartsheet of chartAssets.chartsheets) {
    const match = chartsheet.charts?.find((chart) => chart.chartPath === chartPath);
    if (match) return match;
  }
  return null;
}

/** Saves through the engine, then merges authored parts exactly like export. */
function exportWithAuthoring(workbook, assets) {
  const saved = workbook.saveXlsxBytes();
  const authored = core.mergeWorkbookSheetAuthoringAssets(saved, assets.archive);
  return core.mergeWorkbookImageAssets(authored, assets, assets.sheetOrigins);
}

function parsePart(archive, partPath) {
  const document = new DOMParser().parseFromString(strFromU8(archive[partPath]), "application/xml");
  assert.equal(document.getElementsByTagName("parsererror").length, 0, `${partPath} is well-formed XML`);
  return document;
}

function localElements(parent, localName) {
  return [...parent.getElementsByTagName("*")].filter((node) => node.localName === localName);
}

function childNames(element) {
  return [...element.childNodes]
    .filter((node) => node.nodeType === 1)
    .map((node) => node.localName);
}

function firstChild(element, localName) {
  return [...element.childNodes].find((node) => node.nodeType === 1 && node.localName === localName);
}

function contentTypes(archive) {
  const document = parsePart(archive, "[Content_Types].xml");
  return {
    defaults: localElements(document, "Default").map((node) => ({
      contentType: node.getAttribute("ContentType"),
      extension: node.getAttribute("Extension"),
    })),
    overrides: localElements(document, "Override").map((node) => ({
      contentType: node.getAttribute("ContentType"),
      partName: node.getAttribute("PartName"),
    })),
  };
}

function contentTypeFor(archive, partPath) {
  return contentTypes(archive).overrides.find((entry) => entry.partName === `/${partPath}`)
    ?.contentType ?? null;
}

function resolveRelsTarget(relsPath, target) {
  if (target.startsWith("/")) return target.slice(1);
  // A `.rels` part lives in `_rels/`; relative targets resolve against the
  // directory of the part that owns it.
  const ownerDirectory = relsPath.startsWith("_rels/")
    ? ""
    : relsPath.replace(/\/_rels\/[^/]+\.rels$/, "").replace(/^\/_rels\/[^/]+\.rels$/, "");
  const segments = ownerDirectory ? ownerDirectory.split("/") : [];
  for (const segment of target.split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") segments.pop();
    else segments.push(segment);
  }
  return segments.join("/");
}

function relationships(archive, partPath) {
  const fileName = partPath.split("/").pop();
  const relsPath = `${partPath.slice(0, -fileName.length)}_rels/${fileName}.rels`;
  if (!archive[relsPath]) return { entries: [], relsPath };
  const document = parsePart(archive, relsPath);
  return {
    entries: localElements(document, "Relationship").map((node) => ({
      id: node.getAttribute("Id"),
      target: resolveRelsTarget(relsPath, node.getAttribute("Target") ?? ""),
      type: node.getAttribute("Type"),
    })),
    relsPath,
  };
}

function assertEveryRelationshipTargetExists(archive) {
  for (const path of Object.keys(archive)) {
    if (!path.endsWith(".rels")) continue;
    const document = parsePart(archive, path);
    for (const node of localElements(document, "Relationship")) {
      if (node.getAttribute("TargetMode") === "External") continue;
      const target = node.getAttribute("Target") ?? "";
      if (!target || target.startsWith("#")) continue;
      const resolved = resolveRelsTarget(path, target);
      assert.ok(archive[resolved], `${path} target ${target} exists (${resolved})`);
    }
  }
}

function insertSeriesChart(assets, { range = { start: { row: 0, col: 0 }, end: { row: 4, col: 2 } } } = {}) {
  const inserted = core.insertWorkbookChart(assets, 0, {
    family: "bar",
    range,
    sheetName: "Dashboard",
    title: "Quarterly revenue",
  });
  assert.ok(inserted, "chart is inserted");
  return inserted;
}

function seriesSummary(chart) {
  return chart.series.map((series) => ({
    categories: series.categories,
    color: series.color,
    name: series.name,
    values: series.values,
  }));
}

test("chart type catalog covers every gallery family with unique ids", () => {
  const options = core.chartTypeOptions;
  assert.ok(options.length >= 57, `catalog exposes every classic and modern option (${options.length})`);
  assert.equal(new Set(options.map((option) => option.id)).size, options.length, "option ids are unique");
  assert.deepEqual(
    [...new Set(core.chartTypeFamilies().map((family) => family.family))],
    [
      "Column",
      "Bar",
      "Line",
      "Area",
      "Pie",
      "XY (Scatter)",
      "Stock",
      "Surface",
      "Radar",
      "Combo",
      "Treemap",
      "Sunburst",
      "Histogram",
      "Box & Whisker",
      "Waterfall",
      "Funnel",
      "Map",
    ],
    "families keep the gallery order",
  );
  for (const id of ["ColumnClustered", "Pie3D", "Treemap", "Sunburst", "Histogram", "Pareto", "BoxWhisker", "Waterfall", "Funnel", "RegionMap"]) {
    assert.ok(core.resolveChartTypeOption(id), `${id} resolves`);
  }
  assert.equal(core.resolveChartTypeOption("Nope"), null);

  const modern = core.chartTypeOptions.filter((option) => option.layout);
  assert.deepEqual(
    modern.map((option) => [option.id, option.layout]),
    [
      ["Treemap", "treemap"],
      ["Sunburst", "sunburst"],
      ["Histogram", "clusteredColumn"],
      ["Pareto", "paretoLine"],
      ["BoxWhisker", "boxWhisker"],
      ["Waterfall", "waterfall"],
      ["Funnel", "funnel"],
      ["RegionMap", "regionMap"],
    ],
    "chartEx options carry their layout ids",
  );
});

test("chart type options are disabled when the chart cannot express them", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets);
    const chartAssets = chartAssetsFor(workbook, assets);
    const chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.ok(chart);
    assert.equal(core.chartTypeUnavailable(chart, core.resolveChartTypeOption("StockHLC")), true);
    assert.equal(core.chartTypeUnavailable(chart, core.resolveChartTypeOption("Pie3D")), false);
    assert.equal(core.chartTypeUnavailable({ ...chart, editable: false }, core.resolveChartTypeOption("Pie")), true);
    assert.equal(core.chartTypeUnavailable({ ...chart, chartPath: undefined }, core.resolveChartTypeOption("Pie")), true);
    assert.equal(core.matchesChartType(chart, core.resolveChartTypeOption("ColumnClustered")), true);
    assert.equal(core.matchesChartType(chart, core.resolveChartTypeOption("BarClustered")), false);
  } finally {
    workbook.free();
  }
});

test("classic → classic rewrite keeps series names, colours and data labels", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets);
    let chartAssets = chartAssetsFor(workbook, assets);
    const chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.ok(chart);

    assert.equal(
      core.updateWorkbookChartDefinition(assets, chartAssets, chart.id, {
        dataLabels: { position: "outEnd", showValue: true },
        title: "Revenue and exceptions",
      }),
      true,
      "data labels are written before the switch",
    );

    const before = seriesSummary(chart);
    assert.equal(
      core.changeWorkbookChartType(assets, chartAssets, chart.id, "Pie3D"),
      true,
      "the chart is rewritten as a 3-D pie",
    );

    const archive = assets.archive;
    const chartXml = strFromU8(archive[inserted.chartPath]);
    assert.match(chartXml, /<c:pie3DChart>/, "plot area uses pie3DChart");
    assert.doesNotMatch(chartXml, /<c:barChart>/, "the previous plot area is gone");
    assert.match(chartXml, /<c:view3D>/, "3-D charts carry a view3D node");
    assert.match(chartXml, /<c:dLbls>[\s\S]*<c:showVal val="1"\/>/, "chart data labels survive");
    assert.equal(
      (chartXml.match(/<c:ser>/g) ?? []).length,
      before.length,
      "every series is written",
    );
    assert.match(chartXml, /<c:tx><c:strRef><c:f>'?Dashboard'?!\$B\$1<\/c:f><\/c:strRef><\/c:tx>|Revenue/, "series names survive");
    assert.match(chartXml, /<a:srgbClr val="4F81BD"\/>/, "series colour survives");

    // Excel rejects plot areas whose children are out of schema order.
    const chartDocument = parsePart(archive, inserted.chartPath);
    const pieNode = localElements(chartDocument, "pie3DChart")[0];
    assert.deepEqual(childNames(pieNode), ["varyColors", "ser", "ser", "dLbls"], "pie3DChart keeps schema order");
    const seriesChildren = childNames(localElements(pieNode, "ser")[0]);
    assert.deepEqual(
      seriesChildren,
      ["idx", "order", "tx", "spPr", "cat", "val"],
      "series children keep schema order",
    );

    chartAssets = chartAssetsFor(workbook, assets);
    const next = findChartByPath(chartAssets, inserted.chartPath);
    assert.ok(next, "rewritten chart reloads");
    assert.equal(next.chartType, "Pie3D");
    assert.equal(next.is3d, true);
    assert.deepEqual(next.series.map((series) => series.name), before.map((series) => series.name));
    assert.deepEqual(next.series.map((series) => series.values), before.map((series) => series.values));
    assert.equal(next.dataLabels?.showValue, true, "data labels are parsed from the rewritten part");
  } finally {
    workbook.free();
  }
});

test("classic → chartEx migration writes chartEx parts, styles and markup-compatibility", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets);
    const chartAssets = chartAssetsFor(workbook, assets);
    const chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.ok(chart);
    const before = seriesSummary(chart);
    assert.equal(
      core.updateWorkbookChartDefinition(assets, chartAssets, chart.id, {
        dataLabels: { position: "outEnd", showValue: true },
      }),
      true,
      "data labels are written before the migration",
    );

    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "Treemap"), true);

    const archive = assets.archive;
    const chartXml = strFromU8(archive[inserted.chartPath]);
    assert.match(chartXml, /^<\?xml[^>]*\?><cx:chartSpace /, "the part is a chartEx chartSpace");
    assert.match(chartXml, /<cx:chartData>[\s\S]*<cx:strDim type="cat">/, "categories use strDim");
    assert.match(chartXml, /<cx:numDim type="size">/, "treemap values use the size dimension");
    assert.match(chartXml, /<cx:series [^>]*layoutId="treemap"/, "series declare the treemap layout");
    assert.match(
      chartXml,
      /<cx:dataLabels><cx:dLblPos val="outEnd"\/><cx:showVal val="1"\/><\/cx:dataLabels>/,
      "chart data labels move to the chartEx series",
    );
    assert.equal(contentTypeFor(archive, inserted.chartPath), CHART_EX_CONTENT_TYPE);

    const drawingRels = relationships(archive, inserted.drawingPath);
    const chartRelationship = drawingRels.entries.find((entry) => entry.target === inserted.chartPath);
    assert.ok(chartRelationship, "the drawing still references the chart part");
    assert.equal(chartRelationship.type, CHART_EX_REL_TYPE, "the relationship type becomes chartEx");
    const drawing = parsePart(archive, inserted.drawingPath);
    const alternate = localElements(drawing, "AlternateContent")[0];
    assert.ok(alternate, "the graphic frame is wrapped in mc:AlternateContent");
    assert.equal(alternate.parentNode.localName, "oneCellAnchor", "the wrapper replaces the graphic frame in the anchor");
    assert.deepEqual(
      childNames(alternate),
      ["Choice", "Fallback"],
      "AlternateContent carries a choice and a fallback",
    );
    assert.equal(localElements(alternate, "Choice")[0]?.getAttribute("Requires"), "cx1");
    assert.ok(localElements(alternate, "Fallback")[0], "mc:Fallback renders the legacy placeholder");
    assert.equal(
      localElements(drawing, "graphicData")[0]?.getAttribute("uri"),
      "http://schemas.microsoft.com/office/drawing/2014/chartex",
      "graphicData points at the chartEx namespace",
    );

    const styleRelationships = relationships(archive, inserted.chartPath);
    const stylePart = styleRelationships.entries.find((entry) => entry.type === CHART_STYLE_REL_TYPE);
    const colorStylePart = styleRelationships.entries.find((entry) => entry.type === CHART_COLOR_STYLE_REL_TYPE);
    assert.ok(stylePart && colorStylePart, "chartStyle and chartColorStyle relationships exist");
    assert.ok(archive[stylePart.target], "chartStyle part exists");
    assert.ok(archive[colorStylePart.target], "chartColorStyle part exists");
    assert.equal(contentTypeFor(archive, stylePart.target), "application/vnd.ms-office.chartstyle+xml");
    assert.equal(contentTypeFor(archive, colorStylePart.target), "application/vnd.ms-office.chartcolorstyle+xml");

    // Switching between modern layouts must not duplicate the style parts.
    const stylePartsBefore = Object.keys(archive).filter((path) => path.includes("chartStyle")).length;
    const nextCharts = chartAssetsFor(workbook, assets);
    const treemap = findChartByPath(nextCharts, inserted.chartPath);
    assert.equal(core.changeWorkbookChartType(assets, nextCharts, treemap.id, "Sunburst"), true);
    assert.equal(
      Object.keys(archive).filter((path) => path.includes("chartStyle")).length,
      stylePartsBefore,
      "modern style parts are not duplicated",
    );

    const sunburstAssets = chartAssetsFor(workbook, assets);
    const sunburst = findChartByPath(sunburstAssets, inserted.chartPath);
    assert.ok(sunburst, "chartEx chart is read from the archive");
    assert.equal(sunburst.chartType, "Sunburst");
    assert.equal(sunburst.dataLabels?.showValue, true, "chartEx data labels are parsed back");
    assert.equal(sunburst.chartExLayout, "sunburst");
    assert.deepEqual(
      sunburst.series.map((series) => [series.name, series.values]),
      [[before[0].name, before[0].values]],
      "treemap/sunburst keep the plotted series and its values",
    );
    assert.deepEqual(
      sunburst.raw.chartExSourceSeries.map((series) => series.name),
      before.map((series) => series.name),
      "every declared series stays recoverable",
    );
  } finally {
    workbook.free();
  }
});

test("chartEx → classic migration clears the wrapper and content type", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets);
    let chartAssets = chartAssetsFor(workbook, assets);
    let chart = findChartByPath(chartAssets, inserted.chartPath);
    const before = seriesSummary(chart);
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "Treemap"), true);

    chartAssets = chartAssetsFor(workbook, assets);
    chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.equal(chart.id.startsWith("chart-ex-"), true, "modern charts use the chartEx id");
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "ColumnClustered"), true);

    const archive = assets.archive;
    const chartXml = strFromU8(archive[inserted.chartPath]);
    assert.match(chartXml, /^<\?xml[^>]*\?><c:chartSpace /, "the part is a classic chartSpace again");
    assert.match(chartXml, /<c:barChart>/, "plot area uses barChart");
    assert.equal(contentTypeFor(archive, inserted.chartPath), CLASSIC_CHART_CONTENT_TYPE);

    const drawingRels = relationships(archive, inserted.drawingPath);
    assert.equal(
      drawingRels.entries.find((entry) => entry.target === inserted.chartPath)?.type,
      CHART_REL_TYPE,
      "the drawing relationship returns to the classic type",
    );
    const drawing = parsePart(archive, inserted.drawingPath);
    assert.equal(localElements(drawing, "AlternateContent").length, 0, "mc:AlternateContent is unwrapped");
    assert.equal(
      localElements(drawing, "graphicData")[0]?.getAttribute("uri"),
      "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "graphicData points back at the classic namespace",
    );

    const nextAssets = chartAssetsFor(workbook, assets);
    const next = findChartByPath(nextAssets, inserted.chartPath);
    assert.equal(next.chartType, "ColumnClustered");
    assert.equal(next.series.length, before.length, "both series are restored");
    assert.deepEqual(
      next.series.map((series) => [series.name, series.color, series.values]),
      before.map((series) => [series.name, series.color, series.values]),
      "names, colours and values survive the round trip",
    );
  } finally {
    workbook.free();
  }
});

test("column → 3-D pie → treemap → clustered column keeps series, undoable steps and export", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets);
    let chartAssets = chartAssetsFor(workbook, assets);
    let chart = findChartByPath(chartAssets, inserted.chartPath);
    const original = seriesSummary(chart);
    const expectedTypes = ["ColumnClustered", "Pie3D", "Treemap", "ColumnClustered"];

    for (const type of expectedTypes) {
      assert.equal(
        core.changeWorkbookChartType(assets, chartAssets, chart.id, type),
        true,
        `${type} is applied to ${chart.id}`,
      );
      chartAssets = chartAssetsFor(workbook, assets);
      chart = findChartByPath(chartAssets, inserted.chartPath);
      assert.ok(chart, `${type} chart reloads`);
      assert.equal(chart.chartType, type);
      assert.deepEqual(
        seriesSummary(chart).map((series) => series.name),
        original.map((series) => series.name).slice(0, chart.series.length),
        `${type} keeps the series names`,
      );
      assert.deepEqual(
        seriesSummary(chart).map((series) => series.color),
        original.map((series) => series.color).slice(0, chart.series.length),
        `${type} keeps the series colours`,
      );
    }

    assert.equal(chart.series.length, original.length, "the final clustered column has both series again");
    assert.deepEqual(seriesSummary(chart), original, "the final chart is identical to the original model");

    const exported = exportWithAuthoring(workbook, assets);
    const archive = unzipSync(exported);
    assertEveryRelationshipTargetExists(archive);
    assert.equal(contentTypeFor(archive, inserted.chartPath), CLASSIC_CHART_CONTENT_TYPE);

    const reloaded = wasm.Workbook.fromBytes(exported);
    try {
      const reloadedCharts = core.loadWorkbookChartAssets(
        reloaded,
        core.parseWorkbookImageAssets(exported),
        visibleMap(reloaded.sheetCount),
        false,
      );
      const reloadedChart = findChartByPath(reloadedCharts, inserted.chartPath);
      assert.ok(reloadedChart, "the exported chart reloads from the package");
      assert.equal(reloadedChart.chartType, "ColumnClustered");
      assert.deepEqual(seriesSummary(reloadedChart), original);
    } finally {
      reloaded.free();
    }
  } finally {
    workbook.free();
  }
});

test("an exported chartEx chart keeps its parts, styles and data through reload", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets);
    const chartAssets = chartAssetsFor(workbook, assets);
    const chart = findChartByPath(chartAssets, inserted.chartPath);
    const before = seriesSummary(chart);
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "Funnel"), true);

    const exported = exportWithAuthoring(workbook, assets);
    const archive = unzipSync(exported);
    assertEveryRelationshipTargetExists(archive);
    assert.equal(contentTypeFor(archive, inserted.chartPath), CHART_EX_CONTENT_TYPE);
    assert.match(strFromU8(archive[inserted.chartPath]), /layoutId="funnel"/);

    const reloaded = wasm.Workbook.fromBytes(exported);
    try {
      const reloadedCharts = core.loadWorkbookChartAssets(
        reloaded,
        core.parseWorkbookImageAssets(exported),
        visibleMap(reloaded.sheetCount),
        false,
      );
      const reloadedChart = findChartByPath(reloadedCharts, inserted.chartPath);
      assert.ok(reloadedChart, "chartEx charts survive an export/reload cycle");
      assert.equal(reloadedChart.chartType, "Funnel");
      assert.equal(reloadedChart.chartExLayout, "funnel");
      assert.equal(reloadedChart.series[0].name, before[0].name);
      assert.deepEqual(reloadedChart.series[0].values, before[0].values);
    } finally {
      reloaded.free();
    }
  } finally {
    workbook.free();
  }
});

test("histogram and pareto binning is written and not double-applied on reload", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets, {
      range: { start: { row: 0, col: 0 }, end: { row: 4, col: 1 } },
    });
    let chartAssets = chartAssetsFor(workbook, assets);
    let chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "Histogram"), true);
    const histogramXml = strFromU8(assets.archive[inserted.chartPath]);
    assert.match(histogramXml, /<cx:binning intervalClosed="r"><cx:binCount val="10"\/><\/cx:binning>/);

    chartAssets = chartAssetsFor(workbook, assets);
    chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.equal(chart.chartType, "ColumnClustered");
    assert.equal(chart.chartExLayout, "clusteredColumn");
    const histogramCounts = chart.series[0].values;
    assert.equal(histogramCounts.length, 11, "10 bins span 120000…210000");
    assert.equal(
      histogramCounts.reduce((total, value) => total + (value ?? 0), 0),
      4,
      "every source value falls into exactly one bin",
    );

    const histogramExport = exportWithAuthoring(workbook, assets);
    const reloadedHistogram = wasm.Workbook.fromBytes(histogramExport);
    try {
      const reloadedCharts = core.loadWorkbookChartAssets(
        reloadedHistogram,
        core.parseWorkbookImageAssets(histogramExport),
        visibleMap(reloadedHistogram.sheetCount),
        false,
      );
      const reloadedChart = findChartByPath(reloadedCharts, inserted.chartPath);
      assert.equal(reloadedChart.chartType, "ColumnClustered");
      assert.deepEqual(
        reloadedChart.series[0].values,
        histogramCounts,
        "reloading does not bin the counts a second time",
      );
    } finally {
      reloadedHistogram.free();
    }

    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "Pareto"), true);
    chartAssets = chartAssetsFor(workbook, assets);
    chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.equal(chart.chartExLayout, "paretoLine");
    assert.deepEqual(chart.typeGroups?.map((group) => group.chartType), ["ColumnClustered", "Line"]);
    assert.deepEqual(chart.series[1].values.slice(0, 4), [25, 50, 75, 100], "the pareto line is cumulative");
    assert.equal(
      chart.series[0].values.filter((value) => value !== 0).length,
      4,
      "the pareto histogram keeps one count per source value",
    );

    const paretoExport = exportWithAuthoring(workbook, assets);
    const reloadedPareto = wasm.Workbook.fromBytes(paretoExport);
    try {
      const reloadedCharts = core.loadWorkbookChartAssets(
        reloadedPareto,
        core.parseWorkbookImageAssets(paretoExport),
        visibleMap(reloadedPareto.sheetCount),
        false,
      );
      const reloadedChart = findChartByPath(reloadedCharts, inserted.chartPath);
      assert.equal(reloadedChart.chartExLayout, "paretoLine");
      assert.deepEqual(reloadedChart.series[1].values.slice(0, 4), [25, 50, 75, 100]);
    } finally {
      reloadedPareto.free();
    }
  } finally {
    workbook.free();
  }
});

test("combination charts keep every plot group's series when switching back", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets);
    let chartAssets = chartAssetsFor(workbook, assets);
    let chart = findChartByPath(chartAssets, inserted.chartPath);
    const original = seriesSummary(chart);
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "ColumnLine"), true);

    chartAssets = chartAssetsFor(workbook, assets);
    chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.match(strFromU8(assets.archive[inserted.chartPath]), /<c:barChart>/, "the first group is a column chart");
    assert.match(strFromU8(assets.archive[inserted.chartPath]), /<c:lineChart>/, "the last group is a line chart");
    assert.deepEqual(
      chart.typeGroups?.map((group) => group.chartType),
      ["ColumnClustered", "Line"],
      "the reloaded chart keeps both plot groups",
    );
    assert.equal(
      core.matchesChartType(chart, core.resolveChartTypeOption("ColumnLine")),
      true,
      "the gallery marks the combination entry as active",
    );

    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "ColumnClustered"), true);
    chartAssets = chartAssetsFor(workbook, assets);
    chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.deepEqual(
      seriesSummary(chart),
      original,
      "both plot groups' series return to the single column group",
    );
  } finally {
    workbook.free();
  }
});

test("switching chart type rejects unknown options and read-only charts", () => {
  const { assets, workbook } = loadAssets();
  try {
    const inserted = insertSeriesChart(assets);
    const chartAssets = chartAssetsFor(workbook, assets);
    const chart = findChartByPath(chartAssets, inserted.chartPath);
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "NotAChart"), false);
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, "missing-chart", "Pie3D"), false);
    const readOnlyChart = { ...chart, editable: false };
    chartAssets.chartsByWorkbookSheetIndex = chartAssets.chartsByWorkbookSheetIndex.map((sheetCharts) => (
      sheetCharts.map((entry) => (entry.id === chart.id ? readOnlyChart : entry))
    ));
    assert.equal(core.changeWorkbookChartType(assets, chartAssets, chart.id, "Pie3D"), false);
  } finally {
    workbook.free();
  }
});
