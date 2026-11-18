import { strToU8 } from "fflate";

import type { XlsxChart, XlsxChartDataLabels, XlsxChartSeries } from "../types";
import type { WorkbookImageAssets } from "../images";

import { relsPathForDocument } from "../images";
import { ensureContentTypeOverride, escapeXmlText } from "../authoring/sheet-authoring";

import {
  CHART_COLOR_STYLE_REL_TYPE,
  CHART_EX_REL_TYPE,
  CHART_NS,
  CHART_REL_TYPE,
  CHART_STYLE_REL_TYPE,
  DRAWINGML_NS,
  PKG_REL_NS,
} from "./chart-series";

import { normalizeHexColor } from "./chart-colors";

import { findPrimaryChartTypeNode, type WorkbookChartAssets, type WorkbookChartOrigin } from "./chart-types";
import { parseChartDataLabelsFromXml } from "./chart-parser";

import { resolveChartTypeOption, type ChartTypeOption } from "./chart-type-options";

import {
  getFirstLocalChild,
  getFirstLocalDescendant,
  getLocalChildren,
  getLocalDescendants,
  normalizeArchivePath,
  parseXml,
  readArchiveText,
  resolveRelationshipPath,
  serializeXml,
  setLeafValue,
} from "./chart-xml-utils";

const CHART_EX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const CHART_EX_CHOICE_NS = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
const CHART_STYLE_NS = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const MARKUP_COMPATIBILITY_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const CHART_EX_CONTENT_TYPE = "application/vnd.ms-office.chartex+xml";
const CLASSIC_CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const PIE_LIKE_NODES = new Set(["pieChart", "pie3DChart", "doughnutChart", "ofPieChart"]);
const SURFACE_NODES = new Set(["surfaceChart", "surface3DChart"]);
const MARKER_NODES = new Set(["lineChart", "line3DChart", "scatterChart", "radarChart", "stockChart"]);
const XY_NODES = new Set(["scatterChart", "bubbleChart"]);

type ClassicAxisTemplates = {
  categoryAxis: Element | null;
  seriesAxis: Element | null;
  valueAxes: Element[];
};

type ClassicChartTemplates = {
  axes: ClassicAxisTemplates;
  series: Element[];
  typeNodes: Element[];
};

const EMPTY_TEMPLATES: ClassicChartTemplates = {
  axes: { categoryAxis: null, seriesAxis: null, valueAxes: [] },
  series: [],
  typeNodes: [],
};

// ── Chart part lookup ─────────────────────────────────────────────────────

function findChart(chartAssets: WorkbookChartAssets, chartId: string) {
  for (const sheetCharts of chartAssets.chartsByWorkbookSheetIndex) {
    const match = sheetCharts.find((chart) => chart.id === chartId);
    if (match) {
      return match;
    }
  }
  for (const chartsheet of chartAssets.chartsheets) {
    const match = chartsheet.charts?.find((chart) => chart.id === chartId);
    if (match) {
      return match;
    }
  }
  return null;
}

/**
 * Series to write when the chart is rewritten.
 *
 * Point-series layouts (treemap, funnel, waterfall, …) collapse to a single
 * plotted series in the normalized model; the chart part still declares every
 * source series, which keeps a switch back to a classic type lossless.
 */
function resolveSeriesForRewrite(chart: XlsxChart) {
  const raw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
  const declared = raw?.chartExSourceSeries;
  return Array.isArray(declared) && declared.length > 0
    ? declared as XlsxChartSeries[]
    : chart.series;
}

/** Reads chart-level data labels from the part so an un-refreshed model still keeps them. */
function readSourceDataLabels(chartXml: string) {
  const document = parseXml(chartXml);
  const chartNode = document?.documentElement
    ? getFirstLocalDescendant(document.documentElement, "chart")
    : null;
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, "plotArea") : null;
  const typeNode = findPrimaryChartTypeNode(plotAreaNode);
  return parseChartDataLabelsFromXml(typeNode ? getFirstLocalChild(typeNode, "dLbls") : null);
}

function collectClassicTemplates(plotAreaNode: Element): ClassicChartTemplates {
  const typeNodes = Array.from(plotAreaNode.childNodes).filter(
    (node): node is Element => (
      node.nodeType === Node.ELEMENT_NODE
      && (/Chart$/.test((node as Element).localName))
    )
  );
  return {
    axes: {
      categoryAxis: getLocalChildren(plotAreaNode, "catAx")[0]
        ?? getLocalChildren(plotAreaNode, "dateAx")[0]
        ?? null,
      seriesAxis: getLocalChildren(plotAreaNode, "serAx")[0] ?? null,
      valueAxes: getLocalChildren(plotAreaNode, "valAx")
    },
    series: typeNodes.flatMap((node) => getLocalChildren(node, "ser")),
    typeNodes
  };
}

// ── DOM helpers ───────────────────────────────────────────────────────────

function chartElement(document: XMLDocument, localName: string) {
  return document.createElementNS(CHART_NS, `c:${localName}`);
}

function appendVal(parent: Element, localName: string, value: number | string) {
  const node = chartElement(parent.ownerDocument, localName);
  node.setAttribute("val", String(value));
  parent.appendChild(node);
  return node;
}

function numericChildValue(parent: Element | null, localName: string) {
  const value = Number(parent ? getFirstLocalChild(parent, localName)?.getAttribute("val") ?? Number.NaN : Number.NaN);
  return Number.isFinite(value) ? value : undefined;
}

const CHART_CHILD_ORDER = [
  "title",
  "autoTitleDeleted",
  "pivotFmts",
  "view3D",
  "floor",
  "sideWall",
  "backWall",
  "plotArea",
  "legend",
  "plotVisOnly",
  "dispBlanksAs",
  "showDLblsOverMax",
  "extLst"
];

function insertChartChild(chartNode: Element, node: Element) {
  const order = CHART_CHILD_ORDER.indexOf(node.localName);
  for (const child of Array.from(chartNode.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }
    const childOrder = CHART_CHILD_ORDER.indexOf((child as Element).localName);
    if (childOrder > order) {
      chartNode.insertBefore(node, child);
      return;
    }
  }
  chartNode.appendChild(node);
}

// ── Classic shape properties ──────────────────────────────────────────────

function solidFill(parent: Element, color: string) {
  const fill = parent.ownerDocument.createElementNS(DRAWINGML_NS, "a:solidFill");
  const srgb = parent.ownerDocument.createElementNS(DRAWINGML_NS, "a:srgbClr");
  srgb.setAttribute("val", color.replace(/^#/, "").toUpperCase());
  fill.appendChild(srgb);
  parent.appendChild(fill);
  return fill;
}

function buildSeriesShapeProperties(document: XMLDocument, series: XlsxChartSeries, node: string) {
  const fillColor = normalizeHexColor(series.color ?? "");
  const lineColor = normalizeHexColor(series.lineColor ?? "") ?? fillColor;
  if (!fillColor && !lineColor) {
    return null;
  }

  const shapeProperties = chartElement(document, "spPr");
  const filled = !(node === "lineChart" || node === "scatterChart" || node === "radarChart" || node === "stockChart");
  if (filled && fillColor) {
    solidFill(shapeProperties, fillColor);
  }
  if (lineColor) {
    const line = document.createElementNS(DRAWINGML_NS, "a:ln");
    if (typeof series.lineWidthPx === "number" && series.lineWidthPx > 0) {
      line.setAttribute("w", String(Math.max(1, Math.round(series.lineWidthPx * 12700))));
    }
    solidFill(line, lineColor);
    shapeProperties.appendChild(line);
  }
  return shapeProperties;
}

// ── Classic data references ───────────────────────────────────────────────

type ChartReferenceLike = { formula?: string } | null | undefined;

function appendCachePoints(document: XMLDocument, cache: Element, values: Array<number | string | null>, numeric: boolean) {
  cache.appendChild(chartElement(document, "ptCount")).setAttribute("val", String(values.length));
  values.forEach((value, index) => {
    if (value == null || value === "") {
      return;
    }
    const point = chartElement(document, "pt");
    point.setAttribute("idx", String(index));
    const pointValue = chartElement(document, "v");
    pointValue.textContent = numeric
      ? String(typeof value === "number" ? value : Number(value) || 0)
      : String(value);
    point.appendChild(pointValue);
    cache.appendChild(point);
  });
}

function buildDataNode(
  document: XMLDocument,
  localName: string,
  reference: ChartReferenceLike,
  values: Array<number | string | null>,
  numeric: boolean
) {
  const formula = reference?.formula?.replace(/^=/, "") ?? "";
  if (!formula && values.length === 0) {
    return null;
  }

  const parent = chartElement(document, localName);
  const cache = chartElement(document, numeric ? "numCache" : "strCache");
  if (numeric) {
    const formatCode = chartElement(document, "formatCode");
    formatCode.textContent = "General";
    cache.appendChild(formatCode);
  }
  appendCachePoints(document, cache, values, numeric);

  if (formula) {
    const referenceNode = chartElement(document, numeric ? "numRef" : "strRef");
    setLeafValue(referenceNode, "f", formula);
    referenceNode.appendChild(cache);
    parent.appendChild(referenceNode);
  } else {
    const literalNode = chartElement(document, numeric ? "numLit" : "strLit");
    literalNode.appendChild(cache);
    parent.appendChild(literalNode);
  }
  return parent;
}

// ── Classic series ────────────────────────────────────────────────────────

function setSeriesTextNode(seriesNode: Element, series: XlsxChartSeries) {
  const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
  const nameFormula = typeof raw?.name === "string" && raw.name.length > 0 ? raw.name : undefined;
  if (!nameFormula && !series.name) {
    return;
  }

  const tx = chartElement(seriesNode.ownerDocument, "tx");
  if (nameFormula) {
    const stringReference = chartElement(seriesNode.ownerDocument, "strRef");
    setLeafValue(stringReference, "f", nameFormula.replace(/^=/, ""));
    tx.appendChild(stringReference);
  } else {
    const value = chartElement(seriesNode.ownerDocument, "v");
    value.textContent = series.name ?? "";
    tx.appendChild(value);
  }
  seriesNode.appendChild(tx);
}

function buildMarkerNode(document: XMLDocument, series: XlsxChartSeries, option: ChartTypeOption, template: Element | null) {
  const templateMarker = template ? getFirstLocalChild(template, "marker") : null;
  const marker = templateMarker
    ? document.importNode(templateMarker, true) as Element
    : chartElement(document, "marker");
  const existingSymbol = getFirstLocalChild(marker, "symbol");
  // The chosen gallery entry decides whether markers are drawn; the template
  // only contributes marker styling (colour, size).
  const markerSymbol = option.marker === undefined
    ? series.markerSymbol ?? existingSymbol?.getAttribute("val") ?? "none"
    : option.marker
      ? (series.markerSymbol && series.markerSymbol !== "none"
        ? series.markerSymbol
        : existingSymbol && existingSymbol.getAttribute("val") !== "none"
          ? existingSymbol.getAttribute("val") as string
          : "circle")
      : "none";
  const symbol = existingSymbol ?? chartElement(document, "symbol");
  symbol.setAttribute("val", markerSymbol);
  if (!existingSymbol) {
    marker.appendChild(symbol);
  }
  if (series.markerSize != null) {
    const size = getFirstLocalChild(marker, "size") ?? chartElement(document, "size");
    size.setAttribute("val", String(Math.round(series.markerSize)));
    if (!size.parentNode) {
      const after = getFirstLocalChild(marker, "spPr") ?? getFirstLocalChild(marker, "extLst");
      if (after) {
        marker.insertBefore(size, after);
      } else {
        marker.appendChild(size);
      }
    }
  }
  return marker;
}

function buildSeriesNode(
  document: XMLDocument,
  series: XlsxChartSeries,
  index: number,
  option: ChartTypeOption,
  node: string,
  template: Element | null
) {
  const seriesNode = chartElement(document, "ser");
  appendVal(seriesNode, "idx", index);
  appendVal(seriesNode, "order", index);

  const templateTx = template ? getFirstLocalChild(template, "tx") : null;
  if (templateTx) {
    seriesNode.appendChild(document.importNode(templateTx, true));
  } else {
    setSeriesTextNode(seriesNode, series);
  }

  const templateShapeProperties = template ? getFirstLocalChild(template, "spPr") : null;
  if (templateShapeProperties) {
    seriesNode.appendChild(document.importNode(templateShapeProperties, true));
  } else {
    const shapeProperties = buildSeriesShapeProperties(document, series, node);
    if (shapeProperties) {
      seriesNode.appendChild(shapeProperties);
    }
  }

  const barLike = node === "barChart" || node === "bar3DChart";
  if (barLike && series.invertIfNegative) {
    appendVal(seriesNode, "invertIfNegative", 1);
  }

  if (MARKER_NODES.has(node)) {
    seriesNode.appendChild(buildMarkerNode(document, series, option, template));
  }

  if (PIE_LIKE_NODES.has(node) && template) {
    const explosion = getFirstLocalChild(template, "explosion");
    if (explosion) {
      seriesNode.appendChild(document.importNode(explosion, true));
    }
  }

  if (!SURFACE_NODES.has(node) && template) {
    for (const point of getLocalChildren(template, "dPt")) {
      seriesNode.appendChild(document.importNode(point, true));
    }
    const labels = getFirstLocalChild(template, "dLbls");
    if (labels) {
      seriesNode.appendChild(document.importNode(labels, true));
    }
  }

  const xy = XY_NODES.has(node);
  const categoryValues = xy && series.categories.some((value) => value != null && typeof value !== "number")
    ? series.values.map((_, valueIndex) => valueIndex + 1)
    : series.categories;
  const categoryReference = categoryValues === series.categories ? series.categoriesRef : null;
  const categoryNode = buildDataNode(
    document,
    xy ? "xVal" : "cat",
    categoryReference,
    categoryValues,
    xy || categoryValues.every((value) => value == null || typeof value === "number")
  );
  if (categoryNode) {
    seriesNode.appendChild(categoryNode);
  }

  const valueNode = buildDataNode(document, xy ? "yVal" : "val", series.valuesRef, series.values, true);
  if (valueNode) {
    seriesNode.appendChild(valueNode);
  }

  if (node === "bubbleChart") {
    const bubbleSizes = series.bubbleSizes && series.bubbleSizes.length > 0
      ? series.bubbleSizes
      : series.values.map(() => 1 as number);
    const bubbleNode = buildDataNode(document, "bubbleSize", series.bubbleSizeRef, bubbleSizes, true);
    if (bubbleNode) {
      seriesNode.appendChild(bubbleNode);
    }
    appendVal(seriesNode, "bubble3D", option.bubble3d ? 1 : 0);
  }

  if (node === "lineChart" || node === "line3DChart" || node === "scatterChart") {
    const smooth = series.smooth ?? Boolean(option.scatterStyle?.startsWith("smooth"));
    appendVal(seriesNode, "smooth", smooth ? 1 : 0);
  }

  return seriesNode;
}

// ── Classic chart-type nodes ──────────────────────────────────────────────

function buildDataLabelsNode(document: XMLDocument, labels: XlsxChartDataLabels) {
  const labelsNode = chartElement(document, "dLbls");
  if (labels.numberFormat?.formatCode) {
    const numFmt = chartElement(document, "numFmt");
    numFmt.setAttribute("formatCode", labels.numberFormat.formatCode);
    if (labels.numberFormat.sourceLinked !== undefined) {
      numFmt.setAttribute("sourceLinked", labels.numberFormat.sourceLinked ? "1" : "0");
    }
    labelsNode.appendChild(numFmt);
  }
  if (labels.position) {
    appendVal(labelsNode, "dLblPos", labels.position);
  }
  const flags: Array<[keyof XlsxChartDataLabels, string]> = [
    ["showLegendKey", "showLegendKey"],
    ["showValue", "showVal"],
    ["showCategoryName", "showCatName"],
    ["showSeriesName", "showSerName"],
    ["showPercent", "showPercent"],
    ["showBubbleSize", "showBubbleSize"]
  ];
  for (const [property, localName] of flags) {
    const value = labels[property];
    if (typeof value === "boolean") {
      appendVal(labelsNode, localName, value ? 1 : 0);
    }
  }
  return labelsNode;
}

type SeriesGroup = {
  axisIds: number[];
  node: string;
  option: ChartTypeOption;
  series: XlsxChartSeries[];
  startIndex: number;
};

function buildChartTypeNode(
  document: XMLDocument,
  chart: XlsxChart,
  group: SeriesGroup,
  templates: ClassicChartTemplates,
  withLabels: boolean
) {
  const { option, node, series, startIndex } = group;
  const typeNode = chartElement(document, node);

  if (node === "barChart" || node === "bar3DChart") {
    appendVal(typeNode, "barDir", option.barDir ?? "col");
    appendVal(typeNode, "grouping", option.grouping ?? "clustered");
    appendVal(typeNode, "varyColors", 0);
  } else if (node === "lineChart" || node === "line3DChart" || node === "areaChart" || node === "area3DChart") {
    appendVal(typeNode, "grouping", option.grouping ?? "standard");
    appendVal(typeNode, "varyColors", 0);
  } else if (node === "scatterChart") {
    appendVal(typeNode, "scatterStyle", option.scatterStyle ?? "lineMarker");
    appendVal(typeNode, "varyColors", 0);
  } else if (node === "bubbleChart") {
    appendVal(typeNode, "varyColors", 0);
  } else if (node === "radarChart") {
    appendVal(typeNode, "radarStyle", option.radarStyle ?? "standard");
    appendVal(typeNode, "varyColors", 0);
  } else if (node === "ofPieChart") {
    appendVal(typeNode, "ofPieType", option.ofPieType ?? "pie");
    appendVal(typeNode, "varyColors", 1);
  } else if (PIE_LIKE_NODES.has(node)) {
    appendVal(typeNode, "varyColors", 1);
  } else if (SURFACE_NODES.has(node)) {
    appendVal(typeNode, "wireframe", option.wireframe ? 1 : 0);
  }

  // Data labels live on the chart type node; when the target family differs the
  // labels of the chart's current primary node still describe the chart.
  const templateTypeNode = templates.typeNodes.find((entry) => entry.localName === node)
    ?? (withLabels ? templates.typeNodes[0] ?? null : null);
  series.forEach((entry, seriesIndex) => {
    typeNode.appendChild(buildSeriesNode(
      document,
      entry,
      startIndex + seriesIndex,
      option,
      node,
      templates.series[startIndex + seriesIndex] ?? null
    ));
  });

  if (!SURFACE_NODES.has(node) && withLabels) {
    const templateLabels = templateTypeNode ? getFirstLocalChild(templateTypeNode, "dLbls") : null;
    if (templateLabels) {
      typeNode.appendChild(document.importNode(templateLabels, true));
    } else if (chart.dataLabels) {
      typeNode.appendChild(buildDataLabelsNode(document, chart.dataLabels));
    }
  }

  if (node === "barChart") {
    appendVal(typeNode, "gapWidth", Math.round(chart.gapWidth ?? 150));
    if (chart.overlap !== undefined) {
      appendVal(typeNode, "overlap", Math.round(chart.overlap));
    } else if (option.grouping?.includes("tacked")) {
      // Matches both `stacked` and `percentStacked`.
      appendVal(typeNode, "overlap", 100);
    }
  } else if (node === "bar3DChart") {
    // `overlap` is not part of CT_Bar3DChart; 3-D bars always overlap in Excel.
    appendVal(typeNode, "gapWidth", Math.round(chart.gapWidth ?? 150));
  } else if (node === "pieChart") {
    appendVal(typeNode, "firstSliceAngle", Math.round(chart.firstSliceAngle ?? 0));
  } else if (node === "doughnutChart") {
    appendVal(typeNode, "firstSliceAngle", Math.round(chart.firstSliceAngle ?? 0));
    appendVal(typeNode, "holeSize", Math.round(chart.holeSize ?? 50));
  } else if (node === "ofPieChart") {
    appendVal(typeNode, "gapWidth", Math.round(chart.gapWidth ?? 150));
    appendVal(typeNode, "splitType", "auto");
    appendVal(typeNode, "secondPieSize", 75);
  } else if (node === "lineChart") {
    const templateMarker = templateTypeNode ? getFirstLocalChild(templateTypeNode, "marker") : null;
    if (templateMarker) {
      typeNode.appendChild(document.importNode(templateMarker, true));
    } else {
      appendVal(typeNode, "marker", 1);
    }
  } else if (node === "bubbleChart") {
    appendVal(typeNode, "bubble3D", option.bubble3d ? 1 : 0);
    appendVal(typeNode, "bubbleScale", Math.round(chart.bubbleScale ?? 100));
    appendVal(typeNode, "showNegBubbles", 0);
    appendVal(typeNode, "sizeRepresents", "area");
  } else if (node === "stockChart") {
    typeNode.appendChild(chartElement(document, "hiLowLines"));
    if (series.length === 4) {
      const upDownBars = chartElement(document, "upDownBars");
      appendVal(upDownBars, "gapWidth", 150);
      upDownBars.appendChild(chartElement(document, "upBars"));
      upDownBars.appendChild(chartElement(document, "downBars"));
      typeNode.appendChild(upDownBars);
    }
  }

  for (const axisId of group.axisIds) {
    appendVal(typeNode, "axId", axisId);
  }
  return typeNode;
}

// ── Classic axes ──────────────────────────────────────────────────────────

const CATEGORY_AXIS_CHILD_ORDER = [
  "axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt",
  "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt",
  "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "extLst"
];
const VALUE_AXIS_CHILD_ORDER = [
  "axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt",
  "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt",
  "crossBetween", "majorUnit", "minorUnit", "dispUnits", "extLst"
];

function ensureOrderedAxisChild(axisNode: Element, localName: string, order: string[]) {
  const existing = getFirstLocalChild(axisNode, localName);
  if (existing) {
    return existing;
  }
  const node = chartElement(axisNode.ownerDocument, localName);
  const index = order.indexOf(localName);
  for (const child of Array.from(axisNode.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }
    if (order.indexOf((child as Element).localName) > index) {
      axisNode.insertBefore(node, child);
      return node;
    }
  }
  axisNode.appendChild(node);
  return node;
}

function defaultAxisXml(kind: "catAx" | "valAx" | "serAx", position: string) {
  if (kind === "valAx") {
    return `<c:valAx xmlns:c="${CHART_NS}"><c:axId val="0"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${position}"/><c:majorGridlines/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="0"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  }
  return `<c:${kind} xmlns:c="${CHART_NS}"><c:axId val="0"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${position}"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="0"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:${kind}>`;
}

function buildAxisNode(options: {
  axisId: number;
  crossAxisId: number;
  crossBetween?: string;
  document: XMLDocument;
  hidden?: boolean;
  kind: "catAx" | "valAx" | "serAx";
  position: string;
  template: Element | null;
}) {
  const { axisId, crossAxisId, document, hidden, kind, position, template } = options;
  const source = template ?? parseXml(defaultAxisXml(kind, position))?.documentElement ?? null;
  const axisNode = source
    ? document.importNode(source, true) as Element
    : document.createElementNS(CHART_NS, `c:${kind}`);
  const order = kind === "valAx" ? VALUE_AXIS_CHILD_ORDER : CATEGORY_AXIS_CHILD_ORDER;

  ensureOrderedAxisChild(axisNode, "axId", order).setAttribute("val", String(axisId));
  ensureOrderedAxisChild(axisNode, "crossAx", order).setAttribute("val", String(crossAxisId));
  const positionNode = ensureOrderedAxisChild(axisNode, "axPos", order);
  positionNode.setAttribute("val", position);
  const deleteNode = ensureOrderedAxisChild(axisNode, "delete", order);
  deleteNode.setAttribute("val", hidden ? "1" : "0");
  if (kind === "valAx" && options.crossBetween) {
    ensureOrderedAxisChild(axisNode, "crossBetween", order).setAttribute("val", options.crossBetween);
  }
  return axisNode;
}

// ── Classic plot area ─────────────────────────────────────────────────────

function resolveVolumeStock(option: ChartTypeOption) {
  return option.id === "StockVHLC" || option.id === "StockVOHLC";
}

function buildSeriesGroups(chart: XlsxChart, option: ChartTypeOption): Array<{ node: string; option: ChartTypeOption; series: XlsxChartSeries[]; startIndex: number }> {
  const series = chart.series;
  if (option.combo) {
    const lineSeries = series.slice(-1);
    return [
      { node: option.node, option, series: series.slice(0, -1), startIndex: 0 },
      {
        node: "lineChart",
        option: { ...option, barDir: undefined, grouping: "standard", marker: true, node: "lineChart" },
        series: lineSeries,
        startIndex: series.length - 1
      }
    ];
  }
  if (resolveVolumeStock(option)) {
    return [
      {
        node: "barChart",
        option: { ...option, barDir: "col", grouping: "clustered", node: "barChart" },
        series: series.slice(0, 1),
        startIndex: 0
      },
      { node: option.node, option, series: series.slice(1), startIndex: 1 }
    ];
  }
  return [{ node: option.node, option, series, startIndex: 0 }];
}

function buildClassicPlotArea(
  document: XMLDocument,
  chart: XlsxChart,
  option: ChartTypeOption,
  templates: ClassicChartTemplates
) {
  const plotArea = chartElement(document, "plotArea");
  plotArea.appendChild(chartElement(document, "layout"));

  const noAxes = PIE_LIKE_NODES.has(option.node);
  const xy = option.node === "scatterChart" || option.node === "bubbleChart";
  const thirdAxis = option.node === "line3DChart" || option.node === "surface3DChart";
  const volume = resolveVolumeStock(option);
  const existingIds = [
    ...templates.axes.valueAxes.map((axis) => numericChildValue(axis, "axId")),
    numericChildValue(templates.axes.categoryAxis, "axId"),
    numericChildValue(templates.axes.seriesAxis, "axId")
  ].filter((value): value is number => typeof value === "number");
  const usedIds = new Set<number>();
  let nextAxisId = Math.max(100000, ...existingIds) + 1;
  const takeAxisId = (...candidates: Array<number | undefined>) => {
    for (const candidate of candidates) {
      if (typeof candidate === "number" && !usedIds.has(candidate)) {
        usedIds.add(candidate);
        return candidate;
      }
    }
    const id = nextAxisId;
    nextAxisId += 1;
    usedIds.add(id);
    return id;
  };

  const valueAxisIds = templates.axes.valueAxes.map((axis) => numericChildValue(axis, "axId"));
  const categoryAxisId = noAxes ? 0 : takeAxisId(
    numericChildValue(templates.axes.categoryAxis, "axId"),
    ...(xy ? [] : valueAxisIds)
  );
  const primaryAxisIds: number[] = [];
  let secondaryAxisIds: number[] = [];
  if (!noAxes) {
    // X/Y charts reuse the same two ids; the axis *kind* differs, not the ids.
    primaryAxisIds.push(categoryAxisId, takeAxisId(...valueAxisIds));
  }
  let thirdAxisId: number | undefined;
  if (thirdAxis) {
    thirdAxisId = takeAxisId(numericChildValue(templates.axes.seriesAxis, "axId"));
  }
  if (volume) {
    secondaryAxisIds = [takeAxisId(), takeAxisId()];
  }

  const groups = buildSeriesGroups(chart, option);
  const groupNodes: SeriesGroup[] = groups.map((group, groupIndex) => ({
    ...group,
    axisIds: noAxes
      ? []
      : volume
        ? (groupIndex === 0 ? primaryAxisIds : secondaryAxisIds)
        : thirdAxis
          ? [...primaryAxisIds, thirdAxisId as number]
          : primaryAxisIds
  }));
  groupNodes.forEach((group, groupIndex) => {
    plotArea.appendChild(buildChartTypeNode(
      document,
      chart,
      group,
      templates,
      groupIndex === 0
    ));
  });

  if (noAxes) {
    return plotArea;
  }

  const crossBetween = option.grouping === "percentStacked" ? "percent" : "between";
  const horizontalBar = option.barDir === "bar";
  const categoryPosition = horizontalBar ? "l" : "b";
  const valuePosition = horizontalBar ? "b" : "l";

  if (xy) {
    plotArea.appendChild(buildAxisNode({
      axisId: primaryAxisIds[0],
      crossAxisId: primaryAxisIds[1],
      document,
      kind: "valAx",
      position: "b",
      template: templates.axes.valueAxes[0] ?? null
    }));
    plotArea.appendChild(buildAxisNode({
      axisId: primaryAxisIds[1],
      crossAxisId: primaryAxisIds[0],
      crossBetween,
      document,
      kind: "valAx",
      position: "l",
      template: templates.axes.valueAxes[1] ?? null
    }));
    return plotArea;
  }

  plotArea.appendChild(buildAxisNode({
    axisId: primaryAxisIds[0],
    crossAxisId: primaryAxisIds[1],
    document,
    hidden: volume,
    kind: "catAx",
    position: categoryPosition,
    template: templates.axes.categoryAxis
  }));
  plotArea.appendChild(buildAxisNode({
    axisId: primaryAxisIds[1],
    crossAxisId: primaryAxisIds[0],
    crossBetween,
    document,
    kind: "valAx",
    position: valuePosition,
    template: templates.axes.valueAxes[0] ?? null
  }));
  if (thirdAxisId !== undefined) {
    plotArea.appendChild(buildAxisNode({
      axisId: thirdAxisId,
      crossAxisId: primaryAxisIds[1],
      document,
      kind: "serAx",
      position: "b",
      template: templates.axes.seriesAxis
    }));
  }
  if (volume) {
    plotArea.appendChild(buildAxisNode({
      axisId: secondaryAxisIds[0],
      crossAxisId: secondaryAxisIds[1],
      document,
      hidden: true,
      kind: "catAx",
      position: "b",
      template: templates.axes.categoryAxis
    }));
    plotArea.appendChild(buildAxisNode({
      axisId: secondaryAxisIds[1],
      crossAxisId: secondaryAxisIds[0],
      crossBetween,
      document,
      kind: "valAx",
      position: "r",
      template: templates.axes.valueAxes[0] ?? null
    }));
  }
  return plotArea;
}

function applyClassicPlot(
  document: XMLDocument,
  chartNode: Element,
  chart: XlsxChart,
  option: ChartTypeOption,
  templates: ClassicChartTemplates
) {
  const existingPlotArea = getFirstLocalChild(chartNode, "plotArea");
  const plotArea = buildClassicPlotArea(document, chart, option, templates);
  if (existingPlotArea?.parentNode) {
    existingPlotArea.parentNode.replaceChild(plotArea, existingPlotArea);
  } else {
    insertChartChild(chartNode, plotArea);
  }

  for (const localName of ["view3D", "floor", "sideWall", "backWall"]) {
    getLocalChildren(chartNode, localName).forEach((node) => node.parentNode?.removeChild(node));
  }
  if (option.is3d) {
    const view = chartElement(document, "view3D");
    appendVal(view, "rotX", 15);
    appendVal(view, "rotY", 20);
    appendVal(view, "rAngAx", 1);
    insertChartChild(chartNode, view);
  }
}

function createClassicChartDocument(chart: XlsxChart) {
  const document = parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="${RELATIONSHIPS_NS}"><c:roundedCorners val="0"/><c:chart/></c:chartSpace>`
  );
  const chartSpace = document?.documentElement;
  const chartNode = chartSpace ? getFirstLocalChild(chartSpace, "chart") : null;
  if (!document || !chartSpace || !chartNode) {
    return null;
  }

  if (chart.title) {
    setChartTitleNode(document, chartNode, chart.title);
  }
  const autoTitleDeleted = chart.autoTitleDeleted !== undefined
    ? chart.autoTitleDeleted
    : !chart.title;
  const autoTitleDeletedNode = chartElement(document, "autoTitleDeleted");
  autoTitleDeletedNode.setAttribute("val", autoTitleDeleted ? "1" : "0");
  insertChartChild(chartNode, autoTitleDeletedNode);
  return { chartNode, document };
}

function setChartTitleNode(document: XMLDocument, chartNode: Element, title: string) {
  const titleNode = chartElement(document, "title");
  const tx = chartElement(document, "tx");
  const rich = chartElement(document, "rich");
  const bodyPr = document.createElementNS(DRAWINGML_NS, "a:bodyPr");
  const paragraph = document.createElementNS(DRAWINGML_NS, "a:p");
  const run = document.createElementNS(DRAWINGML_NS, "a:r");
  const text = document.createElementNS(DRAWINGML_NS, "a:t");
  text.textContent = title;
  run.appendChild(text);
  paragraph.appendChild(run);
  rich.appendChild(bodyPr);
  rich.appendChild(paragraph);
  tx.appendChild(rich);
  titleNode.appendChild(tx);
  insertChartChild(chartNode, titleNode);
}

function buildClassicChartXml(
  chart: XlsxChart,
  option: ChartTypeOption,
  oldXml: string,
  fromModern: boolean
) {
  if (fromModern) {
    const created = createClassicChartDocument(chart);
    if (!created) {
      return null;
    }
    const { chartNode, document } = created;
    applyClassicPlot(document, chartNode, chart, option, EMPTY_TEMPLATES);
    insertChartChild(chartNode, buildLegendNode(document, chart));
    const plotVisibleOnlyNode = chartElement(document, "plotVisOnly");
    plotVisibleOnlyNode.setAttribute("val", chart.plotVisibleOnly === false ? "0" : "1");
    insertChartChild(chartNode, plotVisibleOnlyNode);
    const displayBlanksNode = chartElement(document, "dispBlanksAs");
    displayBlanksNode.setAttribute("val", chart.displayBlanksAs ?? "gap");
    insertChartChild(chartNode, displayBlanksNode);
    return serializeXml(document);
  }

  const document = parseXml(oldXml);
  const chartNode = document ? getFirstLocalDescendant(document.documentElement, "chart") : null;
  const plotArea = chartNode ? getFirstLocalChild(chartNode, "plotArea") : null;
  if (!document || !chartNode || !plotArea) {
    return null;
  }
  applyClassicPlot(document, chartNode, chart, option, collectClassicTemplates(plotArea));
  return serializeXml(document);
}

function buildLegendNode(document: XMLDocument, chart: XlsxChart) {
  const legend = chartElement(document, "legend");
  const position = chart.legend?.position ?? "r";
  appendVal(legend, "legendPos", position);
  appendVal(legend, "overlay", chart.legend?.overlay ? 1 : 0);
  return legend;
}

// ── chartEx generation ────────────────────────────────────────────────────

function stableGuid(seed: string) {
  let hashA = 0x811c9dc5;
  let hashB = 0x01000193;
  for (let index = 0; index < seed.length; index += 1) {
    const code = seed.charCodeAt(index);
    hashA = (hashA ^ code) * 16777619 >>> 0;
    hashB = (hashB + code) * 2654435761 >>> 0;
  }
  const part = (value: number) => value.toString(16).toUpperCase().padStart(8, "0");
  const bytes = `${part(hashA)}${part(hashB)}${part(hashA ^ hashB)}${part((hashA + hashB) >>> 0)}`;
  return `{${bytes.slice(0, 8)}-${bytes.slice(8, 12)}-${bytes.slice(12, 16)}-${bytes.slice(16, 20)}-${bytes.slice(20, 32)}}`;
}

function chartExDimension(
  kind: "numDim" | "strDim",
  type: string,
  reference: ChartReferenceLike,
  values: Array<number | string | null>
) {
  const formula = reference?.formula?.replace(/^=/, "") ?? "";
  const points = values
    .map((value, index) => (
      value == null || value === ""
        ? ""
        : `<cx:pt idx="${index}">${escapeXmlText(String(value))}</cx:pt>`
    ))
    .join("");
  const levelAttributes = `ptCount="${values.length}"${kind === "numDim" ? ' formatCode="General"' : ""}`;
  return `<cx:${kind} type="${type}">${formula ? `<cx:f>${escapeXmlText(formula)}</cx:f>` : ""}<cx:lvl ${levelAttributes}>${points}</cx:lvl></cx:${kind}>`;
}

function chartExSeriesShapeProperties(series: XlsxChartSeries) {
  const fillColor = normalizeHexColor(series.color ?? "");
  const lineColor = normalizeHexColor(series.lineColor ?? "");
  if (!fillColor && !lineColor) {
    return "";
  }
  const fill = fillColor
    ? `<a:solidFill><a:srgbClr val="${fillColor.replace(/^#/, "").toUpperCase()}"/></a:solidFill>`
    : "";
  const line = lineColor
    ? `<a:ln>${`<a:solidFill><a:srgbClr val="${lineColor.replace(/^#/, "").toUpperCase()}"/></a:solidFill>`}</a:ln>`
    : "";
  return `<cx:spPr>${fill}${line}</cx:spPr>`;
}

function formatChartExDataLabels(labels: XlsxChartDataLabels | null | undefined) {
  if (!labels) {
    return "";
  }
  const parts: string[] = [];
  if (labels.position) {
    parts.push(`<cx:dLblPos val="${escapeXmlText(labels.position)}"/>`);
  }
  const flags: Array<[keyof XlsxChartDataLabels, string]> = [
    ["showLegendKey", "showLegendKey"],
    ["showValue", "showVal"],
    ["showCategoryName", "showCatName"],
    ["showSeriesName", "showSerName"],
    ["showPercent", "showPercent"],
    ["showBubbleSize", "showBubbleSize"]
  ];
  for (const [property, localName] of flags) {
    const value = labels[property];
    if (typeof value === "boolean") {
      parts.push(`<cx:${localName} val="${value ? 1 : 0}"/>`);
    }
  }
  return parts.length > 0 ? `<cx:dataLabels>${parts.join("")}</cx:dataLabels>` : "";
}

function formatChartExTitle(title: string | undefined) {
  if (!title) {
    return "";
  }
  return `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXmlText(title)}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>`;
}

function formatChartExLegend(chart: XlsxChart) {
  const position = chart.legend?.position;
  const pos = position === "b" || position === "l" || position === "t" || position === "r" ? position : "r";
  return `<cx:legend pos="${pos}" align="ctr" overlay="${chart.legend?.overlay ? 1 : 0}"/>`;
}

function chartExLayout(option: ChartTypeOption) {
  return option.layout ?? "clusteredColumn";
}

function isChartExHistogram(option: ChartTypeOption) {
  return option.id === "Histogram" || option.id === "Pareto";
}

function buildChartExXml(chart: XlsxChart, option: ChartTypeOption) {
  const layout = chartExLayout(option);
  const histogram = isChartExHistogram(option);
  const sizeDimension = layout === "treemap" || layout === "sunburst";
  const noAxes = layout === "treemap" || layout === "sunburst" || layout === "funnel" || layout === "regionMap";
  const seriesLayout = histogram ? "clusteredColumn" : layout;

  const data = chart.series.map((series, index) => {
    const categories = series.categories.map((value) => (value == null ? null : String(value)));
    return `<cx:data id="${index}">${chartExDimension("strDim", "cat", series.categoriesRef, categories)}${chartExDimension("numDim", sizeDimension ? "size" : "val", series.valuesRef, series.values)}</cx:data>`;
  }).join("");

  const series = chart.series.map((entry, index) => {
    const axisIds = noAxes ? "" : '<cx:axisId val="0"/><cx:axisId val="1"/>';
    const binning = histogram
      ? '<cx:layoutPr><cx:binning intervalClosed="r"><cx:binCount val="10"/></cx:binning></cx:layoutPr>'
      : "";
    const labels = entry.name
      ? `<cx:tx><cx:txData><cx:v>${escapeXmlText(entry.name)}</cx:v></cx:txData></cx:tx>`
      : "";
    const dataLabels = formatChartExDataLabels(chart.dataLabels);
    return `<cx:series uniqueId="${stableGuid(`${chart.id}-${index}`)}" formatIdx="${index}" layoutId="${seriesLayout}">${labels}${chartExSeriesShapeProperties(entry)}${dataLabels}<cx:dataId val="${index}"/>${binning}${axisIds}</cx:series>`;
  }).join("");

  const paretoSeries = option.id === "Pareto"
    ? `<cx:series layoutId="paretoLine" ownerIdx="0"><cx:dataId val="0"/><cx:axisId val="0"/><cx:axisId val="2"/></cx:series>`
    : "";

  const axis = (id: number, category = false) => (
    `<cx:axis id="${id}"><cx:${category ? 'catScaling gapWidth="1"' : "valScaling"}/><cx:majorTickMarks type="none"/><cx:minorTickMarks type="none"/><cx:tickLabels/></cx:axis>`
  );
  const axes = noAxes
    ? ""
    : `${axis(0, true)}${axis(1)}${option.id === "Pareto" ? axis(2) : ""}`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:cx="${CHART_EX_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="${RELATIONSHIPS_NS}"><cx:chartData>${data}</cx:chartData><cx:chart>${formatChartExTitle(chart.title)}<cx:plotArea><cx:plotAreaRegion><cx:plotSurface/>${series}${paretoSeries}</cx:plotAreaRegion>${axes}</cx:plotArea>${formatChartExLegend(chart)}<cx:plotVisOnly val="${chart.plotVisibleOnly === false ? 0 : 1}"/></cx:chart></cx:chartSpace>`;
}

// ── Modern chart style parts ──────────────────────────────────────────────

const MODERN_STYLE_ELEMENT_NAMES = [
  "axisTitle",
  "categoryAxis",
  "chartArea",
  "dataLabel",
  "dataLabelCallout",
  "dataPoint",
  "dataPoint3D",
  "dataPointLine",
  "dataPointMarker",
  "dataPointMarkerLayout",
  "dataPointWireframe",
  "dataTable",
  "downBar",
  "dropLine",
  "errorBar",
  "floor",
  "gridlineMajor",
  "gridlineMinor",
  "hiLoLine",
  "leaderLine",
  "legend",
  "plotArea",
  "plotArea3D",
  "seriesAxis",
  "seriesLine",
  "title",
  "trendline",
  "trendlineLabel",
  "upBar",
  "valueAxis",
  "wall"
];

function nextStylePartPath(archive: Record<string, Uint8Array>, chartPath: string, kind: string) {
  const base = chartPath.replace(/\.xml$/i, "");
  let index = 1;
  while (archive[`${base}-${kind}-${index}.xml`]) {
    index += 1;
  }
  return `${base}-${kind}-${index}.xml`;
}

function chartStyleXml() {
  const entries = MODERN_STYLE_ELEMENT_NAMES.map((name) => (
    name === "dataPointMarkerLayout"
      ? '<cs:dataPointMarkerLayout symbol="circle" size="5"/>'
      : `<cs:${name}><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${name.startsWith("dataPoint") ? '<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr>' : ""}</cs:${name}>`
  )).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:chartStyle xmlns:cs="${CHART_STYLE_NS}" xmlns:a="${DRAWINGML_NS}" id="201">${entries}</cs:chartStyle>`;
}

function chartColorStyleXml() {
  const accents = [1, 2, 3, 4, 5, 6]
    .map((index) => `<a:schemeClr val="accent${index}"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:colorStyle xmlns:cs="${CHART_STYLE_NS}" xmlns:a="${DRAWINGML_NS}" meth="cycle" id="10">${accents}<cs:variation/></cs:colorStyle>`;
}

/**
 * Adds the `chartStyle`/`chartColorStyle` parts a chartEx chart needs, their
 * relationships and content-type overrides. Existing style parts are kept, so
 * repeated type switches never duplicate the parts.
 */
export function ensureModernChartStyles(archive: Record<string, Uint8Array>, chartPath: string) {
  const relsPath = relsPathForDocument(chartPath);
  const relsXml = readArchiveText(archive, relsPath)
    ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"></Relationships>`;
  const relationships = parseXml(relsXml);
  const root = relationships?.documentElement;
  if (!relationships || !root) {
    return false;
  }

  const existingTypes = new Set(
    getLocalDescendants(relationships, "Relationship")
      .map((node) => node.getAttribute("Type"))
      .filter((value): value is string => Boolean(value))
  );
  const parts: Array<[string, string, string]> = [
    [CHART_STYLE_REL_TYPE, "chartStyle", chartStyleXml()],
    [CHART_COLOR_STYLE_REL_TYPE, "chartColorStyle", chartColorStyleXml()]
  ];
  let changed = false;
  for (const [relationshipType, kind, xml] of parts) {
    if (existingTypes.has(relationshipType)) {
      continue;
    }
    const partPath = nextStylePartPath(archive, chartPath, kind);
    const existingIds = new Set(
      getLocalDescendants(relationships, "Relationship")
        .map((node) => node.getAttribute("Id"))
        .filter((value): value is string => Boolean(value))
    );
    let relationshipIndex = 1;
    while (existingIds.has(`rId${kind}${relationshipIndex}`)) {
      relationshipIndex += 1;
    }
    const relationship = relationships.createElementNS(PKG_REL_NS, "Relationship");
    relationship.setAttribute("Id", `rId${kind}${relationshipIndex}`);
    relationship.setAttribute("Type", relationshipType);
    relationship.setAttribute("Target", `/${normalizeArchivePath(partPath)}`);
    root.appendChild(relationship);
    archive[partPath] = strToU8(xml);
    ensureContentTypeOverride(
      archive,
      partPath,
      `application/vnd.ms-office.${kind.toLowerCase()}+xml`
    );
    changed = true;
  }
  if (changed) {
    archive[relsPath] = strToU8(serializeXml(relationships));
  }
  return true;
}

// ── classic ↔ chartEx migration ───────────────────────────────────────────

function buildChartExFallbackShape(document: XMLDocument) {
  const shapeXml = `<xdr:sp xmlns:xdr="${SPREADSHEET_DRAWING_NS}" xmlns:a="${DRAWINGML_NS}"><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>This chart requires a newer version of Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp>`;
  const parsed = parseXml(shapeXml)?.documentElement;
  return parsed ? document.importNode(parsed, true) as Element : null;
}

function migrateChartPartReference(
  archive: Record<string, Uint8Array>,
  origin: WorkbookChartOrigin,
  nextModern: boolean
) {
  const drawingXml = readArchiveText(archive, origin.drawingPath);
  const relsPath = relsPathForDocument(origin.drawingPath);
  const relsXml = readArchiveText(archive, relsPath);
  if (!drawingXml || !relsXml) {
    return false;
  }
  const drawing = parseXml(drawingXml);
  const relationships = parseXml(relsXml);
  const drawingRoot = drawing?.documentElement;
  const relationshipsRoot = relationships?.documentElement;
  if (!drawing || !relationships || !drawingRoot || !relationshipsRoot) {
    return false;
  }

  const chartPath = normalizeArchivePath(origin.chartPath ?? "");
  const relationship = getLocalDescendants(relationships, "Relationship").find((node) => {
    const target = node.getAttribute("Target");
    return Boolean(target) && resolveRelationshipPath(relsPath, target as string) === chartPath;
  });
  const relationshipId = relationship?.getAttribute("Id");
  if (!relationship || !relationshipId) {
    return false;
  }
  relationship.setAttribute("Type", nextModern ? CHART_EX_REL_TYPE : CHART_REL_TYPE);

  const chartReference = getLocalDescendants(drawing, "chart").find((node) => (
    (node.getAttributeNS(RELATIONSHIPS_NS, "id") ?? node.getAttribute("r:id")) === relationshipId
  ));
  if (!chartReference?.parentNode) {
    return false;
  }
  const graphicData = chartReference.parentNode.nodeType === Node.ELEMENT_NODE
    ? chartReference.parentNode as Element
    : null;
  if (!graphicData || graphicData.localName !== "graphicData") {
    return false;
  }

  const replacement = drawing.createElementNS(
    nextModern ? CHART_EX_NS : CHART_NS,
    nextModern ? "cx:chart" : "c:chart"
  );
  replacement.setAttributeNS(RELATIONSHIPS_NS, "r:id", relationshipId);
  graphicData.setAttribute("uri", nextModern ? CHART_EX_NS : CHART_NS);
  graphicData.insertBefore(replacement, chartReference);
  graphicData.removeChild(chartReference);

  const frame = replacement.parentNode?.parentNode?.parentNode;
  if (frame && frame.nodeType === Node.ELEMENT_NODE && (frame as Element).localName === "graphicFrame") {
    const frameElement = frame as Element;
    const frameParent = frameElement.parentNode as Element | null;
    if (nextModern && frameParent?.localName !== "AlternateContent") {
      const alternate = drawing.createElementNS(MARKUP_COMPATIBILITY_NS, "mc:AlternateContent");
      const choice = drawing.createElementNS(MARKUP_COMPATIBILITY_NS, "mc:Choice");
      choice.setAttribute("xmlns:cx1", CHART_EX_CHOICE_NS);
      choice.setAttribute("Requires", "cx1");
      const parent = frameElement.parentNode;
      if (!parent) {
        return false;
      }
      parent.insertBefore(alternate, frameElement);
      alternate.appendChild(choice);
      choice.appendChild(frameElement);
      const fallback = drawing.createElementNS(MARKUP_COMPATIBILITY_NS, "mc:Fallback");
      const shape = buildChartExFallbackShape(drawing);
      if (shape) {
        fallback.appendChild(shape);
      }
      alternate.appendChild(fallback);
    } else if (!nextModern && frameParent?.localName === "Choice") {
      const choice = frameParent;
      const alternate = choice.parentNode;
      const parent = alternate?.parentNode;
      if (alternate && parent) {
        parent.insertBefore(frameElement, alternate);
        parent.removeChild(alternate);
      }
    }
  }

  archive[normalizeArchivePath(origin.drawingPath)] = strToU8(serializeXml(drawing));
  archive[normalizeArchivePath(relsPath)] = strToU8(serializeXml(relationships));
  ensureContentTypeOverride(
    archive,
    origin.chartPath as string,
    nextModern ? CHART_EX_CONTENT_TYPE : CLASSIC_CHART_CONTENT_TYPE
  );
  return true;
}

// ── Public API ────────────────────────────────────────────────────────────

/**
 * Rewrites an existing chart part as another chart type.
 *
 * Classic families rebuild the `c:plotArea` (series names, colours, data labels,
 * point styles and axis formatting are carried over). chartEx families rebuild
 * the part as a `cx:chartSpace`, add the chart style parts it needs, and
 * migrating in either direction keeps the drawing relationship, the
 * `mc:AlternateContent` wrapper and the chart content type consistent.
 *
 * The caller re-parses chart assets afterwards; this function only mutates the
 * archive and the origin's `chartKind`.
 *
 * @returns `true` when the chart part was rewritten.
 */
export function changeWorkbookChartType(
  imageAssets: Pick<WorkbookImageAssets, "archive">,
  chartAssets: WorkbookChartAssets,
  chartId: string,
  chartTypeId: string
): boolean {
  const option = resolveChartTypeOption(chartTypeId);
  if (!option) {
    return false;
  }
  const origin = chartAssets.chartOriginsById.get(chartId);
  if (!origin?.chartPath) {
    return false;
  }
  const chart = findChart(chartAssets, chartId);
  if (!chart || chart.editable === false) {
    return false;
  }

  const archive = imageAssets.archive;
  const oldXml = readArchiveText(archive, origin.chartPath);
  if (!oldXml) {
    return false;
  }

  const nextModern = Boolean(option.layout);
  const nextSeries = resolveSeriesForRewrite(chart);
  const dataLabels = chart.dataLabels ?? readSourceDataLabels(oldXml);
  const chartForRewrite: XlsxChart = {
    ...chart,
    dataLabels,
    series: nextSeries
  };
  const nextXml = nextModern
    ? buildChartExXml(chartForRewrite, option)
    : buildClassicChartXml(chartForRewrite, option, oldXml, origin.chartKind === "modern");
  if (!nextXml) {
    return false;
  }

  const chartPartPath = normalizeArchivePath(origin.chartPath);
  archive[chartPartPath] = strToU8(nextXml);
  if (nextModern) {
    ensureModernChartStyles(archive, origin.chartPath);
  }
  if ((origin.chartKind === "modern") !== nextModern) {
    if (!migrateChartPartReference(archive, origin, nextModern)) {
      // Without the relationship/wrapper/content-type migration the part and the
      // package would disagree; leave the chart exactly as it was.
      archive[chartPartPath] = strToU8(oldXml);
      return false;
    }
  }
  origin.chartKind = nextModern ? "modern" : "classic";
  return true;
}
