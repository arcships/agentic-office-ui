import {
  getFirstLocalChild,
  getLocalChildren,
  getLocalDescendants,
  parseXml,
} from "./chart-xml-utils";

/**
 * Parses a `cx:chartSpace` chart part into the raw chartEx shape the chart
 * normalizer (`normalizeChartExChart`) consumes.
 *
 * The workbook engine used by this package does not expose chartEx charts, so
 * chart parts are read from the package archive exactly like classic charts.
 * The shape mirrors what an engine would report: `chartData` dimensions, the
 * plot-area series (layout, binning, axis ids) and the chartEx axes.
 */
export function parseChartExChartXml(chartXml: string): Record<string, unknown> | null {
  const document = parseXml(chartXml);
  const chartSpace = document?.documentElement;
  if (!chartSpace || chartSpace.localName !== "chartSpace") {
    return null;
  }

  const chartNode = getFirstLocalChild(chartSpace, "chart");
  if (!chartNode) {
    return null;
  }

  const plotAreaNode = getFirstLocalChild(chartNode, "plotArea");
  const regionNode = plotAreaNode ? getFirstLocalChild(plotAreaNode, "plotAreaRegion") : null;
  const series = regionNode
    ? getLocalChildren(regionNode, "series").map(parseChartExSeries)
    : [];
  const axes = plotAreaNode
    ? getLocalChildren(plotAreaNode, "axis").map(parseChartExAxis)
    : [];
  const legendNode = getFirstLocalChild(chartNode, "legend");
  const titleNode = getFirstLocalChild(chartNode, "title");
  const plotVisibleOnlyNode = getFirstLocalChild(chartNode, "plotVisOnly");

  return {
    anchor: undefined,
    data: parseChartExData(chartSpace),
    layout: resolveChartExLayout(series),
    legend: legendNode
      ? {
          align: legendNode.getAttribute("align") ?? undefined,
          overlay: legendNode.getAttribute("overlay") === "1",
          pos: legendNode.getAttribute("pos") ?? undefined
        }
      : undefined,
    plotArea: {
      axes,
      series
    },
    plotVisOnly: plotVisibleOnlyNode
      ? plotVisibleOnlyNode.getAttribute("val") !== "0"
      : undefined,
    title: titleNode ? { text: readChartExText(titleNode) } : undefined,
    valueColors: undefined
  };
}

function parseChartExData(chartSpace: Element) {
  const chartDataNode = getFirstLocalChild(chartSpace, "chartData");
  if (!chartDataNode) {
    return [];
  }
  return getLocalChildren(chartDataNode, "data").map((dataNode) => {
    const id = Number(dataNode.getAttribute("id") ?? Number.NaN);
    const dimensions = [
      ...getLocalChildren(dataNode, "strDim"),
      ...getLocalChildren(dataNode, "numDim")
    ].map((dimensionNode) => ({
      dimType: dimensionNode.getAttribute("type")
        ?? (dimensionNode.localName === "numDim" ? "val" : "cat"),
      formula: getFirstLocalChild(dimensionNode, "f")?.textContent?.trim() || undefined,
      values: readChartExLevels(dimensionNode)
    }));
    return {
      dimensions,
      id: Number.isFinite(id) ? id : 0
    };
  });
}

function readChartExLevels(dimensionNode: Element) {
  const values: Array<number | string | null> = [];
  for (const levelNode of getLocalChildren(dimensionNode, "lvl")) {
    for (const pointNode of getLocalChildren(levelNode, "pt")) {
      const index = Number(pointNode.getAttribute("idx") ?? Number.NaN);
      if (!Number.isFinite(index) || index < 0) {
        continue;
      }
      const text = pointNode.textContent ?? "";
      const numeric = dimensionNode.localName === "numDim" ? Number(text) : Number.NaN;
      values[index] = Number.isFinite(numeric) && text.trim() !== "" ? numeric : text;
    }
  }
  return values;
}

function readChartExText(node: Element | null) {
  if (!node) {
    return undefined;
  }
  const richText = getLocalDescendants(node, "t")
    .map((textNode) => textNode.textContent ?? "")
    .join("")
    .trim();
  if (richText) {
    return richText;
  }
  const directValue = getFirstLocalDescendantValue(node, "v");
  return directValue || undefined;
}

function getFirstLocalDescendantValue(node: Element, localName: string) {
  const match = getLocalDescendants(node, localName)[0];
  return match?.textContent?.trim() || undefined;
}

function readChartExSeriesText(seriesNode: Element) {
  const titleNode = getFirstLocalChild(seriesNode, "tx");
  if (!titleNode) {
    return undefined;
  }
  const txData = getFirstLocalChild(titleNode, "txData");
  const txValue = txData ? getFirstLocalChild(txData, "v")?.textContent?.trim() : undefined;
  if (txValue) {
    return txValue;
  }
  const directValue = getFirstLocalChild(titleNode, "v")?.textContent?.trim();
  if (directValue) {
    return directValue;
  }
  const formula = getFirstLocalChild(titleNode, "f")?.textContent?.trim();
  return formula || undefined;
}

function readChartExBinning(seriesNode: Element) {
  const layoutProperties = getFirstLocalChild(seriesNode, "layoutPr");
  const binningNode = layoutProperties ? getFirstLocalChild(layoutProperties, "binning") : null;
  if (!binningNode) {
    return undefined;
  }
  const binning: Record<string, unknown> = {};
  for (const attribute of Array.from(binningNode.attributes)) {
    const name = attribute.localName || attribute.name;
    const rawValue = attribute.value;
    const numeric = Number(rawValue);
    binning[name] = rawValue.trim() !== "" && Number.isFinite(numeric) ? numeric : rawValue;
  }
  for (const child of Array.from(binningNode.childNodes)) {
    if (child.nodeType !== 1) {
      continue;
    }
    const element = child as Element;
    const numeric = Number(element.getAttribute("val") ?? Number.NaN);
    binning[element.localName] = Number.isFinite(numeric) ? numeric : element.textContent?.trim();
  }
  return binning;
}

function readChartExLayoutProperties(seriesNode: Element) {
  const layoutProperties = getFirstLocalChild(seriesNode, "layoutPr");
  if (!layoutProperties) {
    return undefined;
  }
  const binning = readChartExBinning(seriesNode);
  const subtotals = getLocalChildren(layoutProperties, "subtotals");
  const layoutPr: Record<string, unknown> = {};
  if (binning) {
    layoutPr.binning = binning;
  }
  if (subtotals.length > 0) {
    layoutPr.subtotals = subtotals.map((subtotalNode) => {
      const subtotal: Record<string, unknown> = {};
      for (const attribute of Array.from(subtotalNode.attributes)) {
        subtotal[attribute.localName || attribute.name] = attribute.value;
      }
      return subtotal;
    });
  }
  const aggregation = getFirstLocalChild(layoutProperties, "aggregation");
  if (aggregation) {
    layoutPr.aggregation = aggregation.getAttribute("val") !== "0";
  }
  return Object.keys(layoutPr).length > 0 ? layoutPr : undefined;
}

function readChartExValueColors(seriesNode: Element) {
  const valueColorsNode = getFirstLocalChild(seriesNode, "valueColors");
  if (!valueColorsNode) {
    return undefined;
  }
  const colors = getLocalChildren(valueColorsNode, "valueColor")
    .map((colorNode) => {
      const srgbNode = getLocalDescendants(colorNode, "srgbClr")[0];
      const value = srgbNode?.getAttribute("val");
      return value ? `#${value.slice(-6).toLowerCase()}` : undefined;
    })
    .filter((value): value is string => Boolean(value));
  return colors.length > 0 ? colors : undefined;
}

function parseChartExSeries(seriesNode: Element) {
  const axisIds = getLocalChildren(seriesNode, "axisId")
    .map((axisNode) => Number(axisNode.getAttribute("val") ?? Number.NaN))
    .filter((value) => Number.isFinite(value));
  const dataId = Number(getFirstLocalChild(seriesNode, "dataId")?.getAttribute("val") ?? Number.NaN);
  const formatIdx = Number(seriesNode.getAttribute("formatIdx") ?? Number.NaN);
  const dataLabelsNode = getFirstLocalChild(seriesNode, "dataLabels");
  return {
    axisId: axisIds,
    dataId: Number.isFinite(dataId) ? dataId : undefined,
    dataLabels: dataLabelsNode ? parseChartExDataLabels(dataLabelsNode) : undefined,
    formatIdx: Number.isFinite(formatIdx) ? formatIdx : undefined,
    layoutId: seriesNode.getAttribute("layoutId") ?? seriesNode.getAttribute("layout") ?? undefined,
    layoutPr: readChartExLayoutProperties(seriesNode),
    text: readChartExSeriesText(seriesNode),
    uniqueId: seriesNode.getAttribute("uniqueId") ?? undefined,
    valueColors: readChartExValueColors(seriesNode)
  };
}

function parseChartExDataLabels(labelsNode: Element) {
  const flag = (localName: string) => {
    const node = getFirstLocalChild(labelsNode, localName);
    if (!node) {
      return undefined;
    }
    const value = node.getAttribute("val");
    return value === null ? true : value !== "0" && value !== "false";
  };
  const positionNode = getFirstLocalChild(labelsNode, "dLblPos")
    ?? getFirstLocalChild(labelsNode, "position");
  return {
    numberFormat: undefined,
    position: positionNode?.getAttribute("val") ?? undefined,
    showBubbleSize: flag("showBubbleSize"),
    showCategoryName: flag("showCatName"),
    showLegendKey: flag("showLegendKey"),
    showPercent: flag("showPercent"),
    showSeriesName: flag("showSerName"),
    showValue: flag("showVal")
  };
}

function parseChartExAxis(axisNode: Element) {
  const id = Number(axisNode.getAttribute("id") ?? Number.NaN);
  const scalingNode = getFirstLocalChild(axisNode, "catScaling")
    ?? getFirstLocalChild(axisNode, "valScaling");
  const scaling: Record<string, unknown> = {};
  if (scalingNode) {
    for (const [attribute, key] of [
      ["max", "max"],
      ["min", "min"],
      ["majorUnit", "majorUnit"],
      ["minorUnit", "minorUnit"],
      ["logBase", "logBase"],
      ["gapWidth", "gapWidth"]
    ] as const) {
      const value = Number(scalingNode.getAttribute(attribute) ?? Number.NaN);
      if (Number.isFinite(value)) {
        scaling[key] = value;
      }
    }
    const orientation = scalingNode.getAttribute("orientation");
    if (orientation) {
      scaling.orientation = orientation;
    }
  }
  const numberFormatNode = getFirstLocalChild(axisNode, "numFmt");
  const deletedNode = getFirstLocalChild(axisNode, "deleted");
  const tickLabelSkipNode = getFirstLocalChild(axisNode, "tickLabelSkip");
  const positionNode = getFirstLocalChild(axisNode, "pos");
  return {
    crossId: undefined,
    hidden: deletedNode ? deletedNode.getAttribute("val") !== "0" : undefined,
    id: Number.isFinite(id) ? id : undefined,
    majorGridlines: getFirstLocalChild(axisNode, "majorGridlines") ? true : undefined,
    minorGridlines: getFirstLocalChild(axisNode, "minorGridlines") ? true : undefined,
    numberFormat: numberFormatNode
      ? {
          formatCode: numberFormatNode.getAttribute("formatCode") ?? undefined,
          sourceLinked: numberFormatNode.getAttribute("sourceLinked") === "1"
            ? true
            : numberFormatNode.getAttribute("sourceLinked") === "0"
              ? false
              : undefined
        }
      : undefined,
    position: positionNode?.getAttribute("val") ?? undefined,
    scaling,
    tickLabelSkip: tickLabelSkipNode
      ? Number(tickLabelSkipNode.getAttribute("val") ?? Number.NaN)
      : undefined
  };
}

function resolveChartExLayout(series: Array<Record<string, unknown>>) {
  const layoutIds = series
    .map((entry) => (typeof entry.layoutId === "string" ? entry.layoutId : undefined))
    .filter((value): value is string => Boolean(value));
  // Pareto charts plot histogram columns plus a cumulative line; reporting the
  // line as the chart layout keeps the gallery entry and chart type in sync.
  if (layoutIds.includes("paretoLine")) {
    return "paretoLine";
  }
  return layoutIds[0];
}
