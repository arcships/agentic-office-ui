import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8 } from "fflate";
import type { WorkbookImageAssets, WorkbookImageSheetOrigin } from "../images";
import type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartLegend,
  XlsxChartPointDataLabel,
  XlsxChartPointStyle,
  XlsxChartReference,
  XlsxChartSeries,
  XlsxChartTypeGroup,
  XlsxChartWall,
  XlsxChartsheet,
  XlsxImageAnchor,
  XlsxThemePalette,
  XlsxWorkbookTab,
} from "../types";

import {
  CHART_NS,
  normalizeChartReference,
  CHART_COLOR_STYLE_REL_TYPE,
  CHART_STYLE_REL_TYPE,
  normalizeChartSeries,
  cellValueToNumber,
} from "./chart-series";

import {
  isElementNode,
  normalizeHexColor,
  resolveChartColorNode,
  findFirstChartColorElement,
  resolveChartFillColor,
  resolveChartLineStyle,
  readChartTextTypeface,
  applyLightnessTransform,
} from "./chart-colors";

import {
  applyChartSeriesStyleFromXml,
  readChartRelationships,
  readChartColorPalette,
  readChartStyleAppearance,
  applyBuiltinChartDefaults,
  parseFallbackPointStylesFromChartXml,
  parseFallbackSeriesStylesFromChartXml,
  parseFallbackBubbleSizesFromChartXml,
  resolveArchiveFallbackBubbleSizes,
  resolveArchiveFallbackPointStyles,
} from "./chart-styles";

import {
  normalizeChartExChart,
  WorkbookChartOrigin,
  WorkbookChartAssets,
  findPrimaryChartTypeNode,
  resolveScatterChartType,
  buildChartExHistogramSeries,
  buildChartExParetoLineSeries,
  resolveChartExTextFormula,
  PRIMARY_CHART_TYPE_LOCAL_NAMES,
} from "./chart-types";

import {
  normalizeChartAxis,
  mergeChartAxis,
  readChartAxisFromXml,
  readChartWallFromXml,
  normalizeChartDataLabels,
  normalizeChartAnchor,
  normalizeChartTypeGroup,
  collectChartOriginsForSheet,
  applyChartOrigins,
  normalizeChartsheet,
  buildTabs,
} from "./chart-export";

export function decodeChartXmlText(value: string) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'");
}

export function normalizeChartTitleForMatch(value: string | null | undefined) {
  return (value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}

export function extractChartTitleFromXml(chartXml: string): string | null {
  const match = chartXml.match(/<c:title\b[\s\S]*?<a:t>([\s\S]*?)<\/a:t>/i);
  if (!match?.[1]) {
    return null;
  }
  const decoded = decodeChartXmlText(match[1]).trim();
  return decoded.length > 0 ? decoded : null;
}

export function parseChartTypeFromXml(chartXml: string) {
  for (const chartType of PRIMARY_CHART_TYPE_LOCAL_NAMES) {
    if (new RegExp(`<c:${chartType}\\b`, "i").test(chartXml)) {
      return chartType;
    }
  }
  return "";
}

export function normalizeLegend(raw: unknown): XlsxChartLegend | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }
  const legend = raw as Record<string, unknown>;
  return {
    overlay: typeof legend.overlay === "boolean" ? legend.overlay : undefined,
    position: typeof legend.position === "string" ? legend.position : undefined,
    raw: legend
  };
}

export function normalizeLegendPosition(position: string | undefined) {
  if (!position) {
    return undefined;
  }
  switch (position) {
    case "bottom":
      return "b";
    case "left":
      return "l";
    case "right":
      return "r";
    case "top":
      return "t";
    default:
      return position;
  }
}

export function readChartNumericAttribute(parent: Element | null, localName: string) {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  const value = Number(node?.getAttribute("val") ?? Number.NaN);
  return Number.isFinite(value) ? value : undefined;
}

export function readChartBooleanAttribute(parent: Element | null, localName: string) {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  if (!node) {
    return undefined;
  }
  const rawValue = node.getAttribute("val");
  if (rawValue == null) {
    return true;
  }
  if (rawValue === "1" || rawValue === "true") {
    return true;
  }
  if (rawValue === "0" || rawValue === "false") {
    return false;
  }
  return undefined;
}

export function readChartLabelFontSizePt(textPropertiesNode: Element | null) {
  if (!textPropertiesNode) {
    return undefined;
  }
  const runPropertiesNode = getFirstLocalDescendant(textPropertiesNode, "defRPr")
    ?? getFirstLocalDescendant(textPropertiesNode, "rPr");
  const rawSize = Number(runPropertiesNode?.getAttribute("sz") ?? Number.NaN);
  if (!Number.isFinite(rawSize) || rawSize <= 0) {
    return undefined;
  }
  return rawSize / 100;
}

export function parseChartPointDataLabelsFromXml(labelsNode: Element): XlsxChartPointDataLabel[] {
  const fallbackFontSizePt = readChartLabelFontSizePt(getFirstLocalChild(labelsNode, "txPr"));
  const labels: XlsxChartPointDataLabel[] = [];
  for (const pointLabelNode of getLocalChildren(labelsNode, "dLbl")) {
    const index = readChartNumericAttribute(pointLabelNode, "idx");
    if (typeof index !== "number" || !Number.isFinite(index)) {
      continue;
    }

    const layoutNode = getFirstLocalChild(pointLabelNode, "layout");
    const manualLayoutNode = layoutNode ? getFirstLocalChild(layoutNode, "manualLayout") : null;
    labels.push({
      deleted: readChartBooleanAttribute(pointLabelNode, "delete"),
      fontSizePt: readChartLabelFontSizePt(getFirstLocalChild(pointLabelNode, "txPr")) ?? fallbackFontSizePt,
      index,
      showBubbleSize: readChartBooleanAttribute(pointLabelNode, "showBubbleSize"),
      showCategoryName: readChartBooleanAttribute(pointLabelNode, "showCatName"),
      showPercent: readChartBooleanAttribute(pointLabelNode, "showPercent"),
      showSeriesName: readChartBooleanAttribute(pointLabelNode, "showSerName"),
      showValue: readChartBooleanAttribute(pointLabelNode, "showVal"),
      x: readChartNumericAttribute(manualLayoutNode, "x"),
      y: readChartNumericAttribute(manualLayoutNode, "y")
    });
  }
  return labels;
}

export function parseChartDataLabelsFromXml(labelsNode: Element | null): XlsxChartDataLabels | null {
  if (!labelsNode) {
    return null;
  }

  const pointLabels = parseChartPointDataLabelsFromXml(labelsNode);
  const labels: XlsxChartDataLabels = {
    pointLabels: pointLabels.length > 0 ? pointLabels : undefined,
    raw: {},
    showBubbleSize: readChartBooleanAttribute(labelsNode, "showBubbleSize"),
    showCategoryName: readChartBooleanAttribute(labelsNode, "showCatName"),
    showLegendKey: readChartBooleanAttribute(labelsNode, "showLegendKey"),
    showPercent: readChartBooleanAttribute(labelsNode, "showPercent"),
    showSeriesName: readChartBooleanAttribute(labelsNode, "showSerName"),
    showValue: readChartBooleanAttribute(labelsNode, "showVal")
  };
  const hasValue = (
    labels.showBubbleSize !== undefined
    || labels.showCategoryName !== undefined
    || labels.showLegendKey !== undefined
    || labels.showPercent !== undefined
    || (labels.pointLabels?.length ?? 0) > 0
    || labels.showSeriesName !== undefined
    || labels.showValue !== undefined
  );
  return hasValue ? labels : null;
}

function applyChartStyleFromXml(
  chart: XlsxChart,
  chartPath: string | undefined,
  archive: Record<string, Uint8Array>,
  themePalette?: XlsxThemePalette | null
) {
  const chartXml = readArchiveText(archive, chartPath);
  if (!chartXml) {
    return;
  }
  const relationships = chartPath ? readChartRelationships(archive, chartPath) : new Map<string, string>();
  const fallbackPointStylesBySeries = parseFallbackPointStylesFromChartXml(chartXml, themePalette);
  const fallbackSeriesStyles = parseFallbackSeriesStylesFromChartXml(chartXml, themePalette);
  const fallbackBubbleSizesBySeries = parseFallbackBubbleSizesFromChartXml(chartXml);
  const applyFallbackSeriesStyles = () => {
    if (fallbackBubbleSizesBySeries.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackBubbleSizes = fallbackBubbleSizesBySeries[seriesIndex] ?? [];
        if (fallbackBubbleSizes.length === 0) {
          return series;
        }

        const currentNumericPointCount = (series.bubbleSizes ?? []).filter(
          (value): value is number => typeof value === "number" && Number.isFinite(value)
        ).length;
        const fallbackNumericPointCount = fallbackBubbleSizes.filter(
          (value): value is number => typeof value === "number" && Number.isFinite(value)
        ).length;

        if (currentNumericPointCount >= fallbackNumericPointCount) {
          return series;
        }

        return {
          ...series,
          bubbleSizes: fallbackBubbleSizes
        };
      });
    }

    if (fallbackPointStylesBySeries.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackPointStyles = fallbackPointStylesBySeries[seriesIndex] ?? [];
        if (fallbackPointStyles.length === 0) {
          return series;
        }

        const existingByIndex = new Map((series.dataPointStyles ?? []).map((entry) => [entry.index, entry]));
        for (const fallbackStyle of fallbackPointStyles) {
          const existing = existingByIndex.get(fallbackStyle.index);
          existingByIndex.set(fallbackStyle.index, {
            color: existing?.color ?? fallbackStyle.color,
            explosion: existing?.explosion ?? fallbackStyle.explosion,
            index: fallbackStyle.index,
            lineColor: existing?.lineColor ?? fallbackStyle.lineColor
          });
        }

        return {
          ...series,
          dataPointStyles: Array.from(existingByIndex.values()).sort((left, right) => left.index - right.index)
        };
      });
    }

    if (fallbackSeriesStyles.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const fallbackStyle = fallbackSeriesStyles[seriesIndex];
        if (!fallbackStyle) {
          return series;
        }
        const fallbackColor = fallbackStyle.color ?? fallbackStyle.lineColor;
        return {
          ...series,
          color: series.color ?? fallbackColor,
          lineColor: series.lineColor ?? fallbackStyle.lineColor ?? fallbackColor,
          markerColor: series.markerColor ?? fallbackColor ?? series.color,
          markerLineColor: series.markerLineColor ?? fallbackStyle.lineColor ?? fallbackColor ?? series.lineColor
        };
      });
    }
  };
  const applyRelationshipStyles = () => {
    chart.chartColorPalette = readChartColorPalette(archive, relationships.get(CHART_COLOR_STYLE_REL_TYPE), themePalette);
    const styleAppearance = readChartStyleAppearance(
      archive,
      relationships.get(CHART_STYLE_REL_TYPE),
      themePalette
    );
    chart.axisLabelColor = styleAppearance.axisLabelColor ?? chart.axisLabelColor;
    chart.axisLineColor = styleAppearance.axisLineColor ?? chart.axisLineColor;
    chart.chartAreaBorderColor = styleAppearance.chartAreaBorderColor ?? chart.chartAreaBorderColor;
    chart.chartAreaFillColor = styleAppearance.chartAreaFillColor ?? chart.chartAreaFillColor;
    chart.chartColorPaletteOffset = styleAppearance.paletteOffset ?? chart.chartColorPaletteOffset;
    chart.textColor = styleAppearance.textColor ?? chart.textColor;
    chart.titleColor = styleAppearance.titleColor ?? chart.titleColor;
    return styleAppearance;
  };
  const applyModernChartExStyles = () => {
    const modernPlotAreaNode = chartDocument?.documentElement
      ? getFirstLocalDescendant(chartDocument.documentElement, "plotArea")
      : null;
    if (!modernPlotAreaNode) {
      return;
    }

    const parseModernBinning = (seriesNode: Element) => {
      const layoutPrNode = getFirstLocalChild(seriesNode, "layoutPr");
      const binningNode = layoutPrNode ? getFirstLocalChild(layoutPrNode, "binning") : null;
      if (!binningNode) {
        return null;
      }
      const binning: Record<string, unknown> = {};
      for (const attribute of Array.from(binningNode.attributes)) {
        const rawValue = attribute.value;
        const numeric = Number(rawValue);
        binning[attribute.localName || attribute.name] = Number.isFinite(numeric) && rawValue.trim() !== ""
          ? numeric
          : rawValue;
      }
      return Object.keys(binning).length > 0 ? binning : {};
    };

    const plotAreaShapeProperties = getFirstLocalChild(modernPlotAreaNode, "spPr");
    if (plotAreaShapeProperties) {
      const plotAreaFillColor = resolveChartFillColor(plotAreaShapeProperties, themePalette);
      const plotAreaLineStyle = resolveChartLineStyle(plotAreaShapeProperties, themePalette);
      if (plotAreaFillColor) {
        chart.chartAreaFillColor = chart.chartAreaFillColor ?? plotAreaFillColor;
      }
      if (plotAreaLineStyle.color) {
        chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? plotAreaLineStyle.color;
      }
    }

    const modernSeriesNodes = getLocalDescendants(modernPlotAreaNode, "series");
    if (modernSeriesNodes.length === 0) {
      return;
    }

    chart.series = chart.series.map((series, seriesIndex) => {
      const modernSeriesNode = modernSeriesNodes[seriesIndex] ?? null;
      if (!modernSeriesNode) {
        return series;
      }
      const valueColorsNode = getFirstLocalChild(modernSeriesNode, "valueColors");
      const valueColors = valueColorsNode
        ? Array.from(valueColorsNode.childNodes)
          .filter((node): node is Element => node.nodeType === Node.ELEMENT_NODE)
          .map((node) => resolveChartColorNode(findFirstChartColorElement(node) ?? node, themePalette))
          .filter((value): value is string => typeof value === "string" && value.length > 0)
        : [];
      const nextRaw = valueColors.length > 0
        ? {
            ...(series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : {}),
            valueColors
          }
        : series.raw;
      const seriesShapeProperties = getFirstLocalChild(modernSeriesNode, "spPr");
      if (!seriesShapeProperties) {
        return nextRaw === series.raw
          ? series
          : {
              ...series,
              raw: nextRaw
            };
      }
      const fillColor = resolveChartFillColor(seriesShapeProperties, themePalette);
      const lineStyle = resolveChartLineStyle(seriesShapeProperties, themePalette);
      const fallbackColor = fillColor ?? lineStyle.color ?? undefined;
      return {
        ...series,
        color: series.color ?? fallbackColor,
        lineColor: series.lineColor ?? lineStyle.color ?? fillColor ?? fallbackColor,
        lineWidthPx: series.lineWidthPx ?? (typeof lineStyle.widthPx === "number" ? lineStyle.widthPx : undefined),
        markerColor: series.markerColor ?? fallbackColor ?? series.color,
        markerLineColor: series.markerLineColor ?? lineStyle.color ?? fallbackColor ?? series.lineColor,
        raw: nextRaw
      };
    });

    const seriesLayouts = modernSeriesNodes.map((node) => node.getAttribute("layoutId") ?? node.getAttribute("layout"));
    const clusteredColumnIndex = seriesLayouts.findIndex((layout) => layout === "clusteredColumn");
    if (clusteredColumnIndex >= 0) {
      const clusteredNode = modernSeriesNodes[clusteredColumnIndex] ?? null;
      const parsedBinning = clusteredNode ? parseModernBinning(clusteredNode) : null;
      if (parsedBinning) {
        const syntheticRawSeries: Record<string, unknown> = {
          layoutId: "clusteredColumn",
          layoutPr: {
            binning: parsedBinning
          }
        };
        const hasParetoLine = seriesLayouts.includes("paretoLine");
        const replaceColumnSeries = (series: XlsxChartSeries | null | undefined) => (
          series ? buildChartExHistogramSeries(series, syntheticRawSeries, hasParetoLine) : null
        );

        if (chart.typeGroups && chart.typeGroups.length > 0) {
          const nextTypeGroups = chart.typeGroups.map((group) => ({ ...group, series: [...group.series] }));
          const columnGroupIndex = nextTypeGroups.findIndex((group) => group.chartType === "ColumnClustered");
          if (columnGroupIndex >= 0) {
            const originalColumnSeries = nextTypeGroups[columnGroupIndex]?.series[0] ?? null;
            const binnedColumnSeries = replaceColumnSeries(originalColumnSeries);
            if (binnedColumnSeries) {
              nextTypeGroups[columnGroupIndex].series = [binnedColumnSeries];
              const lineGroupIndex = nextTypeGroups.findIndex((group) => group.chartType === "Line");
              if (lineGroupIndex >= 0 && nextTypeGroups[lineGroupIndex]?.series[0]) {
                const originalLineSeries = nextTypeGroups[lineGroupIndex].series[0];
                const recomputedLine = buildChartExParetoLineSeries(
                  binnedColumnSeries,
                  {
                    text: originalLineSeries.name,
                    ...(originalLineSeries.raw && typeof originalLineSeries.raw === "object"
                      ? originalLineSeries.raw as Record<string, unknown>
                      : {})
                  },
                  0
                );
                nextTypeGroups[lineGroupIndex].series = [
                  {
                    ...originalLineSeries,
                    categories: recomputedLine.categories,
                    categoriesRef: recomputedLine.categoriesRef,
                    raw: recomputedLine.raw,
                    values: recomputedLine.values
                  }
                ];
                chart.series = [binnedColumnSeries, nextTypeGroups[lineGroupIndex].series[0]];
              } else {
                chart.series = [binnedColumnSeries];
              }
              chart.typeGroups = nextTypeGroups;
            }
          } else if (chart.series[0]) {
            const binnedSeries = replaceColumnSeries(chart.series[0]);
            if (binnedSeries) {
              chart.series = [binnedSeries];
            }
          }
        } else if (chart.series[0]) {
          const binnedSeries = replaceColumnSeries(chart.series[0]);
          if (binnedSeries) {
            chart.series = [binnedSeries];
          }
        }
      }
    }
  };

  const chartDocument = parseXml(chartXml);
  const chartNode = chartDocument ? getFirstLocalDescendant(chartDocument, "chart") : null;
  const plotAreaNode = chartNode ? getFirstLocalChild(chartNode, "plotArea") : null;
  const styleIdNode = chartDocument?.documentElement ? getFirstLocalDescendant(chartDocument.documentElement, "style") : null;
  const chartTypeNode = findPrimaryChartTypeNode(plotAreaNode);

  if (!chartNode || !chartTypeNode) {
    applyRelationshipStyles();
    const fallbackStyleId = readChartNumericAttribute(styleIdNode, "style");
    if (typeof fallbackStyleId === "number" && Number.isFinite(fallbackStyleId)) {
      chart.chartStyleId = fallbackStyleId;
    }
    applyModernChartExStyles();
    applyFallbackSeriesStyles();
    applyBuiltinChartDefaults(chart, themePalette);
    return;
  }
  const plotArea = plotAreaNode;
  if (!plotArea) {
    applyRelationshipStyles();
    applyFallbackSeriesStyles();
    applyBuiltinChartDefaults(chart, themePalette);
    return;
  }

  switch (chartTypeNode.localName) {
    case "barChart":
    case "bar3DChart": {
      const grouping = getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val");
      const barDir = getFirstLocalChild(chartTypeNode, "barDir")?.getAttribute("val");
      const isHorizontalBar = barDir === "bar";
      chart.is3d = chartTypeNode.localName === "bar3DChart" ? true : chart.is3d;
      if (grouping === "percentStacked") {
        chart.chartType = isHorizontalBar ? "BarPercentStacked" : "ColumnPercentStacked";
      } else if (grouping === "stacked") {
        chart.chartType = isHorizontalBar ? "BarStacked" : "ColumnStacked";
      } else {
        chart.chartType = isHorizontalBar ? "BarClustered" : "ColumnClustered";
      }
      break;
    }
    case "areaChart":
    case "area3DChart": {
      const grouping = getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val");
      chart.is3d = chartTypeNode.localName === "area3DChart" ? true : chart.is3d;
      if (grouping === "stacked") {
        chart.chartType = "AreaStacked";
      } else if (grouping === "percentStacked") {
        chart.chartType = "AreaPercentStacked";
      } else {
        chart.chartType = "Area";
      }
      break;
    }
    case "lineChart":
    case "line3DChart": {
      const grouping = getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val");
      chart.is3d = chartTypeNode.localName === "line3DChart" ? true : chart.is3d;
      if (grouping === "stacked") {
        chart.chartType = "LineStacked";
      } else if (grouping === "percentStacked") {
        chart.chartType = "LinePercentStacked";
      } else {
        chart.chartType = "Line";
      }
      break;
    }
    case "pieChart":
      chart.chartType = "Pie";
      break;
    case "pie3DChart":
      chart.chartType = "Pie3D";
      chart.is3d = true;
      break;
    case "doughnutChart":
      chart.chartType = "Doughnut";
      break;
    case "ofPieChart":
      chart.chartType = "BarOfPie";
      break;
    case "scatterChart":
      chart.chartType = resolveScatterChartType(getFirstLocalChild(chartTypeNode, "scatterStyle")?.getAttribute("val"));
      break;
    case "radarChart":
      chart.chartType = "Radar";
      break;
    case "surfaceChart":
      chart.chartType = "Surface";
      chart.is3d = false;
      break;
    case "surface3DChart":
      chart.chartType = "Surface";
      chart.is3d = true;
      break;
    case "stockChart":
      chart.chartType = "Stock";
      break;
    case "bubbleChart":
      chart.chartType = "Bubble";
      break;
    default:
      break;
  }

  const legendNode = getFirstLocalChild(chartNode, "legend");
  const legendPosition = legendNode ? getFirstLocalChild(legendNode, "legendPos")?.getAttribute("val") ?? undefined : undefined;
  const legendOverlay = legendNode ? getFirstLocalChild(legendNode, "overlay")?.getAttribute("val") : undefined;

  chart.legend = legendNode ? {
    overlay: legendOverlay === "1",
    position: normalizeLegendPosition(legendPosition),
    raw: chart.legend?.raw
  } : chart.legend;
  const plotVisibleOnly = readChartBooleanAttribute(chartNode, "plotVisOnly");
  if (plotVisibleOnly !== undefined) {
    chart.plotVisibleOnly = plotVisibleOnly;
  }
  chart.displayBlanksAs = getFirstLocalChild(chartNode, "dispBlanksAs")?.getAttribute("val") ?? chart.displayBlanksAs;
  const styleId = Number(styleIdNode?.getAttribute("val") ?? Number.NaN);
  chart.chartStyleId = Number.isFinite(styleId) ? styleId : chart.chartStyleId;
  chart.firstSliceAngle = readChartNumericAttribute(chartTypeNode, "firstSliceAng") ?? chart.firstSliceAngle;
  chart.gapWidth = readChartNumericAttribute(chartTypeNode, "gapWidth") ?? chart.gapWidth;
  chart.overlap = readChartNumericAttribute(chartTypeNode, "overlap") ?? chart.overlap;
  chart.bubbleScale = readChartNumericAttribute(chartTypeNode, "bubbleScale") ?? chart.bubbleScale;
  chart.varyColors = readChartBooleanAttribute(chartTypeNode, "varyColors") ?? chart.varyColors;
  const bubble3dNode = getFirstLocalChild(chartTypeNode, "bubble3D");
  chart.bubble3d = bubble3dNode
    ? bubble3dNode.getAttribute("val") !== "0"
    : chart.bubble3d;
  chart.holeSize = readChartNumericAttribute(chartTypeNode, "holeSize") ?? chart.holeSize;
  chart.radarStyle = getFirstLocalChild(chartTypeNode, "radarStyle")?.getAttribute("val") ?? chart.radarStyle;
  chart.scatterStyle = getFirstLocalChild(chartTypeNode, "scatterStyle")?.getAttribute("val") ?? chart.scatterStyle;
  chart.shape3d = getFirstLocalChild(chartTypeNode, "shape")?.getAttribute("val") ?? chart.shape3d;
  const wireframeNode = getFirstLocalChild(chartTypeNode, "wireframe");
  chart.wireframe = wireframeNode
    ? wireframeNode.getAttribute("val") !== "0"
    : chart.wireframe;
  const chartTypeDataLabels = parseChartDataLabelsFromXml(getFirstLocalChild(chartTypeNode, "dLbls"));
  const firstSeriesNode = getLocalChildren(chartTypeNode, "ser")[0] ?? null;
  const seriesDataLabels = parseChartDataLabelsFromXml(getFirstLocalChild(firstSeriesNode, "dLbls"));
  chart.dataLabels = chartTypeDataLabels ?? seriesDataLabels ?? chart.dataLabels;
  const seriesSp3dNode = firstSeriesNode ? getFirstLocalDescendant(firstSeriesNode, "sp3d") : null;
  chart.surfaceMaterial = seriesSp3dNode?.getAttribute("prstMaterial") ?? chart.surfaceMaterial;
  const bandFormatsNode = getLocalChildren(chartTypeNode, "bandFmts")[0] ?? null;
  const bandFormatNodes = bandFormatsNode ? getLocalChildren(bandFormatsNode, "bandFmt") : [];
  const bandFormatColors = bandFormatNodes
    .map((bandFormatNode) => {
      const shapeProperties = getFirstLocalChild(bandFormatNode, "spPr");
      return resolveChartFillColor(shapeProperties, themePalette) ?? undefined;
    })
    .filter((color): color is string => typeof color === "string" && color.length > 0);
  const bandFormatLineColors = bandFormatNodes
    .map((bandFormatNode) => {
      const shapeProperties = getFirstLocalChild(bandFormatNode, "spPr");
      return resolveChartLineStyle(shapeProperties, themePalette).color ?? undefined;
    })
    .filter((color): color is string => typeof color === "string" && color.length > 0);

  chart.raw = {
    ...(chart.raw ?? {}),
    bandFormatCount: bandFormatNodes.length > 0 ? bandFormatNodes.length : undefined,
    bandFormatColors: bandFormatColors.length > 0 ? bandFormatColors : undefined,
    bandFormatLineColors: bandFormatLineColors.length > 0 ? bandFormatLineColors : undefined,
    date1904: readChartBooleanAttribute(chartDocument?.documentElement ?? null, "date1904"),
    bubble3d: chart.bubble3d,
    grouping: getFirstLocalChild(chartTypeNode, "grouping")?.getAttribute("val") ?? undefined,
    ofPieType: getFirstLocalChild(chartTypeNode, "ofPieType")?.getAttribute("val") ?? undefined,
    shape: getFirstLocalChild(chartTypeNode, "shape")?.getAttribute("val") ?? undefined,
    secondPieSize: readChartNumericAttribute(chartTypeNode, "secondPieSize"),
    scatterStyle: chart.scatterStyle,
    splitPos: readChartNumericAttribute(chartTypeNode, "splitPos"),
    splitType: getFirstLocalChild(chartTypeNode, "splitType")?.getAttribute("val") ?? undefined,
    xmlChartType: chartTypeNode.localName
  };
  const view3dNode = getFirstLocalDescendant(chartNode, "view3D");
  if (view3dNode) {
    chart.view3d = {
      depthPercent: readChartNumericAttribute(view3dNode, "depthPercent"),
      perspective: readChartNumericAttribute(view3dNode, "perspective"),
      rAngAx: getFirstLocalChild(view3dNode, "rAngAx")?.getAttribute("val") === "1",
      rotX: readChartNumericAttribute(view3dNode, "rotX"),
      rotY: readChartNumericAttribute(view3dNode, "rotY")
    };
  }
  chart.floor = readChartWallFromXml(getFirstLocalChild(chartNode, "floor"), themePalette) ?? chart.floor;
  chart.sideWall = readChartWallFromXml(getFirstLocalChild(chartNode, "sideWall"), themePalette) ?? chart.sideWall;
  chart.backWall = readChartWallFromXml(getFirstLocalChild(chartNode, "backWall"), themePalette) ?? chart.backWall;

  const styleAppearance = applyRelationshipStyles();
  const chartTextTypeface = readChartTextTypeface(getFirstLocalChild(chartNode, "txPr"), themePalette);
  const titleTypeface = readChartTextTypeface(getFirstLocalDescendant(chartNode, "title"), themePalette);
  chart.fontFamily = chartTextTypeface ?? chart.fontFamily;
  chart.titleFontFamily = titleTypeface ?? chart.titleFontFamily ?? chart.fontFamily;

  const chartAreaShapeProperties = chartDocument?.documentElement
    ? getFirstLocalChild(chartDocument.documentElement, "spPr")
    : null;
  const plotAreaShapeProperties = getFirstLocalChild(plotArea, "spPr");
  const chartAreaNoFill = chartAreaShapeProperties ? getFirstLocalChild(chartAreaShapeProperties, "noFill") != null : false;
  const plotAreaNoFill = plotAreaShapeProperties ? getFirstLocalChild(plotAreaShapeProperties, "noFill") != null : false;
  chart.raw = {
    ...(chart.raw ?? {}),
    chartAreaNoFill: styleAppearance.chartAreaNoFill === true || chartAreaNoFill,
    plotAreaNoFill
  };
  if (chartAreaShapeProperties) {
    const chartAreaFillColor = resolveChartFillColor(chartAreaShapeProperties, themePalette);
    if (chartAreaFillColor) {
      chart.chartAreaFillColor = chartAreaFillColor;
    } else if (getFirstLocalChild(chartAreaShapeProperties, "noFill")) {
      chart.chartAreaFillColor = "transparent";
    }
    const chartAreaLineStyle = resolveChartLineStyle(chartAreaShapeProperties, themePalette);
    if (chartAreaLineStyle.hidden) {
      chart.chartAreaBorderColor = "transparent";
    } else if (chartAreaLineStyle.color) {
      chart.chartAreaBorderColor = chartAreaLineStyle.color;
    }
  }
  if (!chart.chartAreaFillColor && (styleAppearance.chartAreaNoFill === true || plotAreaNoFill)) {
    chart.chartAreaFillColor = "transparent";
  }
  const categoryAxisNodes = [
    ...getLocalChildren(plotArea, "catAx"),
    ...getLocalChildren(plotArea, "dateAx")
  ];
  const valueAxisNodes = getLocalChildren(plotArea, "valAx");
  const seriesAxisNode = getLocalChildren(plotArea, "serAx")[0] ?? null;
  const isScatterLikeChart = (
    chart.chartType === "Scatter"
    || chart.chartType === "ScatterLines"
    || chart.chartType === "ScatterSmooth"
    || chart.chartType === "Bubble"
  );
  let categoryAxisNode = categoryAxisNodes[0] ?? null;
  let valueAxisNode = valueAxisNodes[0] ?? null;
  if (!categoryAxisNode && isScatterLikeChart && valueAxisNodes.length >= 2) {
    categoryAxisNode = valueAxisNodes.find((axisNode) => {
      const position = getFirstLocalChild(axisNode, "axPos")?.getAttribute("val");
      return position === "b" || position === "t";
    }) ?? valueAxisNodes[0];
    valueAxisNode = valueAxisNodes.find((axisNode) => {
      const position = getFirstLocalChild(axisNode, "axPos")?.getAttribute("val");
      return position === "l" || position === "r";
    }) ?? valueAxisNodes[1] ?? valueAxisNodes[0];
  }
  chart.categoryAxis = mergeChartAxis(chart.categoryAxis, readChartAxisFromXml(categoryAxisNode));
  chart.valueAxis = mergeChartAxis(chart.valueAxis, readChartAxisFromXml(valueAxisNode));
  chart.seriesAxis = mergeChartAxis(chart.seriesAxis, readChartAxisFromXml(seriesAxisNode));
  chart.axes = chart.axes.length > 0
    ? chart.axes.map((axis, index) => (
      index === 0 && categoryAxisNode
        ? { ...axis, ...readChartAxisFromXml(categoryAxisNode) }
        : index === 1 && valueAxisNode
          ? { ...axis, ...readChartAxisFromXml(valueAxisNode) }
          : axis
    ))
    : chart.axes;
  if (seriesAxisNode) {
    const seriesAxis = readChartAxisFromXml(seriesAxisNode);
    if (seriesAxis && !chart.axes.some((axis) => axis.id != null && axis.id === seriesAxis.id)) {
      chart.axes = [...chart.axes, seriesAxis as XlsxChartAxis];
    }
  }

  applyChartSeriesStyleFromXml(chart, chartTypeNode, themePalette);
  applyFallbackSeriesStyles();
  if (chart.chartType === "Bubble") {
    const archiveFallbackBubbleSizes = resolveArchiveFallbackBubbleSizes(archive, chart.title);
    if (archiveFallbackBubbleSizes.length > 0) {
      chart.series = chart.series.map((series, seriesIndex) => {
        const pointCount = Math.max(series.values.length, series.categories.length);
        if (pointCount <= 1) {
          return series;
        }

        const numericBubbleCount = (series.bubbleSizes ?? []).filter(
          (value): value is number => typeof value === "number" && Number.isFinite(value)
        ).length;
        if (numericBubbleCount >= pointCount) {
          return series;
        }

        const fallbackCandidate = archiveFallbackBubbleSizes[seriesIndex] ?? archiveFallbackBubbleSizes[0] ?? [];
        const fallbackNumericCount = fallbackCandidate.filter(
          (value): value is number => typeof value === "number" && Number.isFinite(value)
        ).length;
        if (fallbackNumericCount < pointCount) {
          return series;
        }

        return {
          ...series,
          bubbleSizes: fallbackCandidate
        };
      });
    }
  }
  if (chart.chartType === "Pie" || chart.chartType === "Pie3D" || chart.chartType === "PieExploded" || chart.chartType === "Doughnut" || chart.chartType === "BarOfPie") {
    const needsPointColorFallback = chart.series.some((series) => {
      const pointCount = Math.max(series.values.length, series.categories.length);
      if (pointCount <= 0) {
        return false;
      }
      const coloredPointCount = (series.dataPointStyles ?? []).filter(
        (style) => typeof style.color === "string" && style.color.length > 0
      ).length;
      return coloredPointCount === 0;
    });
    if (needsPointColorFallback) {
      const archiveFallbackPointStyles = resolveArchiveFallbackPointStyles(
        archive,
        chart.title,
        chartTypeNode.localName,
        themePalette
      );
      if (archiveFallbackPointStyles.length > 0) {
        chart.series = chart.series.map((series, seriesIndex) => {
          const fallbackStyles = archiveFallbackPointStyles[seriesIndex] ?? archiveFallbackPointStyles[0] ?? [];
          if (fallbackStyles.length === 0) {
            return series;
          }
          const existingByIndex = new Map((series.dataPointStyles ?? []).map((entry) => [entry.index, entry]));
          for (const fallbackStyle of fallbackStyles) {
            const existing = existingByIndex.get(fallbackStyle.index);
            existingByIndex.set(fallbackStyle.index, {
              color: existing?.color ?? fallbackStyle.color,
              explosion: existing?.explosion ?? fallbackStyle.explosion,
              index: fallbackStyle.index,
              lineColor: existing?.lineColor ?? fallbackStyle.lineColor
            });
          }
          return {
            ...series,
            dataPointStyles: Array.from(existingByIndex.values()).sort((left, right) => left.index - right.index)
          };
        });
      }
    }
  }
  applyBuiltinChartDefaults(chart, themePalette);
}


export function normalizeArchivePath(path: string) {
  return path.replace(/^\/+/, "").replace(/\\/g, "/");
}

export function dirname(path: string) {
  const normalized = normalizeArchivePath(path);
  const index = normalized.lastIndexOf("/");
  return index >= 0 ? normalized.slice(0, index) : "";
}

export function resolveRelationshipPath(basePath: string, target: string) {
  if (!target) {
    return "";
  }

  const normalizedTarget = target.replace(/\\/g, "/");
  if (normalizedTarget.startsWith("/")) {
    return normalizeArchivePath(normalizedTarget);
  }
  const normalizedBasePath = normalizeArchivePath(basePath);
  let baseDirectory = dirname(normalizedBasePath);
  if (normalizedBasePath.endsWith(".rels")) {
    const relsMarker = "/_rels/";
    const relsMarkerIndex = normalizedBasePath.lastIndexOf(relsMarker);
    if (relsMarkerIndex >= 0) {
      const ownerPrefix = normalizedBasePath.slice(0, relsMarkerIndex);
      const relFileName = normalizedBasePath.slice(relsMarkerIndex + relsMarker.length);
      const ownerFileName = relFileName.endsWith(".rels")
        ? relFileName.slice(0, -".rels".length)
        : relFileName;
      baseDirectory = dirname(`${ownerPrefix}/${ownerFileName}`);
    }
  }

  const segments = [...baseDirectory.split("/").filter(Boolean), ...normalizedTarget.split("/").filter(Boolean)];
  const resolved: string[] = [];
  for (const segment of segments) {
    if (segment === ".") {
      continue;
    }
    if (segment === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }

  return resolved.join("/");
}

export function readArchiveText(archive: Record<string, Uint8Array>, path: string | null | undefined) {
  if (!path) {
    return null;
  }

  const entry = archive[normalizeArchivePath(path)];
  return entry ? strFromU8(entry) : null;
}

export function parseXml(xml: string) {
  if (typeof DOMParser === "undefined") {
    return null;
  }

  try {
    return new DOMParser().parseFromString(xml, "application/xml");
  } catch {
    return null;
  }
}

export function serializeXml(document: XMLDocument) {
  return new XMLSerializer().serializeToString(document);
}

export function getLocalChildren(parent: ParentNode, localName: string) {
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === localName
  );
}

export function removeLocalChildren(parent: ParentNode, localName: string) {
  getLocalChildren(parent, localName).forEach((node) => node.parentNode?.removeChild(node));
}

export function getLocalDescendants(parent: ParentNode, localName: string) {
  return Array.from((parent as Element | Document).getElementsByTagName("*")).filter(
    (node) => node.localName === localName
  );
}

export function getFirstLocalChild(parent: ParentNode, localName: string) {
  return getLocalChildren(parent, localName)[0] ?? null;
}

export function getFirstLocalDescendant(parent: ParentNode, localName: string) {
  return getLocalDescendants(parent, localName)[0] ?? null;
}

export function ensureChild(parent: Element, localName: string, namespace = parent.namespaceURI ?? CHART_NS, prefix = "c") {
  const existing = getFirstLocalChild(parent, localName);
  if (existing) {
    return existing;
  }

  const document = parent.ownerDocument;
  const node = document.createElementNS(namespace, `${prefix}:${localName}`);
  parent.appendChild(node);
  return node;
}

export function setLeafValue(parent: Element, localName: string, value: string, namespace = parent.namespaceURI ?? CHART_NS, prefix = "c") {
  const node = ensureChild(parent, localName, namespace, prefix);
  node.textContent = value;
  return node;
}

export function setBooleanValue(parent: Element, localName: string, value: boolean) {
  const node = ensureChild(parent, localName);
  node.setAttribute("val", value ? "1" : "0");
  return node;
}

export function setNumericValue(parent: Element, localName: string, value: number) {
  const node = ensureChild(parent, localName);
  node.setAttribute("val", String(Math.round(value)));
  return node;
}


export function parseChartCacheValues(
  parentNode: Element | null,
  cacheName: "numCache" | "strCache",
  mode: "category" | "value"
): Array<number | string | null> | null {
  if (!parentNode) {
    return null;
  }

  const referenceNode = getFirstLocalChild(parentNode, "numRef")
    ?? getFirstLocalChild(parentNode, "strRef")
    ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, cacheName);
  if (!cacheNode) {
    return null;
  }

  const pointCount = readChartNumericAttribute(cacheNode, "ptCount");
  const pointNodes = getLocalChildren(cacheNode, "pt")
    .map((pointNode) => {
      const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
      return {
        index: Number.isFinite(rawIndex) ? rawIndex : 0,
        value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
      };
    })
    .sort((left, right) => left.index - right.index);

  if (pointNodes.length === 0) {
    return null;
  }

  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  for (const point of pointNodes) {
    if (mode === "value") {
      values[point.index] = cellValueToNumber(point.value);
    } else {
      values[point.index] = point.value.length > 0 ? point.value : null;
    }
  }
  return values;
}

export function parseChartMultiLevelCacheValues(
  parentNode: Element | null,
  mode: "category" | "value"
): Array<number | string | null> | null {
  if (!parentNode) {
    return null;
  }

  const referenceNode = getFirstLocalChild(parentNode, "multiLvlStrRef") ?? parentNode;
  const cacheNode = getFirstLocalChild(referenceNode, "multiLvlStrCache");
  if (!cacheNode) {
    return null;
  }

  const levelNodes = getLocalChildren(cacheNode, "lvl");
  if (levelNodes.length === 0) {
    return null;
  }

  const pointCount = readChartNumericAttribute(cacheNode, "ptCount");
  const primaryLevelNode = mode === "category"
    ? levelNodes[levelNodes.length - 1] ?? levelNodes[0]
    : levelNodes[0];
  const pointNodes = getLocalChildren(primaryLevelNode, "pt")
    .map((pointNode) => {
      const rawIndex = Number(pointNode.getAttribute("idx") ?? Number.NaN);
      return {
        index: Number.isFinite(rawIndex) ? rawIndex : 0,
        value: getFirstLocalChild(pointNode, "v")?.textContent ?? ""
      };
    })
    .sort((left, right) => left.index - right.index);

  if (pointNodes.length === 0) {
    return null;
  }

  const maxIndex = pointNodes.reduce((max, point) => Math.max(max, point.index), 0);
  const targetLength = Math.max(
    pointNodes.length,
    Number.isFinite(pointCount ?? Number.NaN) ? Number(pointCount) : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: targetLength }, () => null as number | string | null);
  for (const point of pointNodes) {
    if (mode === "value") {
      values[point.index] = cellValueToNumber(point.value);
      continue;
    }
    values[point.index] = point.value.length > 0 ? point.value : null;
  }
  return values;
}

export function loadWorkbookChartAssets(
  workbook: Workbook,
  imageAssets: Pick<WorkbookImageAssets, "archive" | "sheetOrigins" | "themePalette"> | null,
  visibleSheetIndexByWorkbookSheetIndex: Map<number, number>,
  showHiddenSheets = false
): WorkbookChartAssets {
  const chartsByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, (_, workbookSheetIndex) => {
    const worksheet = workbook.getSheet(workbookSheetIndex);
    const rawCharts = Array.isArray(worksheet.charts) ? worksheet.charts : [];
    const rawChartsEx = Array.isArray(worksheet.chartsEx) ? worksheet.chartsEx : [];
    const visibleSheetIndex = visibleSheetIndexByWorkbookSheetIndex.get(workbookSheetIndex) ?? workbookSheetIndex;

    const classicCharts = rawCharts.map((rawChart, chartIndex) => {
      const chartId = `chart-${workbookSheetIndex}-${chartIndex}`;
      const chart = rawChart && typeof rawChart === "object" ? rawChart as Record<string, unknown> : {};
      const rawView3d = chart.view3d && typeof chart.view3d === "object"
        ? chart.view3d as Record<string, unknown>
        : null;
      const rawSeries = Array.isArray(chart.series) ? chart.series : [];
      const chartLevelDataLabels = normalizeChartDataLabels(chart.dataLabels);
      const firstSeriesDataLabels = rawSeries.length > 0 && rawSeries[0] && typeof rawSeries[0] === "object"
        ? normalizeChartDataLabels((rawSeries[0] as Record<string, unknown>).dataLabels)
        : null;
      return {
        anchor: normalizeChartAnchor(chart.anchor),
        autoTitleDeleted: typeof chart.autoTitleDeleted === "boolean" ? chart.autoTitleDeleted : undefined,
        axes: Array.isArray(chart.axes) ? chart.axes.map(normalizeChartAxis).filter((value): value is XlsxChartAxis => Boolean(value)) : [],
        axisLabelColor: undefined,
        axisLineColor: undefined,
        categoryAxis: normalizeChartAxis(chart.categoryAxis),
        chartAreaBorderColor: undefined,
        chartAreaFillColor: undefined,
        chartColorPalette: undefined,
        chartColorPaletteOffset: undefined,
        chartPath: undefined,
        chartStyleId: undefined,
        chartType: typeof chart.chartType === "string" ? chart.chartType : "ColumnClustered",
        dataLabels: chartLevelDataLabels ?? firstSeriesDataLabels,
        displayBlanksAs: typeof chart.displayBlanksAs === "string" ? chart.displayBlanksAs : undefined,
        editable: true,
        firstSliceAngle: typeof chart.firstSliceAngle === "number" ? chart.firstSliceAngle : undefined,
        fontFamily: undefined,
        gapWidth: typeof chart.gapWidth === "number" ? chart.gapWidth : undefined,
        holeSize: typeof chart.holeSize === "number" ? chart.holeSize : undefined,
        id: chartId,
        is3d: typeof chart.is3d === "boolean" ? chart.is3d : undefined,
        legend: normalizeLegend(chart.legend)
          ? {
              ...normalizeLegend(chart.legend),
              position: normalizeLegendPosition(normalizeLegend(chart.legend)?.position)
            }
          : null,
        name: typeof chart.name === "string" ? chart.name : undefined,
        overlap: typeof chart.overlap === "number" ? chart.overlap : undefined,
        plotVisibleOnly: typeof chart.plotVisibleOnly === "boolean" ? chart.plotVisibleOnly : undefined,
        raw: chart,
        radarStyle: typeof chart.radarStyle === "string" ? chart.radarStyle : undefined,
        scatterStyle: typeof chart.scatterStyle === "string" ? chart.scatterStyle : undefined,
        roundedCorners: typeof chart.roundedCorners === "boolean" ? chart.roundedCorners : undefined,
        shape3d: typeof chart.shape === "string"
          ? chart.shape
          : typeof chart.shape3d === "string"
            ? chart.shape3d
            : undefined,
        seriesAxis: null,
        series: rawSeries.map((entry, seriesIndex) => normalizeChartSeries(workbook, workbookSheetIndex, chartId, entry, seriesIndex)),
        sheetIndex: visibleSheetIndex,
        showDlblsOverMax: typeof chart.showDlblsOverMax === "boolean" ? chart.showDlblsOverMax : undefined,
        sideWall: null,
        backWall: null,
        bubbleScale: typeof chart.bubbleScale === "number" ? chart.bubbleScale : undefined,
        bubble3d: typeof chart.bubble3d === "boolean" ? chart.bubble3d : undefined,
        floor: null,
        surfaceMaterial: undefined,
        textColor: undefined,
        title: typeof chart.title === "string" ? chart.title : undefined,
        titleColor: undefined,
        titleFontFamily: undefined,
        typeGroups: Array.isArray(chart.typeGroups)
          ? chart.typeGroups
            .map((entry, groupIndex) => normalizeChartTypeGroup(workbook, workbookSheetIndex, chartId, entry, groupIndex))
            .filter((value): value is XlsxChartTypeGroup => value != null)
          : [],
        valueAxis: normalizeChartAxis(chart.valueAxis),
        varyColors: typeof chart.varyColors === "boolean" ? chart.varyColors : undefined,
        view3d: rawView3d
          ? {
              depthPercent: typeof rawView3d.depthPercent === "number" ? rawView3d.depthPercent : undefined,
              perspective: typeof rawView3d.perspective === "number" ? rawView3d.perspective : undefined,
              rAngAx: typeof rawView3d.rAngAx === "boolean"
                ? rawView3d.rAngAx
                : typeof rawView3d.rightAngleAxes === "boolean"
                  ? rawView3d.rightAngleAxes
                  : undefined,
              rotX: typeof rawView3d.rotX === "number"
                ? rawView3d.rotX
                : typeof rawView3d.rotateX === "number"
                  ? rawView3d.rotateX
                  : undefined,
              rotY: typeof rawView3d.rotY === "number"
                ? rawView3d.rotY
                : typeof rawView3d.rotateY === "number"
                  ? rawView3d.rotateY
                  : undefined
            }
          : undefined,
        wireframe: typeof chart.wireframe === "boolean" ? chart.wireframe : undefined,
        workbookSheetIndex,
        zIndex: 200 + chartIndex
      } satisfies XlsxChart;
    });

    const modernCharts = rawChartsEx.map((rawChartEx, chartExIndex) => (
      normalizeChartExChart(
        workbook,
        workbookSheetIndex,
        visibleSheetIndex,
        rawChartEx,
        chartExIndex,
        imageAssets?.themePalette ?? null
      )
    ));

    return [...classicCharts, ...modernCharts];
  });

  const chartsheets = Array.isArray(workbook.chartsheets)
    ? workbook.chartsheets.map((entry, index) => normalizeChartsheet(entry, index))
    : [];
  const tabs = buildTabs(workbook, chartsheets, visibleSheetIndexByWorkbookSheetIndex, showHiddenSheets);
  const chartOriginsById = new Map<string, WorkbookChartOrigin>();

  if (imageAssets) {
    applyChartOrigins(chartsByWorkbookSheetIndex, chartOriginsById, imageAssets.archive, imageAssets.sheetOrigins);
    for (const charts of chartsByWorkbookSheetIndex) {
      for (const chart of charts) {
        applyChartStyleFromXml(chart, chart.chartPath, imageAssets.archive, imageAssets.themePalette);
        applyBuiltinChartDefaults(chart, imageAssets.themePalette);
      }
    }
  } else {
    for (const charts of chartsByWorkbookSheetIndex) {
      for (const chart of charts) {
        applyBuiltinChartDefaults(chart, null);
      }
    }
  }

  return {
    chartOriginsById,
    chartsByWorkbookSheetIndex,
    chartsheets,
    tabs
  };
}

