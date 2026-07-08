import { strFromU8 } from "fflate";
import type { XlsxChart, XlsxChartPointStyle, XlsxThemePalette } from "../types";

import {
  CHART_NS,
  SERIES_COLORS,
} from "./chart-series";

import {
  isElementNode,
  normalizeHexColor,
  readHexColorFromXmlFragment,
  resolveColorFromXmlFragment,
  resolveChartColorNode,
  findFirstChartColorElement,
  resolveChartFillColor,
  resolveChartLineStyle,
  applyLightnessTransform,
  readChartTextTypeface,
} from "./chart-colors";

import {
  parseChartTypeFromXml,
  normalizeChartTitleForMatch,
  extractChartTitleFromXml,
  readChartLabelFontSizePt,
} from "./chart-parser";

import {
  parseChartCacheValues,
  parseChartMultiLevelCacheValues,
} from "./chart-cache";

import {
  readChartNumericAttribute,
  readChartBooleanAttribute,
  getLocalChildren,
  getFirstLocalChild,
  getFirstLocalDescendant,
  getLocalDescendants,
  parseXml,
  normalizeArchivePath,
  dirname,
  readArchiveText,
  resolveRelationshipPath,
} from "./chart-xml-utils";

import type { ChartStyleAppearance } from "./chart-types";

export type FallbackSeriesStyle = {
  color?: string;
  lineColor?: string;
};

export function parseFallbackSeriesStylesFromChartXml(
  chartXml: string,
  themePalette?: XlsxThemePalette | null
): FallbackSeriesStyle[] {
  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }

  return seriesBlocks.map((seriesBlock) => {
    const shapeBlock = seriesBlock.match(/<c:spPr\b[\s\S]*?<\/c:spPr>/i)?.[0] ?? "";
    return {
      color: readHexColorFromXmlFragment(shapeBlock, false, themePalette),
      lineColor: readHexColorFromXmlFragment(shapeBlock, true, themePalette)
    };
  });
}

export function parseFallbackPointStylesFromChartXml(
  chartXml: string,
  themePalette?: XlsxThemePalette | null
): XlsxChartPointStyle[][] {
  const chartDocument = parseXml(chartXml);
  if (chartDocument) {
    const parsedSeriesStyles = getLocalDescendants(chartDocument, "ser").map((seriesNode) => {
      const styles: XlsxChartPointStyle[] = [];
      for (const dataPointNode of getLocalChildren(seriesNode, "dPt")) {
        const indexValue = readChartNumericAttribute(dataPointNode, "idx");
        if (indexValue === undefined) {
          continue;
        }
        const shapeProperties = getFirstLocalChild(dataPointNode, "spPr");
        const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
        styles.push({
          color: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
          explosion: readChartNumericAttribute(dataPointNode, "explosion"),
          index: indexValue,
          lineColor: lineStyle.color ?? undefined
        });
      }
      return styles;
    });
    if (parsedSeriesStyles.some((styles) => styles.length > 0)) {
      return parsedSeriesStyles;
    }
  }

  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }

  return seriesBlocks.map((seriesBlock) => {
    const pointBlocks = seriesBlock.match(/<c:dPt\b[\s\S]*?<\/c:dPt>/gi) ?? [];
    if (pointBlocks.length === 0) {
      return [];
    }

    const styles: XlsxChartPointStyle[] = [];
    for (const pointBlock of pointBlocks) {
      const indexMatch = pointBlock.match(/<c:idx\b[^>]*\bval="(-?\d+)"/i);
      const index = indexMatch?.[1] ? Number(indexMatch[1]) : Number.NaN;
      if (!Number.isFinite(index)) {
        continue;
      }
      const explosionMatch = pointBlock.match(/<c:explosion\b[^>]*\bval="(-?\d+(?:\.\d+)?)"/i);
      const explosionValue = explosionMatch?.[1] ? Number(explosionMatch[1]) : Number.NaN;
      styles.push({
        color: readHexColorFromXmlFragment(pointBlock, false, themePalette),
        explosion: Number.isFinite(explosionValue) ? explosionValue : undefined,
        index,
        lineColor: readHexColorFromXmlFragment(pointBlock, true, themePalette)
      });
    }

    return styles;
  });
}

export function parseNumericPointCacheFromXmlFragment(fragment: string) {
  const pointMatches = Array.from(fragment.matchAll(/<c:pt\b[^>]*\bidx="(-?\d+)"[^>]*>[\s\S]*?<c:v>([^<]*)<\/c:v>[\s\S]*?<\/c:pt>/gi));
  if (pointMatches.length === 0) {
    return [];
  }

  const explicitPointCountMatch = fragment.match(/<c:ptCount\b[^>]*\bval="(\d+)"/i);
  const explicitPointCount = explicitPointCountMatch?.[1] ? Number(explicitPointCountMatch[1]) : Number.NaN;
  const maxIndex = pointMatches.reduce((max, match) => {
    const current = Number(match[1] ?? Number.NaN);
    return Number.isFinite(current) ? Math.max(max, current) : max;
  }, -1);
  const pointCount = Math.max(
    pointMatches.length,
    Number.isFinite(explicitPointCount) ? explicitPointCount : 0,
    maxIndex + 1
  );
  const values = Array.from({ length: pointCount }, () => null as number | null);

  for (const match of pointMatches) {
    const index = Number(match[1] ?? Number.NaN);
    const rawValue = (match[2] ?? "").trim();
    const numericValue = Number(rawValue);
    if (!Number.isFinite(index) || index < 0 || !Number.isFinite(numericValue)) {
      continue;
    }
    values[index] = numericValue;
  }

  return values;
}

export function parseFallbackBubbleSizesFromChartXml(chartXml: string): Array<Array<number | null>> {
  const seriesBlocks = chartXml.match(/<c:ser\b[\s\S]*?<\/c:ser>/gi) ?? [];
  if (seriesBlocks.length === 0) {
    return [];
  }

  return seriesBlocks.map((seriesBlock) => {
    const bubbleSizeBlock = seriesBlock.match(/<c:bubbleSize\b[\s\S]*?<\/c:bubbleSize>/i)?.[0] ?? "";
    if (!bubbleSizeBlock) {
      return [];
    }

    return parseNumericPointCacheFromXmlFragment(bubbleSizeBlock);
  });
}

export function resolveArchiveFallbackBubbleSizes(
  archive: Record<string, Uint8Array>,
  preferredTitle: string | undefined
) {
  const preferred = normalizeChartTitleForMatch(preferredTitle);
  let bestScore = Number.NEGATIVE_INFINITY;
  let bestCandidate: Array<Array<number | null>> = [];

  for (const [path, bytes] of Object.entries(archive)) {
    if (!/\/charts\/chart\d+\.xml$/i.test(path)) {
      continue;
    }
    const chartXml = strFromU8(bytes);
    if (!/<c:bubbleChart\b/i.test(chartXml)) {
      continue;
    }
    const candidateBubbleSizes = parseFallbackBubbleSizesFromChartXml(chartXml);
    const hasCandidateValues = candidateBubbleSizes.some((seriesValues) => seriesValues.some((value) => value != null));
    if (!hasCandidateValues) {
      continue;
    }

    let score = 0;
    const candidateTitle = normalizeChartTitleForMatch(extractChartTitleFromXml(chartXml));
    if (preferred.length > 0 && candidateTitle.length > 0 && preferred === candidateTitle) {
      score += 100;
    }
    if (bestCandidate.length === 0) {
      score += 1;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCandidate = candidateBubbleSizes;
      if (score >= 100) {
        break;
      }
    }
  }

  return bestCandidate;
}

export function resolveArchiveFallbackPointStyles(
  archive: Record<string, Uint8Array>,
  preferredTitle: string | undefined,
  preferredChartXmlType: string | undefined,
  themePalette?: XlsxThemePalette | null
) {
  const preferred = normalizeChartTitleForMatch(preferredTitle);
  const preferredType = (preferredChartXmlType ?? "").trim();
  let bestScore = Number.NEGATIVE_INFINITY;
  let bestCandidate: XlsxChartPointStyle[][] = [];

  for (const [path, bytes] of Object.entries(archive)) {
    if (!/\/charts\/chart\d+\.xml$/i.test(path)) {
      continue;
    }
    const chartXml = strFromU8(bytes);
    const candidateType = parseChartTypeFromXml(chartXml);
    if (!candidateType) {
      continue;
    }
    const candidatePointStyles = parseFallbackPointStylesFromChartXml(chartXml, themePalette);
    const hasCandidateValues = candidatePointStyles.some((seriesStyles) => seriesStyles.some((style) => (
      (typeof style.color === "string" && style.color.length > 0)
      || typeof style.explosion === "number"
    )));
    if (!hasCandidateValues) {
      continue;
    }

    let score = 0;
    const candidateTitle = normalizeChartTitleForMatch(extractChartTitleFromXml(chartXml));
    if (preferred.length > 0 && candidateTitle.length > 0 && preferred === candidateTitle) {
      score += 100;
    }
    if (preferredType && candidateType === preferredType) {
      score += 20;
    }
    if (bestCandidate.length === 0) {
      score += 1;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCandidate = candidatePointStyles;
      if (score >= 120) {
        break;
      }
    }
  }

  return bestCandidate;
}

export function readChartRelationships(
  archive: Record<string, Uint8Array>,
  chartPath: string
) {
  const relsPath = normalizeArchivePath(`${dirname(chartPath)}/_rels/${chartPath.split("/").pop()}.rels`);
  const relsXml = readArchiveText(archive, relsPath);
  if (!relsXml) {
    return new Map<string, string>();
  }

  const relsDocument = parseXml(relsXml);
  if (!relsDocument) {
    return new Map<string, string>();
  }

  const relationships = new Map<string, string>();
  for (const relationshipNode of getLocalDescendants(relsDocument, "Relationship")) {
    const type = relationshipNode.getAttribute("Type");
    const target = relationshipNode.getAttribute("Target");
    if (!type || !target) {
      continue;
    }
    relationships.set(type, resolveRelationshipPath(relsPath, target));
  }

  return relationships;
}

export function readChartColorPalette(
  archive: Record<string, Uint8Array>,
  colorStylePath: string | null | undefined,
  themePalette?: XlsxThemePalette | null
) {
  const colorStyleXml = readArchiveText(archive, colorStylePath);
  if (!colorStyleXml) {
    return [];
  }

  const colorStyleDocument = parseXml(colorStyleXml);
  if (!colorStyleDocument?.documentElement) {
    return [];
  }

  return Array.from(colorStyleDocument.documentElement.childNodes)
    .filter((child): child is Element => isElementNode(child) && child.localName !== "variation")
    .map((child) => (
      resolveChartColorNode(child, themePalette)
      ?? resolveChartColorNode(findFirstChartColorElement(child), themePalette)
    ))
    .filter((color): color is string => typeof color === "string");
}

export function readChartStyleAppearance(
  archive: Record<string, Uint8Array>,
  stylePath: string | null | undefined,
  themePalette?: XlsxThemePalette | null
): ChartStyleAppearance {
  const styleXml = readArchiveText(archive, stylePath);
  if (!styleXml) {
    return {};
  }

  const styleDocument = parseXml(styleXml);
  if (!styleDocument) {
    return {};
  }

  const dataPointNode = getFirstLocalDescendant(styleDocument, "dataPoint");
  const fillRefNode = dataPointNode ? getFirstLocalChild(dataPointNode, "fillRef") : null;
  const index = Number(fillRefNode?.getAttribute("idx") ?? Number.NaN);
  const chartAreaNode = getFirstLocalDescendant(styleDocument, "chartArea");
  const chartAreaShapeProperties = chartAreaNode ? getFirstLocalChild(chartAreaNode, "spPr") : null;
  const chartAreaFontRef = chartAreaNode ? getFirstLocalChild(chartAreaNode, "fontRef") : null;
  const chartAreaFontColor = chartAreaFontRef
    ? resolveChartColorNode(Array.from(chartAreaFontRef.childNodes).find(isElementNode) ?? null, themePalette)
    : null;
  const titleNode = getFirstLocalDescendant(styleDocument, "title");
  const titleFontRef = titleNode ? getFirstLocalChild(titleNode, "fontRef") : null;
  const titleColor = titleFontRef
    ? resolveChartColorNode(Array.from(titleFontRef.childNodes).find(isElementNode) ?? null, themePalette)
    : null;
  const axisStyleNode = getFirstLocalDescendant(styleDocument, "categoryAxis")
    ?? getFirstLocalDescendant(styleDocument, "valueAxis");
  const axisShapeProperties = axisStyleNode ? getFirstLocalChild(axisStyleNode, "spPr") : null;
  const axisFontRef = axisStyleNode ? getFirstLocalChild(axisStyleNode, "fontRef") : null;
  const chartAreaNoFill = chartAreaShapeProperties ? getFirstLocalChild(chartAreaShapeProperties, "noFill") != null : false;

  return {
    axisLabelColor: axisFontRef
      ? resolveChartColorNode(Array.from(axisFontRef.childNodes).find(isElementNode) ?? null, themePalette) ?? undefined
      : undefined,
    axisLineColor: resolveChartLineStyle(axisShapeProperties, themePalette).color ?? undefined,
    chartAreaBorderColor: resolveChartLineStyle(chartAreaShapeProperties, themePalette).color ?? undefined,
    chartAreaFillColor: resolveChartFillColor(chartAreaShapeProperties, themePalette) ?? undefined,
    chartAreaNoFill,
    paletteOffset: Number.isFinite(index) ? index : undefined,
    textColor: chartAreaFontColor ?? undefined,
    titleColor: titleColor ?? chartAreaFontColor ?? undefined
  };
}

export function buildThemeSeriesPalette(themePalette?: XlsxThemePalette | null) {
  const themeColors = [4, 5, 6, 7, 8, 9]
    .map((index) => themePalette?.colorsByIndex[index] ?? null)
    .filter((color): color is string => Boolean(color));
  return themeColors.length > 0 ? themeColors : SERIES_COLORS;
}

export function normalizeBuiltinSurfaceStyleId(styleId: number | undefined) {
  if (typeof styleId !== "number" || !Number.isFinite(styleId)) {
    return null;
  }
  return styleId >= 100 ? styleId - 100 : styleId;
}

export function getBuiltinSurfacePalette(styleId: number | undefined, wireframe: boolean | undefined) {
  const normalized = normalizeBuiltinSurfaceStyleId(styleId);
  if (normalized === 34 || (wireframe === true && normalized == null)) {
    return ["#5b9bd5", "#ed7d31", "#a5a5a5"];
  }
  if (normalized === 35 || normalized === 36 || (wireframe !== true && normalized == null)) {
    return ["#2f5597", "#4472c4", "#5b9bd5", "#8faadc", "#d9e2f3"];
  }
  return null;
}

export function applyBuiltinSurfaceDefaults(chart: XlsxChart) {
  if (chart.chartType !== "Surface") {
    return;
  }

  const builtinPalette = getBuiltinSurfacePalette(chart.chartStyleId, chart.wireframe);
  if ((!chart.chartColorPalette || chart.chartColorPalette.length === 0) && builtinPalette) {
    chart.chartColorPalette = builtinPalette;
  }

  const wallFill = chart.wireframe ? "#d0d0d0" : "#d9d9df";
  const wallLine = chart.wireframe ? "#a6a6a6" : "#a8adb7";
  chart.floor = {
    ...(chart.floor ?? {}),
    fillColor: chart.floor?.fillColor ?? wallFill,
    lineColor: chart.floor?.lineColor ?? wallLine
  };
  chart.sideWall = {
    ...(chart.sideWall ?? {}),
    fillColor: chart.sideWall?.fillColor ?? wallFill,
    lineColor: chart.sideWall?.lineColor ?? wallLine
  };
  chart.backWall = {
    ...(chart.backWall ?? {}),
    fillColor: chart.backWall?.fillColor ?? wallFill,
    lineColor: chart.backWall?.lineColor ?? wallLine
  };
  if (!chart.surfaceMaterial && chart.wireframe !== true) {
    chart.surfaceMaterial = "flat";
  }
}

export function applyBuiltinChartDefaults(chart: XlsxChart, themePalette?: XlsxThemePalette | null) {
  const darkBuiltInStyle = typeof chart.chartStyleId === "number" && chart.chartStyleId >= 140 && chart.chartStyleId < 150;
  const textColor = themePalette?.colorsByIndex[1] ?? themePalette?.colorsByIndex[3] ?? null;
  const minorTypeface = themePalette?.minorLatinFont?.trim() || undefined;
  const derivedAxisColor = textColor ? applyLightnessTransform(textColor, 0.35, 0.55) : null;
  const derivedBorderColor = textColor
    ? applyLightnessTransform(textColor, chart.is3d ? 0.28 : 0.22, chart.is3d ? 0.6 : 0.7)
    : null;
  if (darkBuiltInStyle) {
    chart.chartAreaFillColor = chart.chartAreaFillColor ?? "#1f1f1f";
    chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? "#1f1f1f";
    chart.textColor = chart.textColor ?? "#f5f5f5";
    chart.titleColor = chart.titleColor ?? "#f5f5f5";
    chart.axisLabelColor = chart.axisLabelColor ?? "#d9d9d9";
    chart.axisLineColor = chart.axisLineColor ?? "#8c8c8c";
  }
  chart.chartAreaBorderColor = chart.chartAreaBorderColor ?? derivedBorderColor ?? undefined;
  chart.textColor = chart.textColor ?? textColor ?? undefined;
  chart.titleColor = chart.titleColor ?? textColor ?? undefined;
  chart.axisLabelColor = chart.axisLabelColor ?? derivedAxisColor ?? textColor ?? undefined;
  chart.axisLineColor = chart.axisLineColor ?? derivedAxisColor ?? textColor ?? undefined;
  chart.fontFamily = chart.fontFamily ?? minorTypeface;
  chart.titleFontFamily = chart.titleFontFamily ?? chart.fontFamily ?? minorTypeface;

  const seriesPalette = chart.chartColorPalette && chart.chartColorPalette.length > 0
    ? chart.chartColorPalette
    : buildThemeSeriesPalette(themePalette);
  if (!chart.chartColorPalette || chart.chartColorPalette.length === 0) {
    chart.chartColorPalette = seriesPalette;
  }

  chart.series = chart.series.map((series, index) => {
    const fallbackColor = seriesPalette[index % seriesPalette.length];
    return {
      ...series,
      color: series.color ?? series.lineColor ?? fallbackColor,
      lineColor: series.lineColor ?? series.color ?? fallbackColor,
      markerColor: series.markerColor ?? series.color ?? series.lineColor ?? fallbackColor,
      markerLineColor: series.markerLineColor ?? series.lineColor ?? series.color ?? fallbackColor
    };
  });
  chart.typeGroups = chart.typeGroups?.map((group, groupIndex) => ({
    ...group,
    series: group.series.map((series, seriesIndex) => {
      const fallbackColor = seriesPalette[(groupIndex + seriesIndex) % seriesPalette.length];
      return {
        ...series,
        color: series.color ?? series.lineColor ?? fallbackColor,
        lineColor: series.lineColor ?? series.color ?? fallbackColor,
        markerColor: series.markerColor ?? series.color ?? series.lineColor ?? fallbackColor,
        markerLineColor: series.markerLineColor ?? series.lineColor ?? series.color ?? fallbackColor
      };
    })
  }));
  applyBuiltinSurfaceDefaults(chart);
}

export function parseChartPointStyles(seriesNode: Element, themePalette?: XlsxThemePalette | null): XlsxChartPointStyle[] {
  const pointStyles: XlsxChartPointStyle[] = [];

  for (const dataPointNode of getLocalChildren(seriesNode, "dPt")) {
    const indexValue = readChartNumericAttribute(dataPointNode, "idx");
    if (indexValue === undefined) {
      continue;
    }
    const shapeProperties = getFirstLocalChild(dataPointNode, "spPr");
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    pointStyles.push({
      color: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
      explosion: readChartNumericAttribute(dataPointNode, "explosion"),
      index: indexValue,
      lineColor: lineStyle.color ?? undefined
    });
  }

  return pointStyles;
}

export function parseInvertNegativeStyle(seriesNode: Element, themePalette?: XlsxThemePalette | null) {
  const invertNode = getFirstLocalDescendant(seriesNode, "invertSolidFillFmt");
  const shapeProperties = invertNode ? getFirstLocalChild(invertNode, "spPr") : null;
  if (!shapeProperties) {
    return {
      color: undefined,
      lineColor: undefined
    };
  }

  const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
  return {
    color: resolveChartFillColor(shapeProperties, themePalette) ?? undefined,
    lineColor: lineStyle.color ?? undefined
  };
}

export function applyChartSeriesStyleFromXml(chart: XlsxChart, chartTypeNode: Element, themePalette?: XlsxThemePalette | null) {
  const seriesNodes = getLocalChildren(chartTypeNode, "ser");
  chart.series = chart.series.map((series, index) => {
    const seriesNode = seriesNodes[index];
    if (!seriesNode) {
      return series;
    }

    const shapeProperties = getFirstLocalChild(seriesNode, "spPr");
    const markerNode = getFirstLocalChild(seriesNode, "marker");
    const markerShapeProperties = getFirstLocalChild(markerNode ?? chartTypeNode, "spPr");
    const lineStyle = resolveChartLineStyle(shapeProperties, themePalette);
    const markerLineStyle = resolveChartLineStyle(markerShapeProperties, themePalette);
    const fillColor = resolveChartFillColor(shapeProperties, themePalette);
    const markerSize = readChartNumericAttribute(markerNode, "size");
    const markerSymbolNode = markerNode ? getFirstLocalChild(markerNode, "symbol") : null;
    const markerSymbol = markerSymbolNode?.getAttribute("val") ?? undefined;
    const pointStyles = parseChartPointStyles(seriesNode, themePalette);
    const seriesExplosion = readChartNumericAttribute(seriesNode, "explosion");
    const invertNegativeStyle = parseInvertNegativeStyle(seriesNode, themePalette);
    const invertIfNegative = readChartBooleanAttribute(seriesNode, "invertIfNegative");
    const isScatterChart = (
      chart.chartType === "Scatter"
      || chart.chartType === "ScatterLines"
      || chart.chartType === "ScatterSmooth"
      || chart.chartType === "Bubble"
    );
    const cachedCategories = isScatterChart
      ? (
          parseChartCacheValues(getFirstLocalChild(seriesNode, "xVal"), "numCache", "value")
          ?? parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, "xVal"), "category")
        )
      : (
          parseChartCacheValues(getFirstLocalChild(seriesNode, "cat"), "strCache", "category")
          ?? parseChartCacheValues(getFirstLocalChild(seriesNode, "cat"), "numCache", "category")
          ?? parseChartMultiLevelCacheValues(getFirstLocalChild(seriesNode, "cat"), "category")
        );
    const cachedValues = isScatterChart
      ? parseChartCacheValues(getFirstLocalChild(seriesNode, "yVal"), "numCache", "value")
      : parseChartCacheValues(getFirstLocalChild(seriesNode, "val"), "numCache", "value");
    const cachedBubbleSizes = chart.chartType === "Bubble"
      ? parseChartCacheValues(getFirstLocalChild(seriesNode, "bubbleSize"), "numCache", "value")
      : null;
    const existingShapeProperties = series.shapeProperties && typeof series.shapeProperties === "object"
      ? series.shapeProperties as Record<string, unknown>
      : null;
    const rawFillColor = typeof existingShapeProperties?.solidFillHex === "string"
      ? normalizeHexColor(existingShapeProperties.solidFillHex)
      : null;
    const rawLineColor = typeof existingShapeProperties?.lineColorHex === "string"
      ? normalizeHexColor(existingShapeProperties.lineColorHex)
      : null;
    const resolvedLineColor = lineStyle.hidden
      ? undefined
      : rawLineColor ?? lineStyle.color ?? rawFillColor ?? fillColor ?? series.lineColor ?? series.color;

    const hasCategoryReference = typeof series.categoriesRef?.formula === "string" && series.categoriesRef.formula.length > 0;
    const hasValueReference = typeof series.valuesRef?.formula === "string" && series.valuesRef.formula.length > 0;
    const hasBubbleSizeReference = typeof series.bubbleSizeRef?.formula === "string" && series.bubbleSizeRef.formula.length > 0;

    return {
      ...series,
      bubbleSizes: !hasBubbleSizeReference && cachedBubbleSizes
        ? cachedBubbleSizes.map((value) => (typeof value === "number" && Number.isFinite(value) ? value : null))
        : series.bubbleSizes,
      categories: !hasCategoryReference && cachedCategories ? cachedCategories : series.categories,
      color: rawFillColor ?? rawLineColor ?? fillColor ?? lineStyle.color ?? series.color,
      dataPointStyles: pointStyles.length > 0 ? pointStyles : series.dataPointStyles,
      lineColor: resolvedLineColor,
      lineWidthPx: lineStyle.hidden ? undefined : lineStyle.widthPx ?? series.lineWidthPx,
      markerColor: rawFillColor
        ?? rawLineColor
        ?? resolveChartFillColor(markerShapeProperties, themePalette)
        ?? fillColor
        ?? lineStyle.color
        ?? undefined,
      markerLineColor: rawLineColor
        ?? rawFillColor
        ?? markerLineStyle.color
        ?? lineStyle.color
        ?? fillColor
        ?? undefined,
      markerSize: markerSize ?? series.markerSize,
      markerSymbol,
      smooth: readChartBooleanAttribute(seriesNode, "smooth") ?? series.smooth,
      invertIfNegative: invertIfNegative ?? series.invertIfNegative,
      shapeProperties: {
        ...series.shapeProperties,
        xmlExplosion: seriesExplosion ?? undefined,
        xmlFillColor: fillColor ?? undefined,
        xmlLineHidden: lineStyle.hidden ? true : undefined,
        xmlLineColor: lineStyle.color ?? undefined,
        xmlLineWidthPx: lineStyle.widthPx ?? undefined,
        xmlNegativeFillColor: invertNegativeStyle.color ?? undefined,
        xmlNegativeLineColor: invertNegativeStyle.lineColor ?? undefined
      },
      negativeColor: invertNegativeStyle.color ?? series.negativeColor,
      negativeLineColor: invertNegativeStyle.lineColor ?? series.negativeLineColor,
      values: !hasValueReference && cachedValues
        ? cachedValues.map((value) => (typeof value === "number" && Number.isFinite(value) ? value : null))
        : series.values
    };
  });
}

