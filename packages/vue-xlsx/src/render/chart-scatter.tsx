import { scaleBand, scaleLinear, scalePoint } from "d3-scale";
import { symbol as d3Symbol, symbolCircle, symbolCross, symbolDiamond, symbolSquare, symbolStar, symbolTriangle, curveCatmullRom, curveLinear, line as d3Line } from "d3-shape";
import type { XlsxChart, XlsxChartElementSelection, XlsxChartSeries, XlsxImageRect } from "@extend-ai/xlsx-core";
import { chartSeriesColor, chartSeriesStrokeColor, chartPointColor, safeNumber, clamp, darkenColor, lightenColor, coerceLooseNumber, resolveRenderableSeriesValue, resolveChartTextColor, resolveChartAxisTextColor, resolveChartMutedTextColor, buildChartFontFamily, formatTickValue, formatPercentTickValue, chartElementDataProps, barChartElementDataProps, resolveChartElementTarget, resolveChartSelectionFromTarget, isSelectedChartSeries, isSelectedChartPoint, isSelectedChartPointOrSeries, renderSelectionRectHandles, renderSelectionPointHandles, normalizeRendererHexColor, toSvgNumber, normalizeChartMarkerSymbol, markerSymbolPath, normalizeLegendPosition, parseRgbColor, mixRgbColor, escapeCssFontFamilyToken, DEFAULT_CHART_FONT_STACK, DEFAULT_CHART_TEXT_COLOR, DEFAULT_CHART_MUTED_TEXT_COLOR, truncateSvgText, getCategoryLabels, normalizeCategoryLabel, estimateReferencePointCount, isHistogramLikeSeries, isHistogramLikeChart, resolve3dFrameOffsets, resolveAxisDomainWithChartOverrides, resolveCategoryBandPadding, chartSeriesBarColors, buildNumericTickValues, buildNiceStep, resolveNumericAxisDomain, renderCartesianAxes, renderExtrudedRect, buildChartStages, resolveChartStageSubtotal, computeBoxWhiskerStats, resolveBoxWhiskerVisibility, resolveBoxWhiskerQuartileMethod, computePercentile, renderLineOrAreaChart3d, projectCartesian3dPoint, scaleProjectedVector, buildLinearSvgPath, buildRibbonSvgPath, renderRadialFrustum, sampleProjectedEllipseArc, isComboChart, getComboLegendSeries, findAxisForGroup, buildComboGroups, resolveStockRoleIndices, resolveStockPalette, normalizeBuiltinPieStyleId, getBuiltinPiePalette, resolvePiePointColor, selectPrimaryPieSeriesIndex, buildPieEntries, buildHierarchyData, resolveHierarchyNodeColor, resolveTreemapNodeColor, excelTreemapTile, normalizePieArc, resolvePieFrontSegments, pieEllipsePoint, buildPieOuterWallPath, isPieFrontFacingAngle, buildPieRadialWallPath, TWO_PI, normalizeRegionMapKey, resolveRegionMapFeature, resolveRegionMapBaseColor, resolveRegionMapDataColor, resolveRegionMapValueColors, resolveRegionMapColorStrings, resolveRegionMapLayoutProperties, resolveRegionMapFeatureSet, getRegionMapBaseFeatures, resolveRegionMapValueColorFromStops, resolveRegionMapValueColor, resolveRegionMapNoDataColor, resolveSurfaceBaseColor, normalizeBuiltinSurfaceStyleId, hasExplicitSurfaceBaseColor, buildMonochromeSurfacePalette, getBuiltinSurfacePalette, shouldPreferBuiltinSurfacePalette, getSurfaceBandCount, getSurfaceColorStops, resolveSurfaceBandPaletteColor, getSurfaceDomain, resolveSurfaceColor, resolveSurfaceBandColor, resolveSurfaceBandIndex, buildSurfaceSmoothPath, splitSurfacePointRuns, connectSurfaceContourSegments, getSurfaceWireframePalette, resolveSurfaceWireframeColor, resolveSurfacePlotRect, buildSurfaceLegendItems, isContourSurfaceChart, getLegendItems, buildLayout, normalizeRenderableChartType, renderSurfaceAxes, renderTitle, renderLegend, buildRegionMapLegendItems, buildStockLegendItems, excelSerialToDate, isLikelyDateFormatCode, formatExcelDateSerial, formatCategoryLabel } from "./chart-renderer";
import type { ChartRendererPalette, ChartSvgProps, ChartLayout, LegendItem, BarRect, PlotRect, ChartStage, ChartHierarchyDatum, ComboRenderableGroup, BoxWhiskerStats, SurfaceDomain, RegionMapFeature, SurfacePathPoint, SurfaceContourSegment, ChartElementDataOptions } from "./chart-renderer";


export function renderScatterChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, smooth: boolean, selectedChartElement?: XlsxChartElementSelection | null) {
  const plot = layout.plot;
  const rawRecord = chart.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const normalizedTitle = typeof chart.title === "string" ? chart.title.trim().toLowerCase() : "";
  const scatterStyle = typeof chart.scatterStyle === "string"
    ? chart.scatterStyle
    : rawRecord && typeof rawRecord.scatterStyle === "string"
      ? rawRecord.scatterStyle
      : undefined;
  const styleDrawsLine = scatterStyle
    ? scatterStyle === "line" || scatterStyle === "lineMarker" || scatterStyle === "smooth" || scatterStyle === "smoothMarker"
    : true;
  const styleShowsMarkers = scatterStyle
    ? scatterStyle === "marker" || scatterStyle === "lineMarker" || scatterStyle === "smoothMarker"
    : true;
  const styleUsesSmoothCurve = scatterStyle
    ? scatterStyle === "smooth" || scatterStyle === "smoothMarker"
    : smooth;
  const titleForcesMarkerOnly = normalizedTitle.includes("marker only");
  const titleForcesNoMarkers = normalizedTitle.includes("no markers");
  const titleForcesWithMarkers = normalizedTitle.includes("with markers");
  const pointsBySeries = chart.series.map((series) => {
    const yLength = series.values.length;
    const directXValues = Array.from({ length: yLength }, (_, index) => coerceLooseNumber(series.categories[index]));

    let resolvedXValues = directXValues;
    let usesSyntheticIndex = false;
    if (directXValues.filter((value) => value != null).length < Math.max(2, yLength - 1)) {
      const categories = series.categories ?? [];
      const pairedXValues = Array.from({ length: yLength }, (_, index) => {
        const candidates = [
          categories[index],
          categories[index * 2],
          categories[index * 2 + 1],
          categories[index + 1]
        ];
        for (const candidate of candidates) {
          const parsed = coerceLooseNumber(candidate);
          if (parsed != null) {
            return parsed;
          }
        }
        return null;
      });
      if (pairedXValues.filter((value) => value != null).length >= 2) {
        resolvedXValues = pairedXValues;
      } else {
        const numericPool = categories
          .map((entry) => coerceLooseNumber(entry))
          .filter((value): value is number => value != null);
        if (numericPool.length >= 2) {
          resolvedXValues = Array.from(
            { length: yLength },
            (_, index) => numericPool[index] ?? numericPool[numericPool.length - 1] ?? null
          );
        } else {
          usesSyntheticIndex = true;
          resolvedXValues = Array.from({ length: yLength }, (_, index) => index + 1);
        }
      }
    }

    const points = series.values.map((value, index) => {
      const x = coerceLooseNumber(resolvedXValues[index]);
      const y = safeNumber(value);
      return x == null || y == null ? null : { x, y };
    }).filter((point): point is { x: number; y: number } => point != null);
    return {
      pointCount: yLength,
      points,
      series,
      usesSyntheticIndex
    };
  });

  const allX = pointsBySeries.flatMap((series) => series.points.map((point) => point.x));
  const allY = pointsBySeries.flatMap((series) => series.points.map((point) => point.y));
  if (allX.length === 0 || allY.length === 0) {
    return null;
  }

  const hasExplicitMinX = typeof chart.categoryAxis?.min === "number" && Number.isFinite(chart.categoryAxis.min);
  const hasExplicitMaxX = typeof chart.categoryAxis?.max === "number" && Number.isFinite(chart.categoryAxis.max);
  const hasExplicitMinY = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min);
  const hasExplicitMaxY = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max);
  const hasSyntheticIndexAxis = pointsBySeries.some((series) => series.usesSyntheticIndex);
  const syntheticPointCount = Math.max(0, ...pointsBySeries.map((series) => series.pointCount));

  let minX = hasExplicitMinX ? Number(chart.categoryAxis?.min) : Math.min(...allX);
  let maxX = hasExplicitMaxX ? Number(chart.categoryAxis?.max) : Math.max(...allX);
  let minY = hasExplicitMinY ? Number(chart.valueAxis?.min) : Math.min(...allY);
  let maxY = hasExplicitMaxY ? Number(chart.valueAxis?.max) : Math.max(...allY);

  if (!hasExplicitMinX) {
    minX = hasSyntheticIndexAxis
      ? 0
      : chart.categoryAxis?.crosses === "autoZero"
        ? Math.min(0, minX)
        : minX;
  }
  if (!hasExplicitMaxX) {
    maxX = hasSyntheticIndexAxis
      ? Math.max(maxX, syntheticPointCount + 1)
      : chart.categoryAxis?.crosses === "autoZero"
        ? Math.max(0, maxX)
        : maxX;
  }
  if (!hasExplicitMinY && chart.valueAxis?.crosses === "autoZero") {
    minY = Math.min(0, minY);
  }
  if (!hasExplicitMaxY && chart.valueAxis?.crosses === "autoZero") {
    maxY = Math.max(0, maxY);
  }

  if (maxX <= minX) {
    maxX = minX + 1;
  }
  if (maxY <= minY) {
    maxY = minY + 1;
  }

  const xStep = typeof chart.categoryAxis?.majorUnit === "number" && chart.categoryAxis.majorUnit > 0
    ? chart.categoryAxis.majorUnit
    : buildNiceStep(minX, maxX, 5);
  const yStep = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : buildNiceStep(minY, maxY, 5);

  if (!hasExplicitMinX && !hasSyntheticIndexAxis && Number.isFinite(xStep) && xStep > 0) {
    minX = Math.floor(minX / xStep) * xStep;
  }
  if (!hasExplicitMaxX && !hasSyntheticIndexAxis && Number.isFinite(xStep) && xStep > 0) {
    maxX = Math.ceil(maxX / xStep) * xStep;
  }
  if (!hasExplicitMinY && Number.isFinite(yStep) && yStep > 0) {
    minY = Math.floor(minY / yStep) * yStep;
  }
  if (!hasExplicitMaxY && Number.isFinite(yStep) && yStep > 0) {
    maxY = Math.ceil(maxY / yStep) * yStep;
  }

  const safeMaxX = maxX <= minX ? minX + 1 : maxX;
  const safeMaxY = maxY <= minY ? minY + 1 : maxY;

  const xScale = scaleLinear().domain([minX, safeMaxX]).range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear().domain([minY, safeMaxY]).range([plot.top + plot.height, plot.top]);
  const xMajorUnit = typeof chart.categoryAxis?.majorUnit === "number" && chart.categoryAxis.majorUnit > 0
    ? chart.categoryAxis.majorUnit
    : hasSyntheticIndexAxis
      ? 1
      : undefined;
  const yMajorUnit = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : buildNiceStep(minY, safeMaxY, 6);
  const xTicks = buildNumericTickValues(minX, safeMaxX, xMajorUnit);
  const yTicks = buildNumericTickValues(minY, safeMaxY, yMajorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);

  const hasZeroX = minX <= 0 && safeMaxX >= 0;
  const hasZeroY = minY <= 0 && safeMaxY >= 0;

  return (
    <g>
      {xTicks.map((tick) => (
        <g key={`scatter-x-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.7)}
            stroke-width={1}
            x1={xScale(tick)}
            x2={xScale(tick)}
            y1={plot.top}
            y2={plot.top + plot.height}
          />
          <text fill={labelColor} font-size={10} text-anchor="middle" x={xScale(tick)} y={plot.top + plot.height + 14}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      {yTicks.map((tick) => (
        <g key={`scatter-y-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.7)}
            stroke-width={1}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={yScale(tick)}
            y2={yScale(tick)}
          />
          <text fill={labelColor} font-size={10} text-anchor="end" x={plot.left - 6} y={yScale(tick) + 3}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      {hasZeroY ? (
        <line
          stroke={axisColor}
          stroke-width={1.2}
          x1={plot.left}
          x2={plot.left + plot.width}
          y1={yScale(0)}
          y2={yScale(0)}
        />
      ) : null}
      {hasZeroX ? (
        <line
          stroke={axisColor}
          stroke-width={1.2}
          x1={xScale(0)}
          x2={xScale(0)}
          y1={plot.top}
          y2={plot.top + plot.height}
        />
      ) : null}
      {pointsBySeries.map((seriesPoints, seriesIndex) => {
        const series = seriesPoints.series;
        const markerSize = Math.max(5, series.markerSize ?? 7);
        const markerPath = markerSymbolPath(
          normalizeChartMarkerSymbol(series.markerSymbol),
          markerSize * 0.55
        );
        const markerFill = series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesStrokeColor(chart, seriesIndex);
        const shouldDrawLine = !titleForcesMarkerOnly
          && styleDrawsLine
          && series.shapeProperties?.xmlLineHidden !== true
          && seriesPoints.points.length > 1;
        const shouldDrawMarkers = !titleForcesNoMarkers
          && (titleForcesMarkerOnly || titleForcesWithMarkers || styleShowsMarkers)
          && markerPath.length > 0;
        const lineCurve = smooth || styleUsesSmoothCurve || series.smooth === true
          ? curveCatmullRom.alpha(0.5)
          : curveLinear;
        const linePath = shouldDrawLine
          ? d3Line<{ x: number; y: number }>()
              .x((point) => xScale(point.x))
              .y((point) => yScale(point.y))
              .curve(lineCurve)(seriesPoints.points) ?? ""
          : "";

        return (
          <g key={`scatter-series-${seriesIndex}`}>
            {shouldDrawLine && linePath.length > 0 ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={linePath}
                fill="none"
                stroke={series.lineColor ?? chartSeriesStrokeColor(chart, seriesIndex)}
                stroke-linecap="round"
                stroke-linejoin="round"
                stroke-width={Math.max(1.2, series.lineWidthPx ?? 2)}
              />
            ) : null}
            {seriesPoints.points.map((point, pointIndex) => {
              if (!shouldDrawMarkers) {
                return null;
              }
              return (
                <g
                  key={`scatter-point-${seriesIndex}-${pointIndex}`}
                  transform={`translate(${xScale(point.x)}, ${yScale(point.y)})`}
                >
                  <path
                    {...chartElementDataProps(seriesIndex, pointIndex, { selectionMode: "seriesFirst" })}
                    d={markerPath}
                    fill={markerFill}
                    stroke="none"
                    stroke-width={0}
                  />
                </g>
              );
            })}
            {seriesPoints.points.map((point, pointIndex) => (
              isSelectedChartPoint(selectedChartElement, chart.id, seriesIndex, pointIndex)
              || (
                isSelectedChartSeries(selectedChartElement, chart.id, seriesIndex)
                && selectedChartElement?.kind !== "point"
              )
                ? renderSelectionPointHandles(
                    `scatter-selection-${seriesIndex}-${pointIndex}`,
                    [{ x: xScale(point.x), y: yScale(point.y) }]
                  )
                : null
            ))}
          </g>
        );
      })}
    </g>
  );
}

export function renderBubbleChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const plot = layout.plot;
  const pointsBySeries = chart.series.map((series) => (
    series.values.map((value, index) => {
      const x = safeNumber(series.categories[index]);
      const y = safeNumber(value);
      const bubble = safeNumber(series.bubbleSizes?.[index]);
      return x == null || y == null ? null : { bubble: bubble ?? 1, index, x, y };
    }).filter((point): point is { bubble: number; index: number; x: number; y: number } => point != null)
  ));

  const allX = pointsBySeries.flatMap((points) => points.map((point) => point.x));
  const allY = pointsBySeries.flatMap((points) => points.map((point) => point.y));
  const allBubble = pointsBySeries.flatMap((points) => points.map((point) => point.bubble));
  if (allX.length === 0 || allY.length === 0) {
    return null;
  }
  const minX = Math.min(...allX);
  const maxX = Math.max(...allX);
  const minY = Math.min(...allY);
  const maxY = Math.max(...allY);
  const bubbleMagnitudes = allBubble.map((value) => Math.sqrt(Math.max(0, value)));
  const minBubbleMagnitude = Math.min(...bubbleMagnitudes);
  const maxBubbleMagnitude = Math.max(...bubbleMagnitudes);

  const bubbleScaleFactor = clamp((chart.bubbleScale ?? 100) / 100, 0.2, 4);
  const minRadius = 4;
  const maxRadius = Math.max(minRadius + 2, 12 * bubbleScaleFactor);
  const radiusScale = scaleLinear()
    .domain([minBubbleMagnitude, maxBubbleMagnitude <= minBubbleMagnitude ? minBubbleMagnitude + 1 : maxBubbleMagnitude])
    .range([minRadius, maxRadius]);
  const safeMaxX = maxX <= minX ? minX + 1 : maxX;
  const safeMaxY = maxY <= minY ? minY + 1 : maxY;
  const xSpan = safeMaxX - minX;
  const ySpan = safeMaxY - minY;
  const xPad = Math.max(xSpan * 0.04, (xSpan * maxRadius) / Math.max(1, plot.width));
  const yPad = Math.max(ySpan * 0.06, (ySpan * maxRadius) / Math.max(1, plot.height));
  const xDomain = resolveAxisDomainWithChartOverrides(
    chart.categoryAxis,
    minX - xPad,
    safeMaxX + xPad,
    chart.categoryAxis?.crosses === "autoZero"
  );
  const yDomain = resolveAxisDomainWithChartOverrides(
    chart.valueAxis,
    minY - yPad,
    safeMaxY + yPad,
    chart.valueAxis?.crosses === "autoZero"
  );
  const xScale = scaleLinear().domain([xDomain.min, xDomain.max]).range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear().domain([yDomain.min, yDomain.max]).range([plot.top + plot.height, plot.top]);

  const xTicks = xDomain.ticks;
  const yTicks = yDomain.ticks;
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
  const isBubble3d = chart.bubble3d === true || chart.is3d === true;
  const labelsEnabled = Boolean(
    chart.dataLabels?.showCategoryName
    || chart.dataLabels?.showSeriesName
    || chart.dataLabels?.showValue
    || chart.dataLabels?.showBubbleSize
    || chart.dataLabels?.showPercent
  );
  const bubbleTotals = pointsBySeries.map((points) => (
    points.reduce((sum, point) => sum + Math.abs(point.bubble), 0)
  ));

  return (
    <g>
      {xTicks.map((tick) => (
        <g key={`bubble-x-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.72)}
            stroke-width={1}
            x1={xScale(tick)}
            x2={xScale(tick)}
            y1={plot.top}
            y2={plot.top + plot.height}
          />
          <text fill={labelColor} font-size={10} text-anchor="middle" x={xScale(tick)} y={plot.top + plot.height + 14}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      {yTicks.map((tick) => (
        <g key={`bubble-y-${tick}`}>
          <line
            stroke={lightenColor(axisColor, 0.72)}
            stroke-width={1}
            x1={plot.left}
            x2={plot.left + plot.width}
            y1={yScale(tick)}
            y2={yScale(tick)}
          />
          <text fill={labelColor} font-size={10} text-anchor="end" x={plot.left - 6} y={yScale(tick) + 3}>
            {formatTickValue(tick)}
          </text>
        </g>
      ))}
      <line stroke={axisColor} stroke-width={1.2} x1={plot.left} x2={plot.left + plot.width} y1={plot.top + plot.height} y2={plot.top + plot.height} />
      <line stroke={axisColor} stroke-width={1.2} x1={plot.left} x2={plot.left} y1={plot.top} y2={plot.top + plot.height} />
      {pointsBySeries.map((points, seriesIndex) => (
        <g key={`bubble-series-${seriesIndex}`}>
          {isBubble3d ? (
            <defs>
              <radialGradient id={`bubble3d-grad-${chart.id}-${seriesIndex}`} cx="35%" cy="30%" r="70%">
                <stop offset="0%" stop-color={lightenColor(chart.series[seriesIndex]?.color ?? chart.series[seriesIndex]?.lineColor ?? chartSeriesColor(chart, seriesIndex), 0.42)} />
                <stop offset="58%" stop-color={chart.series[seriesIndex]?.color ?? chart.series[seriesIndex]?.lineColor ?? chartSeriesColor(chart, seriesIndex)} />
                <stop offset="100%" stop-color={darkenColor(chart.series[seriesIndex]?.color ?? chart.series[seriesIndex]?.lineColor ?? chartSeriesColor(chart, seriesIndex), 0.18)} />
              </radialGradient>
            </defs>
          ) : null}
          {[...points]
            .sort((left, right) => {
              if (left.bubble !== right.bubble) {
                return right.bubble - left.bubble;
              }
              return left.index - right.index;
            })
            .map((point) => {
            const series = chart.series[seriesIndex];
            const baseColor = series?.color ?? series?.lineColor ?? chartSeriesColor(chart, seriesIndex);
            const radius = radiusScale(Math.sqrt(Math.max(0, point.bubble)));
            const pieces: string[] = [];
            if (chart.dataLabels?.showSeriesName && series?.name) {
              pieces.push(series.name);
            }
            if (chart.dataLabels?.showCategoryName) {
              pieces.push(formatTickValue(point.x));
            }
            if (chart.dataLabels?.showValue) {
              pieces.push(formatTickValue(point.y));
            }
            if (chart.dataLabels?.showBubbleSize) {
              pieces.push(formatTickValue(point.bubble));
            }
            if (chart.dataLabels?.showPercent) {
              pieces.push(`${Math.round((Math.abs(point.bubble) / Math.max(1, bubbleTotals[seriesIndex] ?? 1)) * 100)}%`);
            }
            return (
              <g key={`bubble-${seriesIndex}-${point.index}`}>
                <circle
                  {...chartElementDataProps(seriesIndex, point.index, { selectionMode: "seriesFirst" })}
                  cx={xScale(point.x)}
                  cy={yScale(point.y)}
                  fill={isBubble3d ? `url(#bubble3d-grad-${chart.id}-${seriesIndex})` : baseColor}
                  fill-opacity={isBubble3d ? 0.98 : 0.78}
                  r={radius}
                  stroke={darkenColor(baseColor, 0.18)}
                  stroke-width={isBubble3d ? 1.2 : 1}
                />
                {isSelectedChartPoint(selectedChartElement, chart.id, seriesIndex, point.index)
                || (
                  isSelectedChartSeries(selectedChartElement, chart.id, seriesIndex)
                  && selectedChartElement?.kind !== "point"
                )
                  ? renderSelectionRectHandles(
                      `bubble-selection-${seriesIndex}-${point.index}`,
                      xScale(point.x) - radius,
                      yScale(point.y) - radius,
                      radius * 2,
                      radius * 2
                    )
                  : null}
                {isBubble3d ? (
                  <ellipse
                    cx={xScale(point.x) - radius * 0.16}
                    cy={yScale(point.y) - radius * 0.22}
                    fill="#ffffff"
                    opacity={0.22}
                    rx={Math.max(1.5, radius * 0.34)}
                    ry={Math.max(1, radius * 0.2)}
                  />
                ) : null}
                {labelsEnabled && pieces.length > 0 ? (
                  <text
                    fill={resolveChartTextColor(chart)}
                    font-size={10}
                    text-anchor="start"
                    x={xScale(point.x) + radius + 4}
                    y={yScale(point.y) + 3}
                  >
                    {pieces.join(", ")}
                  </text>
                ) : null}
              </g>
            );
          })}
        </g>
      ))}
    </g>
  );
}
