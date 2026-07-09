import { scaleBand, scaleLinear, scalePoint } from "d3-scale";
import { area as d3Area, curveCatmullRom, curveLinear, curveLinearClosed, line as d3Line, symbol as d3Symbol, symbolCircle, symbolCross, symbolDiamond, symbolSquare, symbolStar, symbolTriangle } from "d3-shape";
import type { CurveFactory } from "d3-shape";
import type { XlsxChart, XlsxChartAxis, XlsxChartElementSelection, XlsxChartSeries, XlsxChartTypeGroup, XlsxImageRect } from "@extend-ai/xlsx-core";
import { chartSeriesColor, chartSeriesStrokeColor, chartPointColor, safeNumber, clamp, darkenColor, lightenColor, coerceLooseNumber, resolveRenderableSeriesValue, resolveChartTextColor, resolveChartAxisTextColor, resolveChartMutedTextColor, buildChartFontFamily, formatTickValue, formatPercentTickValue, chartElementDataProps, barChartElementDataProps, resolveChartElementTarget, resolveChartSelectionFromTarget, isSelectedChartSeries, isSelectedChartPoint, isSelectedChartPointOrSeries, renderSelectionRectHandles, renderSelectionPointHandles, normalizeRendererHexColor, toSvgNumber, normalizeChartMarkerSymbol, markerSymbolPath, normalizeLegendPosition, parseRgbColor, mixRgbColor, escapeCssFontFamilyToken, DEFAULT_CHART_FONT_STACK, DEFAULT_CHART_TEXT_COLOR, DEFAULT_CHART_MUTED_TEXT_COLOR, truncateSvgText, getCategoryLabels, normalizeCategoryLabel, estimateReferencePointCount, isHistogramLikeSeries, isHistogramLikeChart, resolve3dFrameOffsets, resolveAxisDomainWithChartOverrides, resolveCategoryBandPadding, chartSeriesBarColors, buildNumericTickValues, buildNiceStep, resolveNumericAxisDomain, renderCartesianAxes, renderExtrudedRect, buildChartStages, resolveChartStageSubtotal, computeBoxWhiskerStats, resolveBoxWhiskerVisibility, resolveBoxWhiskerQuartileMethod, computePercentile, renderLineOrAreaChart3d, projectCartesian3dPoint, scaleProjectedVector, buildLinearSvgPath, buildRibbonSvgPath, renderRadialFrustum, sampleProjectedEllipseArc, isComboChart, getComboLegendSeries, findAxisForGroup, buildComboGroups, resolveStockRoleIndices, resolveStockPalette, normalizeBuiltinPieStyleId, getBuiltinPiePalette, resolvePiePointColor, selectPrimaryPieSeriesIndex, buildPieEntries, buildHierarchyData, resolveHierarchyNodeColor, resolveTreemapNodeColor, excelTreemapTile, normalizePieArc, resolvePieFrontSegments, pieEllipsePoint, buildPieOuterWallPath, isPieFrontFacingAngle, buildPieRadialWallPath, TWO_PI, normalizeRegionMapKey, resolveRegionMapFeature, resolveRegionMapBaseColor, resolveRegionMapDataColor, resolveRegionMapValueColors, resolveRegionMapColorStrings, resolveRegionMapLayoutProperties, resolveRegionMapFeatureSet, getRegionMapBaseFeatures, resolveRegionMapValueColorFromStops, resolveRegionMapValueColor, resolveRegionMapNoDataColor, resolveSurfaceBaseColor, normalizeBuiltinSurfaceStyleId, hasExplicitSurfaceBaseColor, buildMonochromeSurfacePalette, getBuiltinSurfacePalette, shouldPreferBuiltinSurfacePalette, getSurfaceBandCount, getSurfaceColorStops, resolveSurfaceBandPaletteColor, getSurfaceDomain, resolveSurfaceColor, resolveSurfaceBandColor, resolveSurfaceBandIndex, buildSurfaceSmoothPath, splitSurfacePointRuns, connectSurfaceContourSegments, getSurfaceWireframePalette, resolveSurfaceWireframeColor, resolveSurfacePlotRect, buildSurfaceLegendItems, isContourSurfaceChart, getLegendItems, buildLayout, normalizeRenderableChartType, renderSurfaceAxes, renderTitle, renderLegend, buildRegionMapLegendItems, buildStockLegendItems, excelSerialToDate, isLikelyDateFormatCode, formatExcelDateSerial, formatCategoryLabel } from "./chart-renderer";
import type { ChartRendererPalette, ChartSvgProps, ChartLayout, LegendItem, BarRect, PlotRect, ChartStage, ChartHierarchyDatum, ComboRenderableGroup, BoxWhiskerStats, SurfaceDomain, RegionMapFeature, SurfacePathPoint, SurfaceContourSegment, ChartElementDataOptions } from "./chart-renderer";


export function renderLineOrAreaChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string, selectedChartElement?: XlsxChartElementSelection | null) {
  if (chart.series.length === 0) {
    return null;
  }
  const categories = getCategoryLabels(chart);
  if (categories.length === 0) {
    return null;
  }
  const plot = layout.plot;
  const displayBlanksAs = chart.displayBlanksAs;
  const isAreaChart = chartType === "Area" || chartType === "AreaStacked" || chartType === "AreaPercentStacked";
  const isStackedArea = chartType === "AreaStacked" || chartType === "AreaPercentStacked";
  const isStackedLine = chartType === "LineStacked" || chartType === "LinePercentStacked";
  const isStackedSeries = isStackedArea || isStackedLine;
  const isPercentStackedSeries = chartType === "AreaPercentStacked" || chartType === "LinePercentStacked";
  const resolvedValuesBySeries = chart.series.map((series) => (
    categories.map((_, categoryIndex) => resolveRenderableSeriesValue(series.values[categoryIndex], displayBlanksAs))
  ));

  type SeriesPoint = {
    defined: boolean;
    y: number | null;
    y0: number | null;
    y1: number | null;
  };

  const stackedPointsBySeries: SeriesPoint[][] = isStackedSeries
    ? (() => {
        const positive = Array.from({ length: categories.length }, () => 0);
        const negative = Array.from({ length: categories.length }, () => 0);
        const categoryTotals = isPercentStackedSeries
          ? categories.map((_, categoryIndex) => {
              const total = resolvedValuesBySeries.reduce((sum, seriesValues) => (
                sum + (seriesValues[categoryIndex] ?? 0)
              ), 0);
              return Math.abs(total) < 1e-9 ? 1 : total;
            })
          : null;
        return chart.series.map((_, seriesIndex) => (
          categories.map((_, categoryIndex) => {
            const rawValue = resolvedValuesBySeries[seriesIndex]?.[categoryIndex] ?? null;
            if (rawValue == null) {
              return { defined: false, y: null, y0: null, y1: null };
            }
            const value = isPercentStackedSeries
              ? (rawValue / (categoryTotals?.[categoryIndex] ?? 1)) * 100
              : rawValue;
            if (value >= 0) {
              const start = positive[categoryIndex];
              const end = start + value;
              positive[categoryIndex] = end;
              return { defined: true, y: end, y0: start, y1: end };
            }
            const start = negative[categoryIndex];
            const end = start + value;
            negative[categoryIndex] = end;
            return { defined: true, y: end, y0: start, y1: end };
          })
        ));
      })()
    : chart.series.map((_, seriesIndex) => (
        categories.map((_, categoryIndex) => {
          const value = resolvedValuesBySeries[seriesIndex]?.[categoryIndex] ?? null;
          return {
            defined: value != null,
            y: value,
            y0: null,
            y1: value
          };
        })
      ));

  const stackedExtents = stackedPointsBySeries
    .flatMap((seriesPoints) => (
      seriesPoints.flatMap((point) => (
        point.defined
          ? [point.y0, point.y1].filter((value): value is number => value != null && Number.isFinite(value))
          : []
      ))
    ));
  if (stackedExtents.length === 0) {
    return null;
  }
  let minValue = Math.min(...stackedExtents);
  let maxValue = Math.max(...stackedExtents);
  if (isAreaChart) {
    minValue = Math.min(0, minValue);
  } else if ((chartType === "Line" || isStackedLine) && chart.valueAxis?.crosses === "autoZero") {
    minValue = Math.min(0, minValue);
  }
  if (isPercentStackedSeries) {
    const explicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min)
      ? Number(chart.valueAxis.min)
      : Math.min(0, minValue);
    const explicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max)
      ? Number(chart.valueAxis.max)
      : Math.max(100, maxValue);
    minValue = explicitMin;
    maxValue = explicitMax <= explicitMin ? explicitMin + 1 : explicitMax;
  } else {
    const valueDomain = resolveAxisDomainWithChartOverrides(
      chart.valueAxis,
      minValue,
      maxValue,
      isAreaChart || chart.valueAxis?.crosses === "autoZero"
    );
    minValue = valueDomain.min;
    maxValue = valueDomain.max;
  }

  const xScale = scalePoint<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width]);
  const yScale = scaleLinear()
    .domain([minValue, maxValue])
    .range([plot.top + plot.height, plot.top]);

  const ticks = isPercentStackedSeries
    ? buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit ?? 20)
    : resolveAxisDomainWithChartOverrides(
        chart.valueAxis,
        minValue,
        maxValue,
        isAreaChart || chart.valueAxis?.crosses === "autoZero"
      ).ticks;
  const categoryPositions = categories.map((category) => xScale(category) ?? plot.left);

  const curve: CurveFactory = curveLinear;
  const areaBaseline = yScale(Math.max(minValue, 0));
  const plotPointsBySeries = chart.series.map((_, seriesIndex) => (
    categories.map((category, categoryIndex) => {
      const point = stackedPointsBySeries[seriesIndex]?.[categoryIndex] ?? {
        defined: false,
        y: null,
        y0: null,
        y1: null
      };
      return {
        defined: point.defined,
        x: xScale(category) ?? plot.left,
        y: point.y,
        y0: point.y0,
        y1: point.y1
      };
    })
  ));

  if (chart.is3d) {
    return renderLineOrAreaChart3d(
      chart,
      palette,
      layout,
      categories,
      stackedPointsBySeries,
      minValue,
      maxValue,
      isAreaChart,
      isStackedSeries,
      selectedChartElement
    );
  }

  const rawRecord = chart.raw && typeof chart.raw === "object"
    ? chart.raw as Record<string, unknown>
    : null;
  const dropLinesRecord = rawRecord?.dropLines && typeof rawRecord.dropLines === "object"
    ? rawRecord.dropLines as Record<string, unknown>
    : null;
  const highLowLinesRecord = rawRecord?.highLowLines && typeof rawRecord.highLowLines === "object"
    ? rawRecord.highLowLines as Record<string, unknown>
    : null;
  const upDownBarsRecord = rawRecord?.upDownBars && typeof rawRecord.upDownBars === "object"
    ? rawRecord.upDownBars as Record<string, unknown>
    : null;
  const dropLineColor = normalizeRendererHexColor((dropLinesRecord?.shapeProperties as Record<string, unknown> | undefined)?.lineColorHex)
    ?? lightenColor(chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border, 0.45);
  const highLowLineColor = normalizeRendererHexColor((highLowLinesRecord?.shapeProperties as Record<string, unknown> | undefined)?.lineColorHex)
    ?? chart.axisLineColor
    ?? chart.chartAreaBorderColor
    ?? palette.border;
  const upDownGapWidth = typeof upDownBarsRecord?.gapWidth === "number" && Number.isFinite(upDownBarsRecord.gapWidth)
    ? upDownBarsRecord.gapWidth
    : 150;
  const categoryStep = categoryPositions.length >= 2
    ? Math.min(...categoryPositions.slice(1).map((position, index) => Math.abs(position - (categoryPositions[index] ?? 0))).filter((value) => value > 0))
    : plot.width;
  const upDownBarWidth = clamp(categoryStep * (1 - resolveCategoryBandPadding(upDownGapWidth).inner) * 0.82, 3, 28);

  return (
    <g>
      {renderCartesianAxes(
        chart,
        palette,
        plot,
        false,
        categories,
        categoryPositions,
        ticks,
        (value) => yScale(value)
      )}
      {!isAreaChart && !isStackedSeries && highLowLinesRecord
        ? categories.map((_, categoryIndex) => {
            const yValues = plotPointsBySeries
              .map((seriesPoints) => seriesPoints[categoryIndex]?.y)
              .filter((value): value is number => value != null && Number.isFinite(value));
            if (yValues.length < 2) {
              return null;
            }
            const x = categoryPositions[categoryIndex] ?? plot.left;
            return (
              <line
                key={`high-low-${categoryIndex}`}
                stroke={highLowLineColor}
                stroke-width={1}
                x1={x}
                x2={x}
                y1={yScale(Math.min(...yValues))}
                y2={yScale(Math.max(...yValues))}
              />
            );
          })
        : null}
      {!isAreaChart && !isStackedSeries && upDownBarsRecord && plotPointsBySeries.length >= 2
        ? categories.map((_, categoryIndex) => {
            const first = plotPointsBySeries[0]?.[categoryIndex];
            const second = plotPointsBySeries[1]?.[categoryIndex];
            if (!first || !second || first.y == null || second.y == null) {
              return null;
            }
            const top = yScale(Math.max(first.y, second.y));
            const bottom = yScale(Math.min(first.y, second.y));
            const isUpBar = second.y >= first.y;
            const fill = isUpBar ? darkenColor(highLowLineColor, 0.12) : "#c0504d";
            return (
              <rect
                key={`up-down-${categoryIndex}`}
                fill={fill}
                fill-opacity={0.92}
                height={Math.max(1, bottom - top)}
                stroke={darkenColor(fill, 0.18)}
                stroke-width={1}
                width={upDownBarWidth}
                x={(categoryPositions[categoryIndex] ?? plot.left) - upDownBarWidth / 2}
                y={top}
              />
            );
          })
        : null}
      {chart.series.map((series, seriesIndex) => {
        const points = plotPointsBySeries[seriesIndex] ?? [];
        const lineStrokePoints = displayBlanksAs === "span"
          ? points.filter((point) => point.y != null)
          : points;
        const linePath = d3Line<{ x: number; y: number | null }>()
          .defined((point) => point.y != null)
          .x((point) => point.x)
          .y((point) => yScale(point.y ?? 0))
          .curve(curve)(lineStrokePoints) ?? "";

        const areaPath = isAreaChart
          ? d3Area<{ x: number; y: number | null; y0: number | null; y1: number | null }>()
            .defined((point) => (
              isStackedSeries
                ? point.y0 != null && point.y1 != null
                : point.y != null
            ))
            .x((point) => point.x)
            .y0((point) => (
              isStackedSeries
                ? yScale(point.y0 ?? 0)
                : areaBaseline
            ))
            .y1((point) => yScale((isStackedSeries ? point.y1 : point.y) ?? 0))
            .curve(curve)(points) ?? ""
          : "";

        const seriesFillColor = typeof series.shapeProperties?.xmlFillColor === "string"
          ? series.shapeProperties.xmlFillColor
          : chartSeriesColor(chart, seriesIndex);
        return (
          <g key={`line-series-${seriesIndex}`}>
            {!isAreaChart && !isStackedSeries && dropLinesRecord
              ? points.map((point, pointIndex) => (
                  point.y == null ? null : (
                    <line
                      key={`drop-line-${seriesIndex}-${pointIndex}`}
                      stroke={dropLineColor}
                      stroke-width={1}
                      x1={point.x}
                      x2={point.x}
                      y1={yScale(point.y)}
                      y2={areaBaseline}
                    />
                  )
                ))
              : null}
            {isAreaChart ? (
              <path
                {...chartElementDataProps(seriesIndex)}
                d={areaPath}
                fill={seriesFillColor}
                fill-opacity={1}
                stroke="none"
              />
            ) : null}
            <path
              {...chartElementDataProps(seriesIndex)}
              d={linePath}
              fill="none"
              stroke={chartSeriesStrokeColor(chart, seriesIndex)}
              stroke-linecap="round"
              stroke-linejoin="round"
              stroke-width={Math.max(1.5, series.lineWidthPx ?? 2)}
            />
            {points.map((point, pointIndex) => (
              point.y == null ? null : (
                <circle
                  {...chartElementDataProps(seriesIndex, pointIndex)}
                  key={`line-marker-${seriesIndex}-${pointIndex}`}
                  cx={point.x}
                  cy={yScale(point.y)}
                  fill={series.markerColor ?? chartSeriesColor(chart, seriesIndex)}
                  r={Math.max(2, (series.markerSize ?? 6) * 0.25)}
                  stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? chartSeriesStrokeColor(chart, seriesIndex)}
                  stroke-width={1}
                />
              )
            ))}
            {points.map((point, pointIndex) => (
              point.y != null && (
                isSelectedChartPoint(selectedChartElement, chart.id, seriesIndex, pointIndex)
                || (
                  isSelectedChartSeries(selectedChartElement, chart.id, seriesIndex)
                  && selectedChartElement?.kind !== "point"
                )
              )
                ? isAreaChart
                  ? renderSelectionPointHandles(`area-selection-${seriesIndex}-${pointIndex}`, [{ x: point.x, y: yScale(point.y) }])
                  : renderSelectionRectHandles(
                      `line-selection-${seriesIndex}-${pointIndex}`,
                      point.x - 4,
                      yScale(point.y) - 4,
                      8,
                      8
                    )
                : null
            ))}
          </g>
        );
      })}
    </g>
  );
}

export function renderComboChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const groups = buildComboGroups(chart);
  const columnGroup = groups.find((group) => group.chartType.startsWith("Column"));
  const lineGroup = groups.find((group) => group.chartType.startsWith("Line"));
  if (!columnGroup || !lineGroup) {
    return null;
  }

  const primaryChart: XlsxChart = {
    ...chart,
    categoryAxis: columnGroup.categoryAxis ?? chart.categoryAxis,
    chartType: columnGroup.chartType,
    dataLabels: undefined,
    is3d: columnGroup.is3d ?? false,
    series: columnGroup.series,
    valueAxis: columnGroup.valueAxis ?? chart.valueAxis
  };
  const categories = getCategoryLabels(primaryChart);
  if (categories.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const groupGapWidth = typeof columnGroup.raw?.gapWidth === "number" && Number.isFinite(columnGroup.raw.gapWidth)
    ? columnGroup.raw.gapWidth
    : columnGroup.gapWidth ?? chart.gapWidth;
  const histogramColumns = columnGroup.series.some((series) => isHistogramLikeSeries(series));
  const categoryBandPadding = histogramColumns
    ? { inner: 0, outer: 0 }
    : resolveCategoryBandPadding(groupGapWidth);
  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(categoryBandPadding.inner)
    .paddingOuter(categoryBandPadding.outer);
  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: columnGroup.series.length }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(histogramColumns ? 0 : 0.16)
    .paddingOuter(histogramColumns ? 0 : 0.08);
  const categoryPositions = categories.map((category) => (
    (categoryScale(category) ?? plot.left) + categoryScale.bandwidth() / 2
  ));
  const resolveGlobalSeriesIndex = (series: XlsxChartSeries) => {
    const index = chart.series.findIndex((candidate) => candidate.id === series.id);
    return index >= 0 ? index : 0;
  };

  const primaryValues = columnGroup.series.flatMap((series) => (
    series.values.map((value) => safeNumber(value)).filter((value): value is number => value != null)
  ));
  const secondaryValues = lineGroup.series.flatMap((series) => (
    series.values.map((value) => safeNumber(value)).filter((value): value is number => value != null)
  ));
  if (primaryValues.length === 0 || secondaryValues.length === 0) {
    return null;
  }

  const primaryDomain = resolveNumericAxisDomain(
    typeof columnGroup.valueAxis?.min === "number" ? columnGroup.valueAxis.min : Math.min(...primaryValues),
    typeof columnGroup.valueAxis?.max === "number" ? columnGroup.valueAxis.max : Math.max(...primaryValues),
    columnGroup.valueAxis?.majorUnit,
    true
  );
  const secondaryDomain = resolveNumericAxisDomain(
    typeof lineGroup.valueAxis?.min === "number" ? lineGroup.valueAxis.min : Math.min(...secondaryValues),
    typeof lineGroup.valueAxis?.max === "number" ? lineGroup.valueAxis.max : Math.max(...secondaryValues),
    lineGroup.valueAxis?.majorUnit,
    false
  );
  const primaryScale = scaleLinear()
    .domain([primaryDomain.min, primaryDomain.max])
    .range([plot.top + plot.height, plot.top]);
  const secondaryScale = scaleLinear()
    .domain([secondaryDomain.min, secondaryDomain.max])
    .range([plot.top + plot.height, plot.top]);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);

  return (
    <g>
      {renderCartesianAxes(
        primaryChart,
        palette,
        plot,
        false,
        categories,
        categoryPositions,
        primaryDomain.ticks,
        (value) => primaryScale(value)
      )}
      {secondaryDomain.ticks.map((tick) => {
        const y = secondaryScale(tick);
        return (
          <text
            key={`combo-secondary-tick-${tick}`}
            fill={labelColor}
            font-size={10}
            text-anchor="start"
            x={plot.left + plot.width + 6}
            y={y + 3}
          >
            {formatTickValue(tick)}
          </text>
        );
      })}
      <line
        stroke={axisColor}
        stroke-width={1.2}
        x1={plot.left + plot.width}
        x2={plot.left + plot.width}
        y1={plot.top}
        y2={plot.top + plot.height}
      />
      {columnGroup.series.flatMap((series, seriesIndex) => categories.map((category, categoryIndex) => {
        const globalSeriesIndex = resolveGlobalSeriesIndex(series);
        const value = safeNumber(series.values[categoryIndex]) ?? 0;
        const categoryStart = categoryScale(category) ?? plot.left;
        const barWidth = Math.max(1, histogramColumns ? categoryScale.bandwidth() : seriesScale.bandwidth());
        const x = categoryStart + (histogramColumns ? 0 : (seriesScale(String(seriesIndex)) ?? 0));
        const y = primaryScale(Math.max(0, value));
        const zeroY = primaryScale(0);
        const height = Math.max(1, Math.abs(zeroY - primaryScale(value)));
        const selected = isSelectedChartPoint(selectedChartElement, chart.id, globalSeriesIndex, categoryIndex)
          || (
            isSelectedChartSeries(selectedChartElement, chart.id, globalSeriesIndex)
            && selectedChartElement?.kind !== "point"
          );
        return (
          <>
            <rect
              {...barChartElementDataProps(globalSeriesIndex, categoryIndex)}
              fill={series.color ?? series.lineColor ?? chartSeriesColor(primaryChart, seriesIndex)}
              height={height}
              stroke={selected ? "#64748b" : (histogramColumns ? "none" : (series.lineColor ?? series.color ?? chartSeriesStrokeColor(primaryChart, seriesIndex)))}
              stroke-width={selected ? 2 : (histogramColumns ? 0 : 1)}
              width={barWidth}
              x={x}
              y={Math.min(y, zeroY)}
            />
            {selected
              ? renderSelectionRectHandles(
                  `combo-bar-selection-${globalSeriesIndex}-${categoryIndex}`,
                  x,
                  Math.min(y, zeroY),
                  barWidth,
                  height
                )
              : null}
          </>
        );
      }))}
      {lineGroup.series.map((series, seriesIndex) => {
        const globalSeriesIndex = resolveGlobalSeriesIndex(series);
        const points = categories.map((category, categoryIndex) => ({
          x: (categoryScale(category) ?? plot.left) + categoryScale.bandwidth() / 2,
          y: safeNumber(series.values[categoryIndex])
        }));
        const lineStrokePoints = chart.displayBlanksAs === "span"
          ? points.filter((point) => point.y != null)
          : points;
        const path = d3Line<{ x: number; y: number | null }>()
          .defined((point) => point.y != null)
          .x((point) => point.x)
          .y((point) => secondaryScale(point.y ?? 0))
          .curve(curveLinear)(lineStrokePoints) ?? "";
        return (
          <g key={`combo-line-${seriesIndex}`}>
            <path
              {...chartElementDataProps(globalSeriesIndex)}
              d={path}
              fill="none"
              stroke={series.lineColor ?? series.color ?? chartSeriesStrokeColor(chart, columnGroup.series.length + seriesIndex)}
              stroke-linecap="round"
              stroke-linejoin="round"
              stroke-width={Math.max(1.5, series.lineWidthPx ?? 2)}
            />
            {points.map((point, pointIndex) => (
              point.y == null ? null : (
                <path
                  {...chartElementDataProps(globalSeriesIndex, pointIndex)}
                  d={markerSymbolPath(normalizeChartMarkerSymbol(series.markerSymbol), Math.max(4, series.markerSize ?? 7)) || markerSymbolPath("circle", 7)}
                  fill={series.markerColor ?? series.color ?? series.lineColor ?? chartSeriesColor(chart, columnGroup.series.length + seriesIndex)}
                  key={`combo-line-marker-${seriesIndex}-${pointIndex}`}
                  stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? "#ffffff"}
                  stroke-width={1}
                  transform={`translate(${point.x}, ${secondaryScale(point.y)})`}
                />
              )
            ))}
            {points.map((point, pointIndex) => (
              point.y != null && isSelectedChartPointOrSeries(selectedChartElement, chart.id, globalSeriesIndex, pointIndex)
                ? renderSelectionRectHandles(
                    `combo-line-selection-${globalSeriesIndex}-${pointIndex}`,
                    point.x - 4,
                    secondaryScale(point.y) - 4,
                    8,
                    8
                  )
                : null
            ))}
          </g>
        );
      })}
    </g>
  );
}

export function renderRadarChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  if (chart.series.length === 0) {
    return null;
  }
  const categories = getCategoryLabels(chart);
  if (categories.length === 0) {
    return null;
  }
  const plot = layout.plot;
  const centerX = plot.left + plot.width * 0.5;
  const centerY = plot.top + plot.height * 0.52;
  const radius = Math.max(22, Math.min(plot.width, plot.height) * 0.38);
  const values = chart.series.flatMap((series) => (
    series.values
      .slice(0, categories.length)
      .map((value) => safeNumber(value))
      .filter((value): value is number => value != null)
  ));
  const hasExplicitMin = typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min);
  const hasExplicitMax = typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max);
  let minValue = hasExplicitMin
    ? Number(chart.valueAxis?.min)
    : Math.min(0, ...(values.length ? values : [0]));
  let maxValue = hasExplicitMax
    ? Number(chart.valueAxis?.max)
    : Math.max(1, ...(values.length ? values : [1]));
  if (chart.valueAxis?.crosses === "autoZero") {
    minValue = Math.min(0, minValue);
  }
  const safeMax = maxValue <= minValue ? minValue + 1 : maxValue;
  const span = Math.max(1e-6, safeMax - minValue);
  const candidateMajorUnit = typeof chart.valueAxis?.majorUnit === "number" && chart.valueAxis.majorUnit > 0
    ? chart.valueAxis.majorUnit
    : undefined;
  const preferredMajorUnit = candidateMajorUnit && span / candidateMajorUnit >= 3
    ? candidateMajorUnit
    : buildNiceStep(minValue, safeMax, 6);
  const ticks = buildNumericTickValues(minValue, safeMax, preferredMajorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
  const filled = chart.radarStyle === "filled";
  const showSpokes = chart.categoryAxis?.majorGridlines === true;

  const angleAt = (index: number) => (Math.PI * 2 * index) / categories.length - Math.PI / 2;
  const radialPoint = (index: number, valueRatio: number) => {
    const angle = angleAt(index);
    const r = radius * valueRatio;
    return {
      x: centerX + Math.cos(angle) * r,
      y: centerY + Math.sin(angle) * r
    };
  };

  return (
    <g>
      {ticks.map((tick, ringIndex) => {
        const ratio = (tick - minValue) / (safeMax - minValue);
        const ringPoints = categories.map((_, categoryIndex) => radialPoint(categoryIndex, ratio));
        const ringPath = d3Line<{ x: number; y: number }>()
          .x((point) => point.x)
          .y((point) => point.y)
          .curve(curveLinearClosed)(ringPoints) ?? "";
        return (
          <g key={`radar-ring-${ringIndex}`}>
            <path
              d={ringPath}
              fill={ringIndex % 2 === 0 ? "transparent" : lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.04)}
              stroke={lightenColor(axisColor, 0.5)}
              stroke-width={1}
            />
            <text fill={labelColor} font-size={9} x={centerX + 4} y={centerY - radius * ratio + 3}>
              {formatTickValue(tick)}
            </text>
          </g>
        );
      })}
      {categories.map((category, categoryIndex) => {
        const edge = radialPoint(categoryIndex, 1);
        return (
          <g key={`radar-axis-${categoryIndex}`}>
            {showSpokes ? (
              <line stroke={lightenColor(axisColor, 0.52)} stroke-width={1} x1={centerX} x2={edge.x} y1={centerY} y2={edge.y} />
            ) : null}
            <text
              fill={labelColor}
              font-size={10}
              text-anchor="middle"
              x={centerX + (edge.x - centerX) * 1.1}
              y={centerY + (edge.y - centerY) * 1.1}
            >
              {category}
            </text>
          </g>
        );
      })}
      {chart.series.map((series, seriesIndex) => {
        const points = categories.map((_, categoryIndex) => {
          const rawValue = safeNumber(series.values[categoryIndex]);
          if (rawValue == null) {
            return {
              defined: false,
              pointIndex: categoryIndex,
              ...radialPoint(categoryIndex, 0)
            };
          }
          const ratio = clamp((rawValue - minValue) / (safeMax - minValue), 0, 1);
          return {
            defined: true,
            pointIndex: categoryIndex,
            ...radialPoint(categoryIndex, ratio)
          };
        });
        const definedPoints = points.filter((point) => point.defined);
        if (definedPoints.length === 0) {
          return null;
        }
        const hasGap = points.some((point) => !point.defined);
        const polygon = d3Line<{ defined: boolean; x: number; y: number }>()
          .defined((point) => point.defined)
          .x((point) => point.x)
          .y((point) => point.y)
          .curve(hasGap ? curveLinear : curveLinearClosed)(points) ?? "";
        const color = chartSeriesColor(chart, seriesIndex);
        const markerSymbol = normalizeChartMarkerSymbol(series.markerSymbol);
        const markerSize = Math.max(4, series.markerSize ?? 6);
        const markerPath = markerSymbolPath(markerSymbol, markerSize * 0.5);
	        const showMarkers = markerSymbol !== "none" && markerPath.length > 0;
	        return (
	          <g key={`radar-series-${seriesIndex}`}>
	            {definedPoints.length >= 2 ? (
	              <path
	                {...chartElementDataProps(seriesIndex)}
	                d={polygon}
	                fill={filled && !hasGap && definedPoints.length >= 3 ? color : "none"}
	                fill-opacity={filled ? 0.26 : 0}
	                stroke={chartSeriesStrokeColor(chart, seriesIndex)}
	                stroke-width={1.8}
	              />
	            ) : null}
	            {showMarkers
	              ? definedPoints.map((point) => (
	                  <g
	                    key={`radar-point-${seriesIndex}-${point.pointIndex}`}
	                    transform={`translate(${point.x}, ${point.y})`}
	                  >
	                    <path
	                      {...chartElementDataProps(seriesIndex, point.pointIndex)}
	                      d={markerPath}
	                      fill={series.markerColor ?? color}
	                      stroke={series.markerLineColor ?? chart.chartAreaFillColor ?? palette.surface}
	                      stroke-width={1}
	                    />
	                  </g>
	                ))
	              : null}
	            {definedPoints.map((point) => (
	              isSelectedChartPoint(selectedChartElement, chart.id, seriesIndex, point.pointIndex)
	              || (
	                isSelectedChartSeries(selectedChartElement, chart.id, seriesIndex)
	                && selectedChartElement?.kind !== "point"
	              )
	                ? renderSelectionPointHandles(
	                    `radar-selection-${seriesIndex}-${point.pointIndex}`,
	                      [{ x: point.x, y: point.y }]
	                    )
	                : null
	            ))}
	          </g>
	        );
      })}
    </g>
  );
}

export function renderStockChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const categories = getCategoryLabels(chart);
  if (categories.length === 0 || chart.series.length < 2) {
    return null;
  }

  const roles = resolveStockRoleIndices(chart);
  const highIndex = roles.high ?? 0;
  const lowIndex = roles.low ?? Math.min(1, chart.series.length - 1);
  const closeIndex = roles.close ?? Math.min(2, chart.series.length - 1);
  const openIndex = roles.open;
  const volumeIndex = roles.volume != null && ![highIndex, lowIndex, closeIndex, openIndex].includes(roles.volume)
    ? roles.volume
    : null;
  const high = chart.series[highIndex];
  const low = chart.series[lowIndex];
  const close = chart.series[closeIndex];
  const open = openIndex != null ? chart.series[openIndex] : null;
  const volume = volumeIndex != null ? chart.series[volumeIndex] : null;

  const points = categories
    .map((category, index) => {
      const highValue = safeNumber(high.values[index]);
      const lowValue = safeNumber(low.values[index]);
      const closeValue = safeNumber(close.values[index]);
      const openValue = open ? safeNumber(open.values[index]) : null;
      const volumeValue = volume ? safeNumber(volume.values[index]) : null;
      if (highValue == null || lowValue == null || closeValue == null) {
        return null;
      }
      return {
        category,
        close: closeValue,
        high: highValue,
        low: lowValue,
        open: openValue,
        volume: volumeValue
      };
    })
    .filter((entry): entry is { category: string; close: number; high: number; low: number; open: number | null; volume: number | null } => entry != null);
  if (points.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const hasVolume = volume != null && points.some((entry) => (entry.volume ?? 0) > 0);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const stockPalette = resolveStockPalette(chart, axisColor);
  const rawMinValue = hasVolume
    ? Math.min(0, ...points.flatMap((entry) => [entry.low, entry.close, entry.open ?? entry.close, entry.volume ?? 0]))
    : Math.min(...points.map((entry) => entry.low));
  const rawMaxValue = hasVolume
    ? Math.max(...points.flatMap((entry) => [entry.high, entry.close, entry.open ?? entry.close, entry.volume ?? 0]))
    : Math.max(...points.map((entry) => entry.high));
  const resolvedDomain = resolveNumericAxisDomain(
    typeof chart.valueAxis?.min === "number" && Number.isFinite(chart.valueAxis.min) ? chart.valueAxis.min : rawMinValue,
    typeof chart.valueAxis?.max === "number" && Number.isFinite(chart.valueAxis.max) ? chart.valueAxis.max : rawMaxValue,
    chart.valueAxis?.majorUnit,
    hasVolume
  );
  const yScale = scaleLinear()
    .domain([resolvedDomain.min, resolvedDomain.max])
    .range([plot.top + plot.height, plot.top]);
  const xScale = scaleBand<string>()
    .domain(categories)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(0.28)
    .paddingOuter(0.18);
  const ticks = resolvedDomain.ticks;
  const labelColor = resolveChartAxisTextColor(chart);
  const openTickColor = stockPalette.openAccent;
  const closeTickColor = openIndex != null ? stockPalette.lineColor : stockPalette.closeAccent;
  const labelStep = Math.max(1, Math.ceil(categories.length / Math.max(4, Math.floor(plot.width / 68))));
  const zeroY = yScale(Math.max(resolvedDomain.min, 0));

  return (
    <g>
      {ticks.map((tick) => (
        <g key={`stock-tick-${tick}`}>
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
      {categories.map((category, index) => {
        if (index % labelStep !== 0 && index !== categories.length - 1) {
          return null;
        }
        const x = (xScale(category) ?? plot.left) + xScale.bandwidth() * 0.5;
        return (
          <text
            key={`stock-cat-${category}`}
            fill={labelColor}
            font-size={10}
            text-anchor="middle"
            x={x}
            y={plot.top + plot.height + 14}
          >
            {category}
          </text>
        );
      })}
      <line stroke={axisColor} stroke-width={1.2} x1={plot.left} x2={plot.left} y1={plot.top} y2={plot.top + plot.height} />
      <line stroke={axisColor} stroke-width={1.2} x1={plot.left} x2={plot.left + plot.width} y1={plot.top + plot.height} y2={plot.top + plot.height} />
      {points.map((entry, index) => {
        const x = (xScale(entry.category) ?? plot.left) + xScale.bandwidth() * 0.5;
        const previousClose = index > 0 ? points[index - 1]?.close ?? entry.close : entry.close;
        const isUp = entry.open != null
          ? entry.close >= entry.open
          : entry.close >= previousClose;
        const stroke = stockPalette.lineColor;
        const candleWidth = Math.max(5, xScale.bandwidth() * 0.56);
        const bodyLeft = x - candleWidth / 2;
        const openY = entry.open != null ? yScale(entry.open) : null;
        const closeY = yScale(entry.close);
        const highY = yScale(entry.high);
        const lowY = yScale(entry.low);
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, closeIndex, index);
        return (
            <g
              {...chartElementDataProps(closeIndex, index)}
              key={`stock-point-${index}`}
            >
            {hasVolume && entry.volume != null ? (
              <rect
                fill={stockPalette.volumeFill}
                height={Math.max(1, zeroY - yScale(entry.volume))}
                opacity={0.94}
                stroke="none"
                width={Math.max(3, xScale.bandwidth() * 0.64)}
                x={x - Math.max(3, xScale.bandwidth() * 0.64) / 2}
                y={yScale(entry.volume)}
              />
            ) : null}
            <line stroke={stroke} stroke-width={1.6} x1={x} x2={x} y1={highY} y2={lowY} />
            {entry.open != null && openY != null ? (
              Math.abs(openY - closeY) >= 1 ? (
                <rect
                  fill={isUp ? stockPalette.upFill : stockPalette.downFill}
                  height={Math.max(1.4, Math.abs(closeY - openY))}
                  stroke={stroke}
                  stroke-width={1.3}
                  width={candleWidth}
                  x={bodyLeft}
                  y={Math.min(openY, closeY)}
                />
              ) : (
                <line stroke={stroke} stroke-width={1.6} x1={bodyLeft} x2={bodyLeft + candleWidth} y1={closeY} y2={closeY} />
              )
            ) : (
              <line stroke={closeTickColor} stroke-width={1.8} x1={x} x2={x + 7} y1={closeY} y2={closeY} />
            )}
            {entry.open != null && openY != null ? (
              <line stroke={openTickColor} stroke-width={1.8} x1={x - 7} x2={x} y1={openY} y2={openY} />
            ) : null}
            {selected
              ? renderSelectionPointHandles(`stock-selection-${index}`, [{ x, y: closeY }])
              : null}
          </g>
        );
      })}
    </g>
  );
}
