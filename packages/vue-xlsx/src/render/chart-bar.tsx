import { scaleBand, scaleLinear, scalePoint } from "d3-scale";
import { area as d3Area, curveLinear, line as d3Line } from "d3-shape";
import type { CurveFactory } from "d3-shape";
import type { XlsxChart, XlsxChartElementSelection, XlsxChartSeries, XlsxImageRect } from "@extend-ai/xlsx-core";
import { chartSeriesColor, chartSeriesStrokeColor, chartPointColor, safeNumber, clamp, darkenColor, lightenColor, coerceLooseNumber, resolveRenderableSeriesValue, resolveChartTextColor, resolveChartAxisTextColor, resolveChartMutedTextColor, buildChartFontFamily, formatTickValue, formatPercentTickValue, chartElementDataProps, barChartElementDataProps, resolveChartElementTarget, resolveChartSelectionFromTarget, isSelectedChartSeries, isSelectedChartPoint, isSelectedChartPointOrSeries, renderSelectionRectHandles, renderSelectionPointHandles, normalizeRendererHexColor, toSvgNumber, normalizeChartMarkerSymbol, markerSymbolPath, normalizeLegendPosition, parseRgbColor, mixRgbColor, escapeCssFontFamilyToken, DEFAULT_CHART_FONT_STACK, DEFAULT_CHART_TEXT_COLOR, DEFAULT_CHART_MUTED_TEXT_COLOR, truncateSvgText, getCategoryLabels, normalizeCategoryLabel, estimateReferencePointCount, isHistogramLikeSeries, isHistogramLikeChart, resolve3dFrameOffsets, resolveAxisDomainWithChartOverrides, resolveCategoryBandPadding, chartSeriesBarColors, buildNumericTickValues, buildNiceStep, resolveNumericAxisDomain, renderCartesianAxes, renderExtrudedRect, buildChartStages, resolveChartStageSubtotal, computeBoxWhiskerStats, resolveBoxWhiskerVisibility, resolveBoxWhiskerQuartileMethod, computePercentile, renderLineOrAreaChart3d, projectCartesian3dPoint, scaleProjectedVector, buildLinearSvgPath, buildRibbonSvgPath, renderRadialFrustum, sampleProjectedEllipseArc, isComboChart, getComboLegendSeries, findAxisForGroup, buildComboGroups, resolveStockRoleIndices, resolveStockPalette, normalizeBuiltinPieStyleId, getBuiltinPiePalette, resolvePiePointColor, selectPrimaryPieSeriesIndex, buildPieEntries, buildHierarchyData, resolveHierarchyNodeColor, resolveTreemapNodeColor, excelTreemapTile, normalizePieArc, resolvePieFrontSegments, pieEllipsePoint, buildPieOuterWallPath, isPieFrontFacingAngle, buildPieRadialWallPath, TWO_PI, normalizeRegionMapKey, resolveRegionMapFeature, resolveRegionMapBaseColor, resolveRegionMapDataColor, resolveRegionMapValueColors, resolveRegionMapColorStrings, resolveRegionMapLayoutProperties, resolveRegionMapFeatureSet, getRegionMapBaseFeatures, resolveRegionMapValueColorFromStops, resolveRegionMapValueColor, resolveRegionMapNoDataColor, resolveSurfaceBaseColor, normalizeBuiltinSurfaceStyleId, hasExplicitSurfaceBaseColor, buildMonochromeSurfacePalette, getBuiltinSurfacePalette, shouldPreferBuiltinSurfacePalette, getSurfaceBandCount, getSurfaceColorStops, resolveSurfaceBandPaletteColor, getSurfaceDomain, resolveSurfaceColor, resolveSurfaceBandColor, resolveSurfaceBandIndex, buildSurfaceSmoothPath, splitSurfacePointRuns, connectSurfaceContourSegments, getSurfaceWireframePalette, resolveSurfaceWireframeColor, resolveSurfacePlotRect, buildSurfaceLegendItems, isContourSurfaceChart, getLegendItems, buildLayout, normalizeRenderableChartType, renderSurfaceAxes, renderTitle, renderLegend, buildRegionMapLegendItems, buildStockLegendItems, excelSerialToDate, isLikelyDateFormatCode, formatExcelDateSerial, formatCategoryLabel } from "./chart-renderer";
import type { ChartRendererPalette, ChartSvgProps, ChartLayout, LegendItem, BarRect, PlotRect, ChartStage, ChartHierarchyDatum, ComboRenderableGroup, BoxWhiskerStats, SurfaceDomain, RegionMapFeature, SurfacePathPoint, SurfaceContourSegment, ChartElementDataOptions } from "./chart-renderer";


export function renderBarChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string, selectedChartElement?: XlsxChartElementSelection | null) {
  if (chart.series.length === 0) {
    return null;
  }
  const sourceCategories = getCategoryLabels(chart);
  if (sourceCategories.length === 0) {
    return null;
  }
  const isPercentStacked = chartType === "ColumnPercentStacked" || chartType === "BarPercentStacked";
  const isHorizontal = chartType === "BarClustered" || chartType === "BarStacked" || chartType === "BarPercentStacked";
  const isHistogramLike = isHistogramLikeChart(chart) && chartType === "ColumnClustered";
  const shouldReverseCategories = isHorizontal && chart.categoryAxis?.orientation !== "maxMin";
  const categories = shouldReverseCategories ? sourceCategories.slice().reverse() : sourceCategories;
  const isStacked = chartType === "ColumnStacked" || chartType === "BarStacked" || isPercentStacked;
  const negativeFillMode: "chartArea" | "none" | "series" = isPercentStacked
    ? "chartArea"
    : chartType === "ColumnClustered"
      ? "none"
      : "series";
  const categoryCount = categories.length;
  const seriesCount = chart.series.length;
  const normalized3dShape = (() => {
    switch (chart.shape3d) {
      case "cone":
      case "coneToMax":
        return "cone";
      case "cylinder":
        return "cylinder";
      case "pyramid":
      case "pyramidToMax":
        return "pyramid";
      default:
        return "box";
    }
  })();
  const frameOffsets = chart.is3d ? resolve3dFrameOffsets(chart, isHorizontal ? 9 : 11, isHorizontal ? 7 : 8) : null;
  const plot = chart.is3d && frameOffsets
    ? {
        ...layout.plot,
        height: Math.max(20, layout.plot.height - frameOffsets.insetBottom - 2),
        left: layout.plot.left + frameOffsets.insetLeft,
        top: layout.plot.top + frameOffsets.insetTop,
        width: Math.max(20, layout.plot.width - frameOffsets.insetLeft - frameOffsets.insetRight - 2)
      }
    : layout.plot;

  const matrix = chart.series.map((series) => {
    const values = Array.from(
      { length: categoryCount },
      (_, categoryIndex) => safeNumber(series.values[categoryIndex]) ?? 0
    );
    return shouldReverseCategories ? values.reverse() : values;
  });
  const rawMatrix = matrix.map((row) => row.slice());

  if (isPercentStacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      const total = matrix.reduce((sum, row) => sum + Math.max(0, row[categoryIndex] ?? 0), 0);
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
        matrix[seriesIndex][categoryIndex] = total > 0 ? ((matrix[seriesIndex][categoryIndex] ?? 0) / total) * 100 : 0;
      }
    }
  }

  const positiveTotals = Array.from({ length: categoryCount }, (_, categoryIndex) => (
    matrix.reduce((sum, row) => sum + Math.max(0, row[categoryIndex] ?? 0), 0)
  ));
  const negativeTotals = Array.from({ length: categoryCount }, (_, categoryIndex) => (
    Math.abs(matrix.reduce((sum, row) => sum + Math.min(0, row[categoryIndex] ?? 0), 0))
  ));

  let minValue = Number.POSITIVE_INFINITY;
  let maxValue = Number.NEGATIVE_INFINITY;
  if (isStacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      let positive = 0;
      let negative = 0;
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
        const value = matrix[seriesIndex][categoryIndex] ?? 0;
        if (value >= 0) {
          positive += value;
        } else {
          negative += value;
        }
      }
      maxValue = Math.max(maxValue, positive);
      minValue = Math.min(minValue, negative);
    }
  } else {
    for (const row of matrix) {
      for (const value of row) {
        maxValue = Math.max(maxValue, value);
        minValue = Math.min(minValue, value);
      }
    }
  }
  if (!Number.isFinite(minValue) || !Number.isFinite(maxValue)) {
    minValue = 0;
    maxValue = 1;
  }

  const valueDomain = resolveAxisDomainWithChartOverrides(
    chart.valueAxis,
    minValue,
    maxValue,
    true
  );
  minValue = valueDomain.min;
  maxValue = valueDomain.max;

  const categoryBandPadding = isHistogramLike
    ? { inner: 0, outer: 0 }
    : resolveCategoryBandPadding(chart.gapWidth);
  const categoryScale = scaleBand<string>()
    .domain(categories)
    .range(isHorizontal ? [plot.top, plot.top + plot.height] : [plot.left, plot.left + plot.width])
    .paddingInner(categoryBandPadding.inner)
    .paddingOuter(categoryBandPadding.outer);

  const seriesScale = scaleBand<string>()
    .domain(Array.from({ length: seriesCount }, (_, index) => String(index)))
    .range([0, categoryScale.bandwidth()])
    .paddingInner(isHistogramLike ? 0 : 0.16)
    .paddingOuter(isHistogramLike ? 0 : 0.08);
  const usesSeriesDepthAxis = (
    chart.is3d === true
    && !isHorizontal
    && !isStacked
    && seriesCount > 1
    && chart.seriesAxis != null
  );
  const depthLaneSpread = usesSeriesDepthAxis
    ? clamp(1.26 + seriesCount * 0.16, 1.4, 1.96)
    : 1;
  const depthGridSpanX = usesSeriesDepthAxis && frameOffsets
    ? frameOffsets.depthX * depthLaneSpread
    : 0;
  const depthGridSpanY = usesSeriesDepthAxis && frameOffsets
    ? frameOffsets.depthY * depthLaneSpread
    : 0;
  const depthSlotX = usesSeriesDepthAxis ? depthGridSpanX / Math.max(1, seriesCount) : 0;
  const depthSlotY = usesSeriesDepthAxis ? depthGridSpanY / Math.max(1, seriesCount) : 0;
  const depthBarX = usesSeriesDepthAxis
    ? Math.sign(depthSlotX || frameOffsets?.depthX || 1) * Math.max(8, Math.abs(depthSlotX) * 1.14)
    : frameOffsets?.depthX;
  const depthBarY = usesSeriesDepthAxis
    ? Math.sign(depthSlotY || frameOffsets?.depthY || -1) * Math.max(5.5, Math.abs(depthSlotY) * 1.12)
    : frameOffsets?.depthY;

  const shouldReverseValueAxis = chart.valueAxis?.orientation === "maxMin";
  const valueScale = scaleLinear()
    .domain([minValue, maxValue])
    .range(
      isHorizontal
        ? (shouldReverseValueAxis ? [plot.left + plot.width, plot.left] : [plot.left, plot.left + plot.width])
        : (shouldReverseValueAxis ? [plot.top, plot.top + plot.height] : [plot.top + plot.height, plot.top])
    );

  const ticks = isPercentStacked
    ? buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit ?? 20)
    : valueDomain.ticks;
  const categoryPositions = categories.map((category) => {
    const bandStart = categoryScale(category) ?? 0;
    return bandStart + categoryScale.bandwidth() / 2;
  });

  const bars: BarRect[] = [];
  const positive = Array.from({ length: categoryCount }, () => 0);
  const negative = Array.from({ length: categoryCount }, () => 0);

  for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      const value = matrix[seriesIndex][categoryIndex] ?? 0;
      const rawValue = rawMatrix[seriesIndex][categoryIndex] ?? value;
      const pointIndex = shouldReverseCategories ? (categoryCount - 1 - categoryIndex) : categoryIndex;
      const pointStyle = chart.series[seriesIndex]?.dataPointStyles?.find((entry) => entry.index === pointIndex);
      let start = 0;
      let end = value;
      if (isStacked) {
        if (value >= 0) {
          start = positive[categoryIndex];
          positive[categoryIndex] += value;
          end = positive[categoryIndex];
        } else {
          start = negative[categoryIndex];
          negative[categoryIndex] += value;
          end = negative[categoryIndex];
        }
      }

      const colors = chartSeriesBarColors(chart, seriesIndex, rawValue, negativeFillMode);
      const isInvertedNegative = rawValue < 0 && chart.series[seriesIndex]?.invertIfNegative === true;
      if (!(isInvertedNegative)) {
        if (pointStyle?.color) {
          colors.fill = pointStyle.color;
          if (!pointStyle.lineColor) {
            colors.stroke = pointStyle.color;
          }
        }
        if (pointStyle?.lineColor) {
          colors.stroke = pointStyle.lineColor;
        }
      }
      if (isInvertedNegative && isPercentStacked && chart.is3d) {
        colors.fill = "#ffffff";
      }
      const category = categories[categoryIndex];
      const categoryStart = categoryScale(category) ?? 0;
      const barThickness = usesSeriesDepthAxis
        ? Math.max(7, categoryScale.bandwidth() * 0.4)
        : isHistogramLike
          ? categoryScale.bandwidth()
        : isStacked
          ? categoryScale.bandwidth()
          : seriesScale.bandwidth();
      const barOffset = usesSeriesDepthAxis
        ? Math.max(0, (categoryScale.bandwidth() - barThickness) * 0.5)
        : isHistogramLike
          ? 0
        : isStacked
          ? 0
          : (seriesScale(String(seriesIndex)) ?? 0);
      const depthOffsetX = usesSeriesDepthAxis ? depthSlotX * seriesIndex : 0;
      const depthOffsetY = usesSeriesDepthAxis ? depthSlotY * seriesIndex : 0;
      const depthOrder = depthOffsetX * (frameOffsets?.depthX ?? 0) + depthOffsetY * (frameOffsets?.depthY ?? 0);
      const shapeTaper = normalized3dShape === "pyramid"
        ? 0.94
        : normalized3dShape === "cone"
          ? 1
          : 0;
      let topScale = 1;
      let bottomScale = 1;
      if (chart.is3d && shapeTaper > 0) {
        const scaleAt = (ratio: number) => {
          const clampedRatio = clamp(ratio, 0, 1);
          const taperedRatio = normalized3dShape === "cone"
            ? Math.pow(clampedRatio, isHorizontal ? 1.85 : 1.08)
            : clampedRatio;
          return clamp(1 - (taperedRatio * shapeTaper), normalized3dShape === "cone" ? 0 : 0.04, 1);
        };
        if (rawValue >= 0) {
          const total = Math.max(1e-6, isStacked ? positiveTotals[categoryIndex] : Math.abs(rawValue));
          bottomScale = scaleAt(isStacked ? start / total : 0);
          topScale = scaleAt(isStacked ? end / total : 1);
        } else {
          const total = Math.max(1e-6, isStacked ? negativeTotals[categoryIndex] : Math.abs(rawValue));
          topScale = scaleAt(isStacked ? Math.abs(start) / total : 0);
          bottomScale = scaleAt(isStacked ? Math.abs(end) / total : 1);
        }
      }

      if (isHorizontal) {
        const x1 = valueScale(start);
        const x2 = valueScale(end);
        const maxPositive = positiveTotals[categoryIndex] ?? 0;
        const maxNegative = negativeTotals[categoryIndex] ?? 0;
        bars.push({
          capEnd: !isStacked || (rawValue >= 0
            ? Math.abs(end - maxPositive) < 1e-6
            : Math.abs(end - maxNegative) < 1e-6),
          capStart: !isStacked || Math.abs(start) < 1e-6,
          bottomScale,
          categoryIndex,
          color: colors.fill,
          depthOrder,
          depthOffsetX,
          depthOffsetY,
          height: Math.max(1, barThickness),
          isHorizontal: true,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: Math.min(x1, x2),
          depthX: usesSeriesDepthAxis ? depthBarX : frameOffsets?.depthX,
          depthY: usesSeriesDepthAxis ? depthBarY : frameOffsets?.depthY,
          shape3d: normalized3dShape,
          seriesIndex,
          stroke: colors.stroke,
          strokeWidth: colors.strokeWidth,
          topScale,
          top: categoryStart + barOffset,
          invertedNegative: isInvertedNegative,
          value: rawValue,
          width: Math.max(1, Math.abs(x2 - x1))
        });
      } else {
        const y1 = valueScale(start);
        const y2 = valueScale(end);
        const maxPositive = positiveTotals[categoryIndex] ?? 0;
        const maxNegative = negativeTotals[categoryIndex] ?? 0;
        bars.push({
          capEnd: !isStacked || (rawValue >= 0
            ? Math.abs(end - maxPositive) < 1e-6
            : Math.abs(end - maxNegative) < 1e-6),
          capStart: !isStacked || Math.abs(start) < 1e-6,
          bottomScale,
          categoryIndex,
          color: colors.fill,
          depthOrder,
          depthOffsetX,
          depthOffsetY,
          height: Math.max(1, Math.abs(y2 - y1)),
          isHorizontal: false,
          key: `bar-${seriesIndex}-${categoryIndex}`,
          left: categoryStart + barOffset,
          depthX: usesSeriesDepthAxis ? depthBarX : frameOffsets?.depthX,
          depthY: usesSeriesDepthAxis ? depthBarY : frameOffsets?.depthY,
          shape3d: normalized3dShape,
          seriesIndex,
          stroke: colors.stroke,
          strokeWidth: colors.strokeWidth,
          topScale,
          top: Math.min(y1, y2),
          invertedNegative: isInvertedNegative,
          value: rawValue,
          width: Math.max(1, barThickness)
        });
      }
    }
  }

  const renderedBars = chart.is3d && isStacked
    ? bars.slice().sort((left, right) => {
        if (left.categoryIndex !== right.categoryIndex) {
          return left.categoryIndex - right.categoryIndex;
        }
        const leftNegative = left.value < 0 ? 0 : 1;
        const rightNegative = right.value < 0 ? 0 : 1;
        if (leftNegative !== rightNegative) {
          return leftNegative - rightNegative;
        }
        if (isHorizontal) {
          if (left.left !== right.left) {
            return left.left - right.left;
          }
          return left.seriesIndex - right.seriesIndex;
        }
        if (left.top !== right.top) {
          return right.top - left.top;
        }
        return left.seriesIndex - right.seriesIndex;
      })
    : bars;
  const sortedBars = usesSeriesDepthAxis
    ? renderedBars.slice().sort((left, right) => {
        if ((left.depthOrder ?? 0) !== (right.depthOrder ?? 0)) {
          return (right.depthOrder ?? 0) - (left.depthOrder ?? 0);
        }
        if (left.categoryIndex !== right.categoryIndex) {
          return left.categoryIndex - right.categoryIndex;
        }
        return left.top - right.top;
      })
    : renderedBars;

  const useVertical3dGradient = chart.is3d && !isHorizontal;
  const gradientByColor = new Map<string, string>();
  if (useVertical3dGradient) {
    for (const bar of renderedBars) {
      if (!bar.color || bar.color === "none" || bar.color.startsWith("url(")) {
        continue;
      }
      if (!gradientByColor.has(bar.color)) {
        gradientByColor.set(
          bar.color,
          `bar3d-front-${chart.id}-${gradientByColor.size}`.replace(/[^A-Za-z0-9_-]/g, "-")
        );
      }
      bar.gradientId = gradientByColor.get(bar.color);
    }
  }

  const axisNode = renderCartesianAxes(
    chart,
    palette,
    plot,
    isHorizontal,
    categories,
    categoryPositions,
    ticks,
    (value) => valueScale(value),
    isPercentStacked ? formatPercentTickValue : formatTickValue
  );
  const depthAxisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const depthGridColor = lightenColor(depthAxisColor, 0.48);
  const frameDepthX = usesSeriesDepthAxis ? depthGridSpanX : (frameOffsets?.depthX ?? 0);
  const frameDepthY = usesSeriesDepthAxis ? depthGridSpanY : (frameOffsets?.depthY ?? 0);
  const depthAxisNode = usesSeriesDepthAxis && frameOffsets ? (
    <g>
      {Array.from({ length: categoryCount + 1 }, (_, boundaryIndex) => {
        const x = plot.left + ((boundaryIndex / Math.max(1, categoryCount)) * plot.width);
        return (
          <line
            key={`bar3d-floor-x-${boundaryIndex}`}
            stroke={depthGridColor}
            stroke-width={0.9}
            x1={x}
            x2={x + depthGridSpanX}
            y1={plot.top + plot.height}
            y2={plot.top + plot.height + depthGridSpanY}
          />
        );
      })}
      {Array.from({ length: seriesCount + 1 }, (_, boundaryIndex) => {
        const ratio = boundaryIndex / Math.max(1, seriesCount);
        const offsetX = depthGridSpanX * ratio;
        const offsetY = depthGridSpanY * ratio;
        return (
          <line
            key={`bar3d-floor-z-${boundaryIndex}`}
            stroke={depthGridColor}
            stroke-width={0.9}
            x1={plot.left + offsetX}
            x2={plot.left + plot.width + offsetX}
            y1={plot.top + plot.height + offsetY}
            y2={plot.top + plot.height + offsetY}
          />
        );
      })}
      {chart.series.map((series, seriesIndex) => {
        const ratio = (seriesIndex + 0.5) / Math.max(1, seriesCount);
        const x = plot.left + plot.width + depthGridSpanX * ratio;
        const y = plot.top + plot.height + depthGridSpanY * ratio;
        return (
          <text
            key={`bar3d-ser-label-${seriesIndex}`}
            fill={resolveChartAxisTextColor(chart)}
            font-size={10}
            text-anchor="start"
            x={x + 7}
            y={y + 3}
          >
            {series.name ?? `Series ${seriesIndex + 1}`}
          </text>
        );
      })}
    </g>
  ) : null;

  const frameNode = chart.is3d && frameOffsets ? (
    <g>
      <polygon
        fill={lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.05)}
        points={`${plot.left},${plot.top + plot.height} ${plot.left + plot.width},${plot.top + plot.height} ${plot.left + plot.width + frameDepthX},${plot.top + plot.height + frameDepthY} ${plot.left + frameDepthX},${plot.top + plot.height + frameDepthY}`}
        stroke={lightenColor(chart.axisLineColor ?? palette.border, 0.2)}
        stroke-width={1}
      />
      <polygon
        fill={lightenColor(chart.chartAreaFillColor ?? palette.surface, 0.12)}
        points={`${plot.left},${plot.top} ${plot.left + plot.width},${plot.top} ${plot.left + plot.width + frameDepthX},${plot.top + frameDepthY} ${plot.left + frameDepthX},${plot.top + frameDepthY}`}
        stroke={lightenColor(chart.axisLineColor ?? palette.border, 0.2)}
        stroke-width={1}
      />
    </g>
  ) : null;

  return (
    <g>
      {useVertical3dGradient && gradientByColor.size > 0 ? (
        <defs>
          {Array.from(gradientByColor.entries()).map(([color, id]) => (
            <linearGradient id={id} key={id} x1="0%" x2="0%" y1="0%" y2="100%">
              <stop offset="0%" stop-color={lightenColor(color, 0.28)} />
              <stop offset="42%" stop-color={lightenColor(color, 0.12)} />
              <stop offset="100%" stop-color={darkenColor(color, 0.1)} />
            </linearGradient>
          ))}
        </defs>
      ) : null}
      {axisNode}
      {frameNode}
      {depthAxisNode}
      {chart.is3d
        ? sortedBars.map((bar) => (
            <g key={`bar-hit-${bar.key}`} {...barChartElementDataProps(bar.seriesIndex, bar.categoryIndex)}>
              {renderExtrudedRect(bar)}
            </g>
          ))
        : renderedBars.map((bar) => (
            <rect
              {...barChartElementDataProps(bar.seriesIndex, bar.categoryIndex)}
              key={bar.key}
              fill={bar.color}
              height={bar.height}
              stroke={isHistogramLike ? "none" : bar.stroke}
              stroke-width={isHistogramLike ? 0 : bar.strokeWidth}
              width={bar.width}
              x={bar.left}
              y={bar.top}
            />
          ))}
      {renderedBars.map((bar) => (
        isSelectedChartPoint(selectedChartElement, chart.id, bar.seriesIndex, bar.categoryIndex)
        || (
          isSelectedChartSeries(selectedChartElement, chart.id, bar.seriesIndex)
          && selectedChartElement?.kind !== "point"
        )
          ? renderSelectionRectHandles(`bar-selection-${bar.key}`, bar.left, bar.top, bar.width, bar.height)
          : null
      ))}
    </g>
  );
}

export function renderWaterfallChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const stages = buildChartStages(chart);
  if (stages.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const labels = stages.map((stage) => stage.label);
  const xScale = scaleBand<string>()
    .domain(labels)
    .range([plot.left, plot.left + plot.width])
    .paddingInner(0.34)
    .paddingOuter(0.18);

  const bars: Array<ChartStage & { end: number; index: number; start: number }> = [];
  let runningTotal = 0;
  stages.forEach((stage, index) => {
    const start = stage.isSubtotal ? 0 : runningTotal;
    const end = stage.isSubtotal ? runningTotal : runningTotal + stage.value;
    if (!stage.isSubtotal) {
      runningTotal = end;
    }
    bars.push({ ...stage, end, index, start });
  });

  const extents = bars.flatMap((bar) => [0, bar.start, bar.end]);
  let minValue = Math.min(...extents);
  let maxValue = Math.max(...extents);
  if (maxValue <= minValue) {
    maxValue = minValue + 1;
  }

  const yScale = scaleLinear()
    .domain([minValue, maxValue])
    .range([plot.top + plot.height, plot.top]);
  const ticks = buildNumericTickValues(minValue, maxValue, chart.valueAxis?.majorUnit);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);

  return (
    <g>
      {ticks.map((tick) => (
        <g key={`waterfall-y-${tick}`}>
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
      <line stroke={axisColor} stroke-width={1.2} x1={plot.left} x2={plot.left} y1={plot.top} y2={plot.top + plot.height} />
      <line stroke={axisColor} stroke-width={1.2} x1={plot.left} x2={plot.left + plot.width} y1={yScale(0)} y2={yScale(0)} />
      {bars.map((bar, index) => {
        const bandLeft = xScale(bar.label) ?? plot.left;
        const bandWidth = xScale.bandwidth();
        const topValue = Math.max(bar.start, bar.end);
        const bottomValue = Math.min(bar.start, bar.end);
        const top = yScale(topValue);
        const height = Math.max(1, yScale(bottomValue) - top);
        const fill = bar.isSubtotal
          ? darkenColor(bar.color, 0.18)
          : bar.value >= 0
            ? bar.color
            : lightenColor(bar.color, 0.22);
        const connectorStart = index > 0 ? bars[index - 1] : null;
        const connectorY = connectorStart ? yScale(connectorStart.end) : 0;
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, 0, index);

        return (
          <g key={`waterfall-bar-${index}`}>
            {connectorStart ? (
              <line
                stroke={lightenColor(axisColor, 0.35)}
                stroke-dasharray="3 3"
                stroke-width={1}
                x1={(xScale(connectorStart.label) ?? plot.left) + xScale.bandwidth()}
                x2={bandLeft}
                y1={connectorY}
                y2={connectorY}
              />
            ) : null}
              <rect
                {...chartElementDataProps(0, index)}
                fill={fill}
                rx={2}
                ry={2}
                stroke={selected ? "#64748b" : darkenColor(fill, 0.22)}
                stroke-width={selected ? 2 : 1}
              x={bandLeft}
              y={top}
              width={bandWidth}
              height={height}
            />
            {selected
              ? renderSelectionRectHandles(`waterfall-selection-${index}`, bandLeft, top, bandWidth, height)
              : null}
            <text
              fill={labelColor}
              font-size={10}
              text-anchor="middle"
              x={bandLeft + bandWidth * 0.5}
              y={plot.top + plot.height + 14}
            >
              {bar.label}
            </text>
            <text
              fill={labelColor}
              font-size={10}
              text-anchor="middle"
              x={bandLeft + bandWidth * 0.5}
              y={top - 6}
            >
              {formatTickValue(bar.end)}
            </text>
          </g>
        );
      })}
    </g>
  );
}

export function renderFunnelChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const stages = buildChartStages(chart)
    .map((stage) => ({ ...stage, value: Math.max(0, stage.value) }))
    .filter((stage) => stage.value > 0);
  if (stages.length === 0) {
    return null;
  }

  const plot = layout.plot;
  const maxValue = Math.max(...stages.map((stage) => stage.value));
  const sectionHeight = plot.height / stages.length;
  const centerX = plot.left + plot.width * 0.5;
  const labelColor = resolveChartTextColor(chart);

  return (
    <g>
      {stages.map((stage, index) => {
        const stageWidth = (stage.value / maxValue) * plot.width;
        const topY = plot.top + index * sectionHeight;
        const stageHeight = Math.max(6, sectionHeight - 2);
        const left = centerX - stageWidth * 0.5;
        const fill = stage.isSubtotal ? darkenColor(stage.color, 0.14) : stage.color;
        const labelFitsInside = stageWidth > 90;
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, 0, index);

        return (
          <g key={`funnel-stage-${index}`}>
              <rect
                {...chartElementDataProps(0, index)}
                fill={fill}
                height={stageHeight}
              rx={0}
              ry={0}
                stroke={selected ? "#64748b" : darkenColor(fill, 0.2)}
                stroke-width={selected ? 2 : 1}
              width={stageWidth}
              x={left}
              y={topY}
            />
            {selected
              ? renderSelectionRectHandles(`funnel-selection-${index}`, left, topY, stageWidth, stageHeight)
              : null}
            <text
              fill={labelColor}
              font-size={10}
              text-anchor={labelFitsInside ? "middle" : "start"}
              x={labelFitsInside ? centerX : centerX + stageWidth * 0.5 + 8}
              y={topY + sectionHeight * 0.5}
            >
              {`${stage.label} ${formatTickValue(stage.value)}`}
            </text>
          </g>
        );
      })}
    </g>
  );
}

export function renderBoxWhiskerChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const visibleSeries = chart.series.filter((series) => series.hidden !== true);
  if (visibleSeries.length === 0) {
    return null;
  }

  const seriesStats = visibleSeries.map((series, index) => ({
    color: series.color ?? chartSeriesColor(chart, typeof series.formatIdx === "number" ? series.formatIdx : index),
    label: normalizeCategoryLabel(series.name) || `Series ${index + 1}`,
    lineColor: series.lineColor ?? series.color ?? chartSeriesColor(chart, typeof series.formatIdx === "number" ? series.formatIdx : index),
    series,
    stats: computeBoxWhiskerStats(series),
    visibility: resolveBoxWhiskerVisibility(series)
  })).filter((entry): entry is {
    color: string;
    label: string;
    lineColor: string;
    series: XlsxChartSeries;
    stats: BoxWhiskerStats;
    visibility: ReturnType<typeof resolveBoxWhiskerVisibility>;
  } => entry.stats != null);

  if (seriesStats.length === 0) {
    return null;
  }

  const allValues = seriesStats.flatMap((entry) => [
    entry.stats.min,
    entry.stats.lowerWhisker,
    entry.stats.q1,
    entry.stats.median,
    entry.stats.q3,
    entry.stats.upperWhisker,
    entry.stats.max
  ]);
  const minValue = Math.min(...allValues);
  const maxValue = Math.max(...allValues);
  const valueDomain = resolveAxisDomainWithChartOverrides(chart.valueAxis, minValue, maxValue, false);
  const yScale = scaleLinear()
    .domain([valueDomain.min, valueDomain.max])
    .range([layout.plot.top + layout.plot.height, layout.plot.top]);
  const xScale = scalePoint<number>()
    .domain(seriesStats.map((_, index) => index))
    .range([layout.plot.left + 24, layout.plot.left + layout.plot.width - 24]);
  const gap = seriesStats.length > 1
    ? Math.abs((xScale(1) ?? 0) - (xScale(0) ?? 0))
    : layout.plot.width * 0.5;
  const boxWidth = clamp(gap * 0.34, 20, 54);
  const capWidth = Math.max(8, boxWidth * 0.52);
  const axisColor = chart.axisLineColor ?? chart.chartAreaBorderColor ?? palette.border;
  const labelColor = resolveChartAxisTextColor(chart);
  const meanLinePoints = seriesStats
    .filter((entry) => entry.visibility.meanLine)
    .map((entry, index) => `${xScale(index) ?? layout.plot.left},${yScale(entry.stats.mean)}`);

  return (
    <g>
      {valueDomain.ticks.map((tick, index) => {
        const y = yScale(tick);
        return (
          <g key={`box-whisker-grid-${index}`}>
            <line
              stroke={lightenColor(axisColor, 0.22)}
              stroke-width={0.8}
              x1={layout.plot.left}
              x2={layout.plot.left + layout.plot.width}
              y1={y}
              y2={y}
            />
            <text
              fill={labelColor}
              font-size={10}
              text-anchor="end"
              x={layout.plot.left - 6}
              y={y + 3}
            >
              {formatTickValue(tick)}
            </text>
          </g>
        );
      })}
      <line
        stroke={axisColor}
        stroke-width={1}
        x1={layout.plot.left}
        x2={layout.plot.left}
        y1={layout.plot.top}
        y2={layout.plot.top + layout.plot.height}
      />
      <line
        stroke={axisColor}
        stroke-width={1}
        x1={layout.plot.left}
        x2={layout.plot.left + layout.plot.width}
        y1={layout.plot.top + layout.plot.height}
        y2={layout.plot.top + layout.plot.height}
      />
      {meanLinePoints.length >= 2 ? (
        <polyline
          fill="none"
          points={meanLinePoints.join(" ")}
          stroke={darkenColor(axisColor, 0.2)}
          stroke-dasharray="4 3"
          stroke-width={1}
        />
      ) : null}
      {seriesStats.map((entry, index) => {
        const x = xScale(index) ?? layout.plot.left;
        const boxTop = yScale(entry.stats.q3);
        const boxBottom = yScale(entry.stats.q1);
        const medianY = yScale(entry.stats.median);
        const lowerWhiskerY = yScale(entry.stats.lowerWhisker);
        const upperWhiskerY = yScale(entry.stats.upperWhisker);
        const meanY = yScale(entry.stats.mean);
        const visiblePoints = entry.visibility.nonoutliers
          ? entry.stats.visiblePoints
          : [];
        const outliers = entry.visibility.outliers
          ? entry.stats.outliers
          : [];
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, index, index);

        return (
            <g {...chartElementDataProps(index)} key={`box-whisker-series-${index}`}>
            <line
              stroke={entry.lineColor}
              stroke-width={1.5}
              x1={x}
              x2={x}
              y1={upperWhiskerY}
              y2={boxTop}
            />
            <line
              stroke={entry.lineColor}
              stroke-width={1.5}
              x1={x}
              x2={x}
              y1={boxBottom}
              y2={lowerWhiskerY}
            />
            <line
              stroke={entry.lineColor}
              stroke-width={1.5}
              x1={x - capWidth * 0.5}
              x2={x + capWidth * 0.5}
              y1={upperWhiskerY}
              y2={upperWhiskerY}
            />
            <line
              stroke={entry.lineColor}
              stroke-width={1.5}
              x1={x - capWidth * 0.5}
              x2={x + capWidth * 0.5}
              y1={lowerWhiskerY}
              y2={lowerWhiskerY}
            />
              <rect
                {...chartElementDataProps(index, index)}
                fill={lightenColor(entry.color, 0.35)}
              fill-opacity={0.72}
              height={Math.max(1, boxBottom - boxTop)}
                stroke={selected ? "#64748b" : entry.lineColor}
                stroke-width={selected ? 2 : 1.5}
              width={boxWidth}
              x={x - boxWidth * 0.5}
              y={boxTop}
            />
            {selected
              ? renderSelectionRectHandles(
                  `box-whisker-selection-${index}`,
                  x - boxWidth * 0.5,
                  boxTop,
                  boxWidth,
                  Math.max(1, boxBottom - boxTop)
                )
              : null}
            <line
              stroke={darkenColor(entry.lineColor, 0.15)}
              stroke-width={2}
              x1={x - boxWidth * 0.5}
              x2={x + boxWidth * 0.5}
              y1={medianY}
              y2={medianY}
            />
            {entry.visibility.meanMarker ? (
              <>
                <line
                  stroke={darkenColor(entry.lineColor, 0.18)}
                  stroke-width={1.2}
                  x1={x - 4}
                  x2={x + 4}
                  y1={meanY - 4}
                  y2={meanY + 4}
                />
                <line
                  stroke={darkenColor(entry.lineColor, 0.18)}
                  stroke-width={1.2}
                  x1={x - 4}
                  x2={x + 4}
                  y1={meanY + 4}
                  y2={meanY - 4}
                />
              </>
            ) : null}
            {visiblePoints.map((value, pointIndex) => {
              const y = yScale(value);
              const jitter = ((pointIndex % 7) - 3) * (boxWidth / 16);
              return (
                  <circle
                    {...chartElementDataProps(index, pointIndex)}
                    key={`box-whisker-visible-${index}-${pointIndex}`}
                  cx={x + jitter}
                  cy={y}
                  fill={entry.color}
                  fill-opacity={0.45}
                  r={2}
                  stroke="none"
                />
              );
            })}
            {outliers.map((value, pointIndex) => {
              const y = yScale(value);
              return (
                  <circle
                    {...chartElementDataProps(index, pointIndex)}
                    key={`box-whisker-outlier-${index}-${pointIndex}`}
                  cx={x}
                  cy={y}
                  fill="#ffffff"
                  r={3}
                  stroke={entry.lineColor}
                  stroke-width={1.2}
                />
              );
            })}
            <text
              fill={labelColor}
              font-size={10}
              text-anchor="middle"
              x={x}
              y={layout.plot.top + layout.plot.height + 14}
            >
              {entry.label}
            </text>
          </g>
        );
      })}
    </g>
  );
}
