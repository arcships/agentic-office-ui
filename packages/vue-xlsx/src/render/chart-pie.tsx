import { hierarchy as d3Hierarchy, partition as d3Partition, treemap as d3Treemap, treemapBinary, treemapDice, treemapSquarify } from "d3-hierarchy";
import { scaleBand, scaleLinear } from "d3-scale";
import { arc as d3Arc, pie as d3Pie } from "d3-shape";
import type { HierarchyNode, HierarchyRectangularNode } from "d3-hierarchy";
import type { XlsxChart, XlsxChartElementSelection, XlsxChartSeries, XlsxImageRect } from "@arcships/xlsx-core";
import { chartSeriesColor, chartSeriesStrokeColor, chartPointColor, safeNumber, clamp, darkenColor, lightenColor, coerceLooseNumber, resolveRenderableSeriesValue, resolveChartTextColor, resolveChartAxisTextColor, resolveChartMutedTextColor, buildChartFontFamily, formatTickValue, formatPercentTickValue, chartElementDataProps, barChartElementDataProps, resolveChartElementTarget, resolveChartSelectionFromTarget, isSelectedChartSeries, isSelectedChartPoint, isSelectedChartPointOrSeries, renderSelectionRectHandles, renderSelectionPointHandles, normalizeRendererHexColor, toSvgNumber, normalizeChartMarkerSymbol, markerSymbolPath, normalizeLegendPosition, parseRgbColor, mixRgbColor, escapeCssFontFamilyToken, DEFAULT_CHART_FONT_STACK, DEFAULT_CHART_TEXT_COLOR, DEFAULT_CHART_MUTED_TEXT_COLOR, truncateSvgText, getCategoryLabels, normalizeCategoryLabel, estimateReferencePointCount, isHistogramLikeSeries, isHistogramLikeChart, resolve3dFrameOffsets, resolveAxisDomainWithChartOverrides, resolveCategoryBandPadding, chartSeriesBarColors, buildNumericTickValues, buildNiceStep, resolveNumericAxisDomain, renderCartesianAxes, renderExtrudedRect, buildChartStages, resolveChartStageSubtotal, computeBoxWhiskerStats, resolveBoxWhiskerVisibility, resolveBoxWhiskerQuartileMethod, computePercentile, renderLineOrAreaChart3d, projectCartesian3dPoint, scaleProjectedVector, buildLinearSvgPath, buildRibbonSvgPath, renderRadialFrustum, sampleProjectedEllipseArc, isComboChart, getComboLegendSeries, findAxisForGroup, buildComboGroups, resolveStockRoleIndices, resolveStockPalette, normalizeBuiltinPieStyleId, getBuiltinPiePalette, resolvePiePointColor, selectPrimaryPieSeriesIndex, buildPieEntries, buildHierarchyData, resolveHierarchyNodeColor, resolveTreemapNodeColor, excelTreemapTile, normalizePieArc, resolvePieFrontSegments, pieEllipsePoint, buildPieOuterWallPath, isPieFrontFacingAngle, buildPieRadialWallPath, TWO_PI, normalizeRegionMapKey, resolveRegionMapFeature, resolveRegionMapBaseColor, resolveRegionMapDataColor, resolveRegionMapValueColors, resolveRegionMapColorStrings, resolveRegionMapLayoutProperties, resolveRegionMapFeatureSet, getRegionMapBaseFeatures, resolveRegionMapValueColorFromStops, resolveRegionMapValueColor, resolveRegionMapNoDataColor, resolveSurfaceBaseColor, normalizeBuiltinSurfaceStyleId, hasExplicitSurfaceBaseColor, buildMonochromeSurfacePalette, getBuiltinSurfacePalette, shouldPreferBuiltinSurfacePalette, getSurfaceBandCount, getSurfaceColorStops, resolveSurfaceBandPaletteColor, getSurfaceDomain, resolveSurfaceColor, resolveSurfaceBandColor, resolveSurfaceBandIndex, buildSurfaceSmoothPath, splitSurfacePointRuns, connectSurfaceContourSegments, getSurfaceWireframePalette, resolveSurfaceWireframeColor, resolveSurfacePlotRect, buildSurfaceLegendItems, isContourSurfaceChart, getLegendItems, buildLayout, normalizeRenderableChartType, renderSurfaceAxes, renderTitle, renderLegend, buildRegionMapLegendItems, buildStockLegendItems, excelSerialToDate, isLikelyDateFormatCode, formatExcelDateSerial, formatCategoryLabel } from "./chart-renderer";
import type { ChartRendererPalette, ChartSvgProps, ChartLayout, LegendItem, BarRect, PlotRect, ChartStage, ChartHierarchyDatum, ComboRenderableGroup, BoxWhiskerStats, SurfaceDomain, RegionMapFeature, SurfacePathPoint, SurfaceContourSegment, ChartElementDataOptions } from "./chart-renderer";


export function renderPieChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, chartType: string, selectedChartElement?: XlsxChartElementSelection | null) {
  const pieSeriesIndex = selectPrimaryPieSeriesIndex(chart);
  const pieSeries = chart.series[pieSeriesIndex];
  const pieData = buildPieEntries(chart, pieSeriesIndex);
  if (pieData.length === 0) {
    return null;
  }

  const legendOnRight = layout.legendPosition === "right";
  const centerX = legendOnRight ? layout.plot.left + layout.plot.width * 0.42 : layout.plot.left + layout.plot.width * 0.5;
  const centerY = layout.plot.top + layout.plot.height * 0.54;
  const outerRadius = Math.max(16, Math.min(layout.plot.width, layout.plot.height) * (chartType === "Doughnut" ? 0.32 : 0.38));
  const innerRatio = chartType === "Doughnut" ? clamp((chart.holeSize ?? 56) / 100, 0.1, 0.9) : 0;
  const innerRadius = outerRadius * innerRatio;
  const startAngle = ((chart.firstSliceAngle ?? 0) * Math.PI) / 180;
  const arcs = d3Pie<{ color: string; index: number; label: string; value: number }>()
    .value((entry) => entry.value)
    .sort(null)
    .startAngle(startAngle)
    .endAngle(startAngle + Math.PI * 2)(pieData);

  const arcPath = d3Arc<typeof arcs[number]>()
    .innerRadius(innerRadius)
    .outerRadius(outerRadius);

  const isPie3d = chartType === "Pie3D" || (chartType === "PieExploded" && chart.is3d === true);
  const rotX = chart.view3d?.rotX ?? 20;
  const perspective = chart.view3d?.perspective ?? 30;
  const tiltFromRotX = clamp(0.14 + (rotX / 90) * 0.72, 0.12, 0.92);
  const perspectiveCompression = clamp(1 - (perspective / 180), 0.55, 1);
  const tilt = isPie3d ? clamp(tiltFromRotX * perspectiveCompression, 0.14, 0.78) : 1;
  const depthScale = (chart.view3d?.depthPercent ?? 100) / 100;
  const depth = isPie3d ? Math.max(9, outerRadius * 0.28 * depthScale) : 0;
  const dataLabelsEnabled = Boolean(chart.dataLabels?.showCategoryName || chart.dataLabels?.showPercent || chart.dataLabels?.showValue);
  const total = pieData.reduce((sum, entry) => sum + entry.value, 0);
  const shadowId = `pie3d-shadow-${chart.id}`.replace(/[^A-Za-z0-9_-]/g, "-");
  const baseShadowId = `pie3d-base-shadow-${chart.id}`.replace(/[^A-Za-z0-9_-]/g, "-");
  const sliceSeparatorColor = chart.chartAreaFillColor ?? "#ffffff";
  const labelBounds = {
    bottom: layout.height - 6,
    left: 6,
    right: layout.width - 6,
    top: layout.plot.top + 8
  };
  const pointLabelByIndex = new Map((chart.dataLabels?.pointLabels ?? []).map((label) => [label.index, label]));
  const centerValuePointLabel = chartType === "Doughnut"
    ? (chart.dataLabels?.pointLabels ?? []).find((label) => (
      label.deleted !== true
      && label.showValue === true
      && label.showCategoryName !== true
      && label.showPercent !== true
      && label.showSeriesName !== true
      && label.showBubbleSize !== true
    )) ?? null
    : null;
  const centerValueEntry = centerValuePointLabel
    ? pieData.find((entry) => entry.index === centerValuePointLabel.index) ?? null
    : null;
  const shouldRenderCenterValue = chartType === "Doughnut" && centerValueEntry != null;
  const centerValueLooksPercent = centerValueEntry != null
    && total > 0
    && total <= 1.0000001
    && centerValueEntry.value >= 0
    && centerValueEntry.value <= 1.0000001;
  const centerValueText = centerValueEntry == null
    ? ""
    : centerValueLooksPercent
      ? `${Math.round(centerValueEntry.value * 100)}%`
      : formatTickValue(centerValueEntry.value);
  const centerValueFontSize = shouldRenderCenterValue
    ? Math.max(
        14,
        Math.min(
          Math.max(20, innerRadius * 0.72),
          (centerValuePointLabel?.fontSizePt ?? 28) * 0.85
        )
      )
    : 0;

  const resolveSliceExplosion = (pointIndex: number) => {
    const pointStyle = pieSeries?.dataPointStyles?.find((entry) => entry.index === pointIndex);
    const seriesExplosion = typeof pieSeries?.shapeProperties?.xmlExplosion === "number"
      ? pieSeries.shapeProperties.xmlExplosion
      : 0;
    const rawExplosion = Math.max(0, pointStyle?.explosion ?? seriesExplosion ?? 0);
    if (rawExplosion <= 0) {
      return 0;
    }
    // OOXML explosion values are percent-like offsets of the pie radius.
    return outerRadius * clamp(rawExplosion / 100, 0, 4);
  };
  const hasExplodedSlices = arcs.some((arc) => resolveSliceExplosion(arc.data.index) > 0);
  const sliceSeparatorWidth = hasExplodedSlices ? 2 : 1.2;

  return (
    <g>
      {isPie3d ? (
        <defs>
          <filter id={shadowId} x="-40%" y="-40%" width="180%" height="200%">
            <feDropShadow dx="1.2" dy="3.6" flood-color="#000000" flood-opacity="0.28" std-deviation="2.8" />
          </filter>
          <filter id={baseShadowId} x="-50%" y="-50%" width="220%" height="220%">
            <feGaussianBlur std-deviation="3.2" />
          </filter>
        </defs>
      ) : null}
      {isPie3d ? (
        <ellipse
          cx={centerX}
          cy={centerY + depth + 3}
          fill="#000000"
          filter={`url(#${baseShadowId})`}
          opacity={0.14}
          rx={outerRadius * 1.02}
          ry={outerRadius * tilt * 0.68}
        />
      ) : null}
      {isPie3d
          ? arcs.map((arc) => {
              const midAngle = (arc.startAngle + arc.endAngle) / 2;
              const explosion = resolveSliceExplosion(arc.data.index);
              const explodeX = Math.sin(midAngle) * explosion;
              const explodeY = -Math.cos(midAngle) * explosion * tilt;
              const sidePaths = buildPieOuterWallPath(
                centerX,
                centerY,
                outerRadius,
                tilt,
                depth,
                arc.startAngle,
                arc.endAngle,
                explosion > 0
              );
            if (sidePaths.length === 0) {
              return null;
            }
            return (
              <g key={`pie-side-${arc.data.index}`} transform={`translate(${explodeX}, ${explodeY})`}>
                {sidePaths.map((sidePath, sideIndex) => (
                  <path
                    d={sidePath}
                    fill={darkenColor(arc.data.color, 0.34)}
                    key={`pie-side-path-${arc.data.index}-${sideIndex}`}
                    stroke={darkenColor(arc.data.color, 0.5)}
                    stroke-width={0.8}
                  />
                ))}
                {explosion > 0 ? (() => {
                  const startWall = buildPieRadialWallPath(
                    centerX,
                    centerY,
                    outerRadius,
                    tilt,
                    depth,
                    arc.startAngle,
                    true
                  );
                  const endWall = buildPieRadialWallPath(
                    centerX,
                    centerY,
                    outerRadius,
                    tilt,
                    depth,
                    arc.endAngle,
                    true
                  );
                  return (
                    <>
                      {startWall ? (
                        <path
                          d={startWall}
                          fill={darkenColor(arc.data.color, 0.26)}
                          stroke={darkenColor(arc.data.color, 0.44)}
                          stroke-width={0.8}
                        />
                      ) : null}
                      {endWall ? (
                        <path
                          d={endWall}
                          fill={darkenColor(arc.data.color, 0.2)}
                          stroke={darkenColor(arc.data.color, 0.4)}
                          stroke-width={0.8}
                        />
                      ) : null}
                    </>
                  );
                })() : null}
              </g>
            );
          })
        : null}
      {arcs.map((arc) => {
        const explosion = resolveSliceExplosion(arc.data.index);
        const midAngle = (arc.startAngle + arc.endAngle) / 2;
        const explodeX = Math.sin(midAngle) * explosion;
        const explodeY = -Math.cos(midAngle) * explosion * (isPie3d ? tilt : 1);
        const isSliceSelected = isSelectedChartPoint(selectedChartElement, chart.id, pieSeriesIndex, arc.data.index)
          || (
            isSelectedChartSeries(selectedChartElement, chart.id, pieSeriesIndex)
            && selectedChartElement?.kind !== "point"
          );
        const sliceHandlePoints = isSliceSelected
          ? [
              pieEllipsePoint(centerX, centerY, outerRadius, isPie3d ? tilt : 1, arc.startAngle),
              pieEllipsePoint(centerX, centerY, outerRadius, isPie3d ? tilt : 1, midAngle),
              pieEllipsePoint(centerX, centerY, outerRadius, isPie3d ? tilt : 1, arc.endAngle),
              innerRadius > 0
                ? pieEllipsePoint(centerX, centerY, innerRadius, isPie3d ? tilt : 1, midAngle)
                : { x: centerX, y: centerY }
            ].map((point) => ({
              x: point.x + explodeX,
              y: point.y + explodeY
            }))
          : [];
        const labelRadius = outerRadius + (chartType === "PieExploded" ? 8 : 12);
        const labelX = centerX + Math.sin(midAngle) * labelRadius + explodeX;
        const labelY = centerY - Math.cos(midAngle) * labelRadius * (isPie3d ? tilt : 1) + explodeY;
        const pieces: string[] = [];
        if (chart.dataLabels?.showCategoryName && arc.data.label.trim().length > 0) {
          pieces.push(arc.data.label);
        }
        const pointLabel = pointLabelByIndex.get(arc.data.index);
        const showValue = pointLabel?.showValue ?? chart.dataLabels?.showValue;
        const showPercent = pointLabel?.showPercent ?? chart.dataLabels?.showPercent;
        if (showValue) {
          pieces.push(formatTickValue(arc.data.value));
        }
        if (showPercent) {
          pieces.push(`${Math.round((arc.data.value / Math.max(1, total)) * 100)}%`);
        }
        const labelText = pieces.join(", ");
        const truncatedLabelText = truncateSvgText(labelText, Math.max(48, layout.width * 0.42), 10);
        const approxLabelWidth = Math.max(12, truncatedLabelText.length * 5.6);
        let labelAnchor: "end" | "start" = labelX >= centerX ? "start" : "end";
        let resolvedLabelX = labelX;
        if (labelAnchor === "start" && resolvedLabelX + approxLabelWidth > labelBounds.right) {
          labelAnchor = "end";
        }
        if (labelAnchor === "end" && resolvedLabelX - approxLabelWidth < labelBounds.left) {
          labelAnchor = "start";
        }
        resolvedLabelX = clamp(resolvedLabelX, labelBounds.left, labelBounds.right);
        const resolvedLabelY = clamp(labelY, labelBounds.top, labelBounds.bottom);
        return (
          <>
            <g transform={`translate(${explodeX}, ${explodeY})`}>
              <path
                {...chartElementDataProps(pieSeriesIndex, arc.data.index)}
                d={(arcPath(arc) ?? "")}
                fill={arc.data.color}
                filter={isPie3d ? `url(#${shadowId})` : undefined}
                stroke={isSliceSelected ? "#64748b" : sliceSeparatorColor}
                stroke-width={isSliceSelected ? Math.max(2, sliceSeparatorWidth) : sliceSeparatorWidth}
                transform={`translate(${centerX}, ${centerY})${isPie3d ? ` scale(1, ${tilt})` : ""}`}
              />
            </g>
            {isSliceSelected
              ? renderSelectionPointHandles(`pie-selection-${arc.data.index}`, sliceHandlePoints)
              : null}
            {dataLabelsEnabled && truncatedLabelText.length > 0 ? (
              <text
                fill={resolveChartTextColor(chart)}
                font-size={10}
                text-anchor={labelAnchor}
                x={resolvedLabelX}
                y={resolvedLabelY}
              >
                {truncatedLabelText}
              </text>
            ) : null}
          </>
        );
      })}
      {shouldRenderCenterValue ? (
        <text
          fill={chart.textColor ?? chart.titleColor ?? DEFAULT_CHART_TEXT_COLOR}
          font-size={centerValueFontSize}
          font-weight={700}
          text-anchor="middle"
          x={centerX}
          y={centerY + centerValueFontSize * 0.34}
        >
          {centerValueText}
        </text>
      ) : null}
    </g>
  );
}

export function renderBarOfPieChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const pieSeriesIndex = selectPrimaryPieSeriesIndex(chart);
  const pieSeries = chart.series[pieSeriesIndex];
  const categories = getCategoryLabels(chart);
  const values = pieSeries?.values.map((value) => Math.max(0, safeNumber(value) ?? 0)) ?? [];
  if (values.length === 0) {
    return null;
  }
  const raw = (chart.raw ?? {}) as Record<string, unknown>;
  const ofPieType = raw.ofPieType === "pie" ? "pie" : "bar";
  const splitPos = typeof raw.splitPos === "number" ? raw.splitPos : 0;
  let secondaryIndices = values
    .map((value, index) => ({ index, value }))
    .filter(({ value }) => value <= splitPos)
    .map(({ index }) => index);
  if (secondaryIndices.length === 0) {
    secondaryIndices = values
      .map((value, index) => ({ index, value }))
      .sort((left, right) => left.value - right.value)
      .slice(0, Math.min(2, values.length))
      .map(({ index }) => index);
  }
  const secondarySet = new Set(secondaryIndices);
  const secondaryTotal = secondaryIndices.reduce((sum, index) => sum + (values[index] ?? 0), 0);
  const primaryData = values.flatMap((value, index) => secondarySet.has(index)
    ? []
    : [{ color: chartPointColor(chart, index, pieSeriesIndex), index, label: categories[index], value }]);
  if (secondaryTotal > 0) {
    primaryData.push({
      color: chartPointColor(chart, secondaryIndices[0] ?? 0, pieSeriesIndex),
      index: secondaryIndices[0] ?? 0,
      label: "Other",
      value: secondaryTotal
    });
  }

  const pieCenterX = layout.plot.left + layout.plot.width * 0.28;
  const pieCenterY = layout.plot.top + layout.plot.height * 0.55;
  const pieRadius = Math.max(16, Math.min(layout.plot.height, layout.plot.width * 0.45) * 0.3);
  const arc = d3Arc<{ endAngle: number; startAngle: number }>().innerRadius(0).outerRadius(pieRadius);
  const pieArcs = d3Pie<{ color: string; index: number; label: string; value: number }>()
    .value((entry) => entry.value)
    .sort(null)
    .startAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180)
    .endAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180 + Math.PI * 2)(primaryData);

  const secondaryLabels = secondaryIndices.map((index) => categories[index] ?? "");
  const secondaryValues = secondaryIndices.map((index) => values[index] ?? 0);
  const connectorTargetX = layout.plot.left + layout.plot.width * 0.69;
  const secondaryData = secondaryIndices.map((index) => ({
    color: chartPointColor(chart, index, pieSeriesIndex),
    index,
    label: categories[index] ?? "",
    value: values[index] ?? 0
  }));
  const secondaryCenterX = layout.plot.left + layout.plot.width * 0.79;
  const secondaryCenterY = pieCenterY;
  const secondaryRadius = pieRadius * clamp(((typeof raw.secondPieSize === "number" ? raw.secondPieSize : 100) / 100), 0.55, 1.5);
  const secondaryArc = d3Arc<{ endAngle: number; startAngle: number }>().innerRadius(0).outerRadius(secondaryRadius);
  const secondaryPieArcs = d3Pie<{ color: string; index: number; label: string; value: number }>()
    .value((entry) => entry.value)
    .sort(null)
    .startAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180)
    .endAngle(((90 - (chart.firstSliceAngle ?? 0)) * Math.PI) / 180 + Math.PI * 2)(secondaryData);
  const stackedBarLeft = layout.plot.left + layout.plot.width * 0.72;
  const stackedBarWidth = Math.max(20, layout.plot.width * 0.13);
  const stackedBarTop = layout.plot.top + 16;
  const stackedBarHeight = Math.max(28, layout.plot.height - 32);
  const secondaryTotalSafe = Math.max(1e-6, secondaryTotal);
  let stackCursor = stackedBarTop;

  return (
    <g>
        {pieArcs.map((entry, index) => {
          const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, pieSeriesIndex, entry.data.index);
          const midAngle = (entry.startAngle + entry.endAngle) / 2;
          const dot = pieEllipsePoint(pieCenterX, pieCenterY, pieRadius * 0.62, 1, midAngle);
          return (
            <>
              <path
                {...chartElementDataProps(pieSeriesIndex, entry.data.index)}
                d={arc(entry) ?? ""}
                fill={entry.data.color}
                stroke={selected ? "#64748b" : chart.chartAreaFillColor ?? palette.surface}
                stroke-width={selected ? 2 : 1}
                transform={`translate(${pieCenterX}, ${pieCenterY})`}
              />
              {selected
                ? renderSelectionPointHandles(`bar-of-pie-main-selection-${index}`, [dot])
                : null}
            </>
          );
        })}
      <line
        stroke={chart.chartAreaBorderColor ?? palette.border}
        stroke-width={1}
        x1={pieCenterX + pieRadius}
        x2={connectorTargetX}
        y1={pieCenterY - pieRadius * 0.4}
        y2={layout.plot.top + 10}
      />
      <line
        stroke={chart.chartAreaBorderColor ?? palette.border}
        stroke-width={1}
        x1={pieCenterX + pieRadius}
        x2={connectorTargetX}
        y1={pieCenterY + pieRadius * 0.4}
        y2={layout.plot.top + layout.plot.height - 10}
      />
      {ofPieType === "pie"
          ? secondaryPieArcs.map((entry, index) => {
            const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, pieSeriesIndex, entry.data.index);
            const midAngle = (entry.startAngle + entry.endAngle) / 2;
            const dot = pieEllipsePoint(secondaryCenterX, secondaryCenterY, secondaryRadius * 0.62, 1, midAngle);
            return (
              <>
                <path
                  {...chartElementDataProps(pieSeriesIndex, entry.data.index)}
                  d={secondaryArc(entry) ?? ""}
                  fill={entry.data.color}
                  stroke={selected ? "#64748b" : chart.chartAreaFillColor ?? palette.surface}
                  stroke-width={selected ? 2 : 1}
                  transform={`translate(${secondaryCenterX}, ${secondaryCenterY})`}
                />
                {selected
                  ? renderSelectionPointHandles(`bar-of-pie-secondary-pie-selection-${index}`, [dot])
                  : null}
              </>
            );
          })
        : secondaryData.map((entry, index) => {
          const segmentHeight = index === secondaryData.length - 1
            ? Math.max(1, stackedBarTop + stackedBarHeight - stackCursor)
            : Math.max(1, (entry.value / secondaryTotalSafe) * stackedBarHeight);
          const y = stackCursor;
          stackCursor += segmentHeight;
          const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, pieSeriesIndex, entry.index);
          return (
            <g key={`bar-of-pie-secondary-bar-${index}`}>
                <rect
                  {...chartElementDataProps(pieSeriesIndex, entry.index)}
                  fill={entry.color}
                  height={segmentHeight}
                  stroke={selected ? "#64748b" : chart.chartAreaFillColor ?? palette.surface}
                  stroke-width={selected ? 2 : 1}
                width={stackedBarWidth}
                x={stackedBarLeft}
                y={y}
              />
              {selected
                ? renderSelectionRectHandles(
                    `bar-of-pie-secondary-bar-selection-${index}`,
                    stackedBarLeft,
                    y,
                    stackedBarWidth,
                    segmentHeight
                  )
                : null}
              <text
                fill={resolveChartAxisTextColor(chart)}
                font-size={10}
                text-anchor="start"
                x={stackedBarLeft + stackedBarWidth + 6}
                y={y + segmentHeight * 0.5 + 3}
              >
                {entry.label}
              </text>
            </g>
          );
        })}
    </g>
  );
}

export function renderSunburstChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const hierarchyData = buildHierarchyData(chart);
  if (!hierarchyData) {
    return null;
  }

  const root = d3Hierarchy(hierarchyData)
    .sum((node) => node.children && node.children.length > 0 ? 0 : Math.max(0.0001, node.value ?? 0))
    .sort((left, right) => (right.value ?? 0) - (left.value ?? 0));
  const partitioned = d3Partition<ChartHierarchyDatum>()
    .size([Math.PI * 2, root.height + 1])(root);
  const plot = layout.plot;
  const radius = Math.max(24, Math.min(plot.width, plot.height) * 0.5);
  const holeRadius = radius * 0.16;
  const centerX = plot.left + plot.width * 0.5;
  const centerY = plot.top + plot.height * 0.5;
  const ringSpan = Math.max(1, partitioned.height || 1);
  const arcBuilder = d3Arc<HierarchyRectangularNode<ChartHierarchyDatum>>()
    .startAngle((node) => node.x0)
    .endAngle((node) => node.x1)
    .padAngle(0.005)
    .padRadius(radius)
    .innerRadius((node) => holeRadius + ((node.y0 - 1) / ringSpan) * (radius - holeRadius))
    .outerRadius((node) => holeRadius + ((node.y1 - 1) / ringSpan) * (radius - holeRadius) - 1);

  return (
    <g transform={`translate(${centerX}, ${centerY})`}>
      {partitioned.descendants().filter((node) => node.depth > 0).map((node, index) => {
        const path = arcBuilder(node);
        if (!path) {
          return null;
        }
        const labelAngle = (node.x0 + node.x1) * 0.5 - Math.PI * 0.5;
        const labelRadius = holeRadius + (((node.y0 + node.y1) * 0.5 - 1) / ringSpan) * (radius - holeRadius);
        const labelX = Math.cos(labelAngle) * labelRadius;
        const labelY = Math.sin(labelAngle) * labelRadius;
        const arcSpan = node.x1 - node.x0;
        const canShowLabel = arcSpan * labelRadius > 26;
        const fill = resolveHierarchyNodeColor(chart, node);
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, 0, index);

        return (
          <g key={`sunburst-node-${index}`}>
              <path
                {...chartElementDataProps(0, index)}
                d={path}
                fill={fill}
                stroke={selected ? "#64748b" : palette.surface}
                stroke-width={selected ? 2 : 1}
            />
            {selected
              ? renderSelectionPointHandles(`sunburst-selection-${index}`, [{ x: labelX, y: labelY }])
              : null}
            {canShowLabel ? (
              <text
                fill={darkenColor(fill, 0.65)}
                font-size={9}
                text-anchor="middle"
                transform={`translate(${labelX}, ${labelY}) rotate(${(labelAngle * 180) / Math.PI})`}
              >
                {node.data.name}
              </text>
            ) : null}
          </g>
        );
      })}
      <circle fill={chart.chartAreaFillColor ?? palette.surface} r={holeRadius - 2} />
    </g>
  );
}

export function renderTreemapChart(chart: XlsxChart, palette: ChartRendererPalette, layout: ChartLayout, selectedChartElement?: XlsxChartElementSelection | null) {
  const hierarchyData = buildHierarchyData(chart);
  if (!hierarchyData) {
    return null;
  }

  const root = d3Hierarchy(hierarchyData)
    .sum((node) => node.children && node.children.length > 0 ? 0 : Math.max(0.0001, node.value ?? 0))
    .sort((left, right) => (right.value ?? 0) - (left.value ?? 0));
  const treemapRoot = d3Treemap<ChartHierarchyDatum>()
    .size([layout.plot.width, layout.plot.height])
    .paddingInner(2)
    .paddingOuter(1)
    .round(true)
    .tile(excelTreemapTile)(root);

  return (
    <g transform={`translate(${layout.plot.left}, ${layout.plot.top})`}>
      {treemapRoot.leaves().map((leaf, index) => {
        const fill = resolveTreemapNodeColor(chart, leaf);
        const width = Math.max(0, leaf.x1 - leaf.x0);
        const height = Math.max(0, leaf.y1 - leaf.y0);
        const canShowLabel = width > 48 && height > 22;
        const selected = isSelectedChartPointOrSeries(selectedChartElement, chart.id, 0, index);

        return (
          <g key={`treemap-leaf-${index}`}>
              <rect
                {...chartElementDataProps(0, index)}
                fill={fill}
                rx={3}
                ry={3}
                stroke={selected ? "#64748b" : palette.surface}
                stroke-width={selected ? 2 : 1}
              x={leaf.x0}
              y={leaf.y0}
              width={width}
              height={height}
            />
            {selected
              ? renderSelectionRectHandles(`treemap-selection-${index}`, leaf.x0, leaf.y0, width, height)
              : null}
            {canShowLabel ? (
              <>
                <text
                  fill={darkenColor(fill, 0.68)}
                  font-size={10}
                  font-weight={600}
                  x={leaf.x0 + 8}
                  y={leaf.y0 + 14}
                >
                  {leaf.data.name}
                </text>
                <text
                  fill={darkenColor(fill, 0.54)}
                  font-size={9}
                  x={leaf.x0 + 8}
                  y={leaf.y0 + 28}
                >
                  {formatTickValue(leaf.value ?? 0)}
                </text>
              </>
            ) : null}
          </g>
        );
      })}
    </g>
  );
}
