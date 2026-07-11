/** @jsxImportSource vue */
import { defineComponent } from "vue";
import type { XlsxChart, XlsxChartElementSelection, XlsxImageRect } from "@arcships/xlsx-core";

// Re-export all chart utilities from their dedicated modules
export type { ChartRendererPalette, ChartSvgProps, LegendItem, PlotRect, ChartLayout, BarRect, ChartElementDataOptions, ChartHierarchyDatum, ComboRenderableGroup, ChartStage, BoxWhiskerStats, SurfaceDomain, RegionMapFeature, SurfacePathPoint, SurfaceContourSegment } from "./chart-types";

export {
  chartElementDataProps,
  barChartElementDataProps,
  resolveChartElementTarget,
  resolveChartSelectionFromTarget,
  isSelectedChartSeries,
  isSelectedChartPoint,
  isSelectedChartPointOrSeries,
  renderSelectionRectHandles,
  renderSelectionPointHandles
} from "./chart-element";

export {
  parseRgbColor,
  mixRgbColor,
  lightenColor,
  darkenColor,
  clamp,
  normalizeRendererHexColor,
  DEFAULT_CHART_FONT_STACK,
  DEFAULT_CHART_TEXT_COLOR,
  DEFAULT_CHART_MUTED_TEXT_COLOR,
  escapeCssFontFamilyToken,
  buildChartFontFamily,
  resolveChartTextColor,
  resolveChartAxisTextColor,
  resolveChartMutedTextColor,
  safeNumber,
  normalizeLegendPosition,
  normalizeCategoryLabel,
  truncateSvgText
} from "./chart-shared";

export {
  chartSeriesColor,
  chartSeriesStrokeColor,
  chartPointColor,
  isHistogramLikeSeries,
  isHistogramLikeChart,
  normalizeBuiltinPieStyleId,
  getBuiltinPiePalette,
  resolvePiePointColor,
  selectPrimaryPieSeriesIndex,
  chartSeriesBarColors,
  resolveCategoryBandPadding,
  normalizeChartMarkerSymbol,
  normalizeRenderableChartType,
  estimateReferencePointCount,
  getCategoryLabels,
  isComboChart,
  getComboLegendSeries,
  findAxisForGroup,
  buildComboGroups,
  resolveRenderableSeriesValue,
  coerceLooseNumber,
  buildPieEntries,
  isLikelyDateFormatCode,
  excelSerialToDate,
  formatExcelDateSerial,
  formatCategoryLabel
} from "./chart-data";

export {
  WORLD_COUNTRY_FEATURES,
  US_STATE_NAME_BY_ID,
  US_STATE_FEATURES,
  REGION_MAP_COUNTRY_ALIASES,
  REGION_MAP_US_STATE_ALIASES,
  REGION_MAP_FEATURES_BY_KEY,
  REGION_MAP_US_STATE_FEATURES_BY_KEY,
  normalizeRegionMapKey,
  resolveRegionMapFeature,
  resolveRegionMapBaseColor,
  resolveRegionMapDataColor,
  resolveRegionMapValueColors,
  resolveRegionMapColorStrings,
  resolveRegionMapLayoutProperties,
  resolveRegionMapFeatureSet,
  getRegionMapBaseFeatures,
  resolveRegionMapValueColorFromStops,
  resolveRegionMapValueColor,
  resolveRegionMapNoDataColor,
  buildRegionMapLegendItems
} from "./chart-region-map";

export {
  resolveChartStageSubtotal,
  buildChartStages,
  buildHierarchyData,
  resolveBoxWhiskerQuartileMethod,
  resolveBoxWhiskerVisibility,
  computePercentile,
  computeBoxWhiskerStats,
  resolveHierarchyNodeColor,
  resolveTreemapNodeColor,
  excelTreemapTile
} from "./chart-analysis";

export {
  resolveSurfaceBaseColor,
  normalizeBuiltinSurfaceStyleId,
  hasExplicitSurfaceBaseColor,
  buildMonochromeSurfacePalette,
  getBuiltinSurfacePalette,
  shouldPreferBuiltinSurfacePalette,
  getSurfaceBandCount,
  getSurfaceColorStops,
  resolveSurfaceBandPaletteColor,
  getSurfaceDomain,
  resolveSurfaceColor,
  resolveSurfaceBandColor,
  resolveSurfaceBandIndex,
  buildSurfaceSmoothPath,
  splitSurfacePointRuns,
  connectSurfaceContourSegments,
  getSurfaceWireframePalette,
  resolveSurfaceWireframeColor,
  resolveSurfacePlotRect,
  buildSurfaceLegendItems,
  isContourSurfaceChart
} from "./chart-surface-utils";

export {
  TWO_PI,
  toSvgNumber,
  indexByName,
  resolveStockRoleIndices,
  resolveStockPalette,
  buildStockLegendItems,
  markerSymbolPath,
  normalizePieArc,
  resolvePieFrontSegments,
  pieEllipsePoint,
  buildPieOuterWallPath,
  isPieFrontFacingAngle,
  buildPieRadialWallPath,
  renderTitle,
  renderUnsupported
} from "./chart-svg-utils";

export {
  resolve3dFrameOffsets,
  scaleProjectedVector,
  sampleProjectedEllipseArc,
  buildRibbonSvgPath,
  renderRadialFrustum,
  renderExtrudedRect,
  buildLinearSvgPath,
  projectCartesian3dPoint,
  renderLineOrAreaChart3d
} from "./chart-svg-pipeline";

export { buildLayout, getLegendItems, renderLegend } from "./chart-legend";
export { buildNiceStep, buildNumericTickValues, formatPercentTickValue, formatTickValue, renderCartesianAxes, renderSurfaceAxes, resolveAxisDomainWithChartOverrides, resolveNumericAxisDomain } from "./chart-axis";

// Chart type renderers
import { renderBarChart, renderWaterfallChart, renderFunnelChart, renderBoxWhiskerChart } from "./chart-bar";
import { renderLineOrAreaChart, renderComboChart, renderRadarChart, renderStockChart } from "./chart-line";
import { renderPieChart, renderBarOfPieChart, renderSunburstChart, renderTreemapChart } from "./chart-pie";
import { renderScatterChart, renderBubbleChart } from "./chart-scatter";
import { renderSurfaceChart, renderSurfaceHitOverlay, renderRegionMapChart } from "./chart-surface";
import { MemoSurfaceChartComposite } from "./surface-regl";
import type { ChartRendererPalette, ChartLayout } from "./chart-types";
import { buildChartFontFamily, resolveChartMutedTextColor } from "./chart-shared";
import { normalizeRenderableChartType, isComboChart } from "./chart-data";
import { getLegendItems } from "./chart-legend";
import { buildLayout } from "./chart-legend";
import { renderTitle } from "./chart-svg-utils";
import { renderUnsupported } from "./chart-svg-utils";
import { renderLegend } from "./chart-legend";
import { resolveChartSelectionFromTarget, resolveChartElementTarget } from "./chart-element";
import { renderSurfaceAxes } from "./chart-axis";

export function renderChartPlot(
  chart: XlsxChart,
  palette: ChartRendererPalette,
  layout: ChartLayout,
  chartType: string,
  selectedChartElement?: XlsxChartElementSelection | null
) {
  if (isComboChart(chart)) {
    return renderComboChart(chart, palette, layout, selectedChartElement) ?? renderUnsupported(chart, palette, layout, "Combo");
  }
  if (
    chartType === "ColumnClustered"
    || chartType === "ColumnStacked"
    || chartType === "ColumnPercentStacked"
    || chartType === "BarClustered"
    || chartType === "BarStacked"
    || chartType === "BarPercentStacked"
  ) {
    return renderBarChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (
    chartType === "Line"
    || chartType === "LineStacked"
    || chartType === "LinePercentStacked"
    || chartType === "Area"
    || chartType === "AreaStacked"
    || chartType === "AreaPercentStacked"
  ) {
    return renderLineOrAreaChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (chartType === "Scatter") {
    return renderScatterChart(chart, palette, layout, false, selectedChartElement);
  }
  if (chartType === "ScatterLines") {
    return renderScatterChart(chart, palette, layout, false, selectedChartElement);
  }
  if (chartType === "ScatterSmooth") {
    return renderScatterChart(chart, palette, layout, true, selectedChartElement);
  }
  if (chartType === "Bubble") {
    return renderBubbleChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Radar") {
    return renderRadarChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Pie" || chartType === "Pie3D" || chartType === "PieExploded" || chartType === "Doughnut") {
    return renderPieChart(chart, palette, layout, chartType, selectedChartElement);
  }
  if (chartType === "BarOfPie") {
    return renderBarOfPieChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Surface") {
    return renderSurfaceChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Stock") {
    return renderStockChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Waterfall") {
    return renderWaterfallChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Funnel") {
    return renderFunnelChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "BoxWhisker") {
    return renderBoxWhiskerChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Sunburst") {
    return renderSunburstChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "Treemap") {
    return renderTreemapChart(chart, palette, layout, selectedChartElement);
  }
  if (chartType === "RegionMap") {
    return renderRegionMapChart(chart, palette, layout, selectedChartElement);
  }
  return renderUnsupported(chart, palette, layout, chartType);
}

export const MemoChartSvg = /* @__PURE__ */ defineComponent({
  name: "MemoChartSvg",
  props: {
    chart: { type: Object as () => XlsxChart, required: true },
    onChartElementDoubleClick: { type: Function as unknown as () => ((selection: XlsxChartElementSelection, event: MouseEvent) => void) | undefined, default: undefined },
    onChartElementPointerDown: { type: Function as unknown as () => ((selection: XlsxChartElementSelection, event: PointerEvent) => void) | undefined, default: undefined },
    palette: { type: Object as () => ChartRendererPalette, required: true },
    rect: { type: Object as () => XlsxImageRect, required: true },
    selectedChartElement: { type: Object as () => XlsxChartElementSelection | null | undefined, default: null },
  },
  setup(props) {
    return () => {
      const { chart, onChartElementDoubleClick, onChartElementPointerDown, palette, rect, selectedChartElement } = props;
      const renderChartType = normalizeRenderableChartType(chart);
      const legendItems = getLegendItems(chart, renderChartType, palette);
      const layout = buildLayout(chart, rect, legendItems);
      const chartRaw = chart.raw && typeof chart.raw === "object" ? chart.raw as Record<string, unknown> : null;
      const explicitNoFill = chartRaw?.chartAreaNoFill === true || chartRaw?.plotAreaNoFill === true;
      const background = chart.chartAreaFillColor ?? (explicitNoFill ? "transparent" : "#ffffff");
      const borderColor = chart.chartAreaBorderColor ?? "transparent";
      const normalizedBackground = background.trim().toLowerCase();
      const normalizedBorderColor = borderColor.trim().toLowerCase();
      const hideBackgroundRect = normalizedBackground === "transparent" && normalizedBorderColor === "transparent";
      const fontFamily = buildChartFontFamily(chart.fontFamily);

      const handlePointerDown = (event: PointerEvent) => {
        const selection = resolveChartSelectionFromTarget(chart, event.target, selectedChartElement);
        if (!selection) {
          return;
        }
        if (resolveChartElementTarget(event.target)) {
          event.stopPropagation();
        }
        onChartElementPointerDown?.(selection, event);
      };

      const handleDoubleClick = (event: MouseEvent) => {
        const selection = resolveChartSelectionFromTarget(chart, event.target, selectedChartElement);
        if (!selection) {
          return;
        }
        if (resolveChartElementTarget(event.target)) {
          event.stopPropagation();
        }
        onChartElementDoubleClick?.(selection, event);
      };

      if (renderChartType === "Surface") {
        return (
          <MemoSurfaceChartComposite
            background={background}
            borderColor={borderColor}
            chart={chart}
            fallback={renderSurfaceChart(chart, palette, layout, selectedChartElement)}
            fontFamily={fontFamily}
            layout={layout}
            onDoubleClick={handleDoubleClick}
            onPointerDown={handlePointerDown}
            overlay={
              <>
                {renderSurfaceHitOverlay(chart, layout, selectedChartElement)}
                {renderSurfaceAxes(chart, layout)}
                {renderTitle(chart, layout, palette)}
                {renderLegend(chart, layout, palette)}
              </>
            }
            palette={palette}
          />
        );
      }

      return (
        <svg
          aria-label={chart.title ?? chart.name ?? "Chart"}
          onDblclick={handleDoubleClick}
          onPointerdown={handlePointerDown}
          role="img"
          style={{ display: "block", fontFamily, height: "100%", pointerEvents: "auto", width: "100%" }}
          viewBox={`0 0 ${layout.width} ${layout.height}`}
        >
          {hideBackgroundRect
            ? null
            : <rect fill={background} height={layout.height} stroke={borderColor} stroke-width={1} width={layout.width} x={0} y={0} />}
          {renderTitle(chart, layout, palette)}
          {renderLegend(chart, layout, palette)}
          {renderChartPlot(chart, palette, layout, renderChartType, selectedChartElement)}
        </svg>
      );
    };
  },
});
