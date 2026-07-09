/** @jsxImportSource vue */
import type { XlsxChart, XlsxImageRect } from "@extend-ai/xlsx-core";
import {
  buildHierarchyData,
  buildPieEntries,
  buildRegionMapLegendItems,
  buildStockLegendItems,
  buildSurfaceLegendItems,
  chartElementDataProps,
  chartPointColor,
  chartSeriesColor,
  getCategoryLabels,
  getComboLegendSeries,
  isComboChart,
  normalizeCategoryLabel,
  normalizeLegendPosition,
  resolveChartAxisTextColor,
  safeNumber,
} from "./chart-renderer";
import type { ChartLayout, ChartRendererPalette, LegendItem } from "./chart-renderer";

export function getLegendItems(chart: XlsxChart, chartType: string, palette: ChartRendererPalette): LegendItem[] {
  if (!chart.legend) {
    return [];
  }
  if (isComboChart(chart)) {
    return getComboLegendSeries(chart);
  }
  if (chartType === "Stock") {
    return buildStockLegendItems(chart, palette);
  }
  if (chartType === "Surface") {
    return buildSurfaceLegendItems(chart, palette);
  }
  if (chartType === "RegionMap") {
    return buildRegionMapLegendItems(chart);
  }
  if (chartType === "Sunburst" || chartType === "Treemap") {
    const hierarchyData = buildHierarchyData(chart);
    if (!hierarchyData?.children) {
      return [];
    }
    return hierarchyData.children.map((child, index) => ({
      color: chartSeriesColor(chart, child.colorIndex ?? index),
      label: child.name
    }));
  }
  if (chartType === "Pie" || chartType === "Pie3D" || chartType === "PieExploded" || chartType === "Doughnut" || chartType === "BarOfPie") {
    if (chartType === "BarOfPie") {
      const categories = getCategoryLabels(chart);
      const values = chart.series[0]?.values ?? [];
      return categories
        .map((label, index) => ({
          color: chartPointColor(chart, index),
          label: normalizeCategoryLabel(label),
          value: safeNumber(values[index]) ?? 0
        }))
        .filter((entry) => entry.value > 0 && entry.label.trim().length > 0)
        .map((entry) => ({
          color: entry.color,
          label: entry.label
        }));
    }
    return buildPieEntries(chart).map((entry) => ({
      color: entry.color,
      label: entry.label
    }));
  }
  const isXyLegend = (
    chartType === "Scatter"
    || chartType === "ScatterLines"
    || chartType === "ScatterSmooth"
    || chartType === "Bubble"
  );
  return chart.series.map((series, index) => ({
    color: isXyLegend
      ? (series.lineColor ?? series.markerColor ?? series.color ?? chartSeriesColor(chart, index))
      : chartSeriesColor(chart, index),
    label: series.name ?? `Series ${index + 1}`
  }));
}

export function buildLayout(chart: XlsxChart, rect: XlsxImageRect, legendItems: LegendItem[]): ChartLayout {
  const width = Math.max(80, Math.round(rect.width));
  const height = Math.max(60, Math.round(rect.height));
  const isSurfaceChart = chart.chartType === "Surface"
    || (chart.raw && typeof chart.raw === "object" && typeof (chart.raw as Record<string, unknown>).xmlChartType === "string" && String((chart.raw as Record<string, unknown>).xmlChartType).includes("surface"));
  const titleHeight = chart.title ? (isSurfaceChart ? 30 : 24) : 8;
  const legendPosition = normalizeLegendPosition(chart.legend?.position);

  const legendVertical = legendItems.length > 0 && (legendPosition === "right" || legendPosition === "left");
  const legendHorizontal = legendItems.length > 0 && (legendPosition === "top" || legendPosition === "bottom");

  const plotLeft = 42 + (legendPosition === "left" ? 98 : 0);
  const plotRight = 16 + (legendPosition === "right" ? 98 : 0);
  const plotTop = titleHeight + (legendPosition === "top" ? 24 : 0);
  const plotBottom = 28 + (legendPosition === "bottom" ? 24 : 0);

  const plotWidth = Math.max(40, width - plotLeft - plotRight);
  const plotHeight = Math.max(40, height - plotTop - plotBottom);

  const compactWidth = width <= 280;
  const compactHeight = height <= 200;
  const finalLegendPosition = compactWidth || compactHeight
    ? (legendVertical ? "bottom" : legendPosition)
    : legendPosition;

  return {
    height,
    legendItems,
    legendPosition: finalLegendPosition,
    plot: {
      height: Math.max(40, height - plotTop - plotBottom),
      left: plotLeft,
      top: plotTop,
      width: Math.max(40, width - plotLeft - plotRight)
    },
    titleHeight,
    width
  };
}

export function renderLegend(chart: XlsxChart, layout: ChartLayout, palette: ChartRendererPalette) {
  if (!chart.legend || layout.legendItems.length === 0) {
    return null;
  }
  const textColor = resolveChartAxisTextColor(chart);
  const legendPos = layout.legendPosition;
  const items = layout.legendItems;
  const swatchSize = 8;
  const textOffset = swatchSize + 4;

  if (legendPos === "left" || legendPos === "right") {
    const x = legendPos === "right"
      ? layout.plot.left + layout.plot.width + 8
      : 8;
    const startY = layout.plot.top + 6;
    return (
      <g>
        {items.map((item, index) => {
          const y = startY + index * 18;
          return (
              <g
                {...(index < chart.series.length ? chartElementDataProps(index) : {})}
                data-xlsx-chart-element-kind={index < chart.series.length ? "legendEntry" : undefined}
                key={`legend-${index}`}
                transform={`translate(${x}, ${y})`}
              >
              <rect fill={item.color} height={swatchSize} rx={1.2} ry={1.2} width={swatchSize} x={0} y={-7} />
              <text fill={textColor} font-size={10} x={textOffset} y={0}>
                {item.label}
              </text>
            </g>
          );
        })}
      </g>
    );
  }

  const rowY = legendPos === "top" ? (layout.titleHeight + 12) : (layout.height - 8);
  const totalWidth = items.reduce((sum, item) => sum + 24 + Math.min(96, item.label.length * 5.4), 0);
  let cursorX = Math.max(8, (layout.width - totalWidth) / 2);
  return (
    <g>
      {items.map((item, index) => {
        const labelWidth = Math.min(96, item.label.length * 5.4);
        const node = (
            <g
              {...(index < chart.series.length ? chartElementDataProps(index) : {})}
              data-xlsx-chart-element-kind={index < chart.series.length ? "legendEntry" : undefined}
              key={`legend-${index}`}
              transform={`translate(${cursorX}, ${rowY})`}
            >
            <rect fill={item.color} height={swatchSize} rx={1.2} ry={1.2} width={swatchSize} x={0} y={-7} />
            <text fill={textColor} font-size={10} x={textOffset} y={0}>
              {item.label}
            </text>
          </g>
        );
        cursorX += 24 + labelWidth;
        return node;
      })}
    </g>
  );
}
