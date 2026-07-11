<template>
  <div
    class="xlsx-chart-overlay"
    data-testid="xlsx-chart-overlay"
    :data-state="controller.isChartsLoading ? 'loading' : 'ready'"
    :data-chart-count="chartItems.length"
  >
    <div
      v-for="item in chartItems"
      :key="item.chart.id"
      class="xlsx-chart-overlay__item"
      data-testid="xlsx-chart-item"
      :class="{
        'xlsx-chart-overlay__item--selected': item.chart.id === selectedChartId,
      }"
      :style="item.style"
      @pointerdown.stop="onChartPointerDown(item.chart, $event)"
    >
      <MemoChartSvg
        v-if="item.chart.chartType !== 'Unsupported'"
        :chart="item.chart"
        :rect="item.rect"
        :palette="chartPalette"
        :selected-chart-element="
          selectedChartId === item.chart.id ? controller.selectedChartElement : null
        "
        :on-chart-element-pointer-down="
          selectedChartId === item.chart.id ? onChartElementPointerDown : undefined
        "
        :on-chart-element-double-click="
          selectedChartId === item.chart.id ? onChartElementDoubleClick : undefined
        "
      />
      <div v-else class="xlsx-chart-overlay__unsupported">
        📊 不支持的图表类型
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxChart, XlsxChartElementSelection, XlsxImageRect, XlsxImageAnchor, XlsxSheetData } from "@arcships/xlsx-core";
import { anchorToAbsoluteRect, emuToPixels } from "@arcships/xlsx-core";
import { MemoChartSvg } from "../optional/lazy-renderers";
import type { ChartRendererPalette } from "../render/chart-types";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  scrollLeft?: number;
  scrollTop?: number;
}>();

const selectedChartId = computed(() => props.controller.selectedChartId);

const chartPalette = computed<ChartRendererPalette>(() => ({
  border: props.isDark ? "#52525b" : "#d4d4d8",
  mutedText: props.isDark ? "#a1a1aa" : "#71717a",
  surface: props.isDark ? "#27272a" : "#f4f4f5",
  text: props.isDark ? "#e4e4e7" : "#18181b",
}));

interface ChartItem {
  chart: XlsxChart;
  rect: XlsxImageRect;
  style: CSSProperties;
}

function frozenExtent(sheet: XlsxSheetData | null, axis: "column" | "row", zoomScale: number): number {
  if (!sheet?.freezePanes) return 0;
  const limit = axis === "column" ? sheet.freezePanes.col : sheet.freezePanes.row;
  const indices = axis === "column" ? sheet.visibleCols : sheet.visibleRows;
  const sizes = axis === "column" ? sheet.colWidths : sheet.rowHeights;
  const fallback = axis === "column" ? sheet.defaultColWidthPx : sheet.defaultRowHeightPx;
  const zoom = Math.max(0.1, zoomScale / 100);
  return indices.reduce((total, actualIndex, displayIndex) =>
    actualIndex < limit ? total + (sizes[displayIndex] || fallback) * zoom : total,
  0);
}

function resolveAnchorRect(anchor: XlsxImageAnchor, sheet: XlsxSheetData | null, zoomScale: number): XlsxImageRect {
  const rect = anchorToAbsoluteRect(anchor, sheet);
  const zoom = Math.max(0.1, zoomScale / 100);
  return {
    left: emuToPixels(rect.x) * zoom,
    top: emuToPixels(rect.y) * zoom,
    width: Math.max(1, emuToPixels(rect.cx) * zoom),
    height: Math.max(1, emuToPixels(rect.cy) * zoom),
  };
}

const chartItems = computed<ChartItem[]>(() => {
  const charts = props.controller.charts as XlsxChart[];
  return charts
    .filter((chart) => chart.anchor)
    .map((chart) => {
      const rect = resolveAnchorRect(chart.anchor, props.controller.activeSheet, props.controller.zoomScale);
      const isSelected = chart.id === selectedChartId.value;
      const frozenWidth = frozenExtent(props.controller.activeSheet, "column", props.controller.zoomScale);
      const frozenHeight = frozenExtent(props.controller.activeSheet, "row", props.controller.zoomScale);
      const scrollX = rect.left < frozenWidth ? 0 : props.scrollLeft ?? 0;
      const scrollY = rect.top < frozenHeight ? 0 : props.scrollTop ?? 0;
      return {
        chart,
        rect,
        style: {
          left: `${rect.left + 48 - scrollX}px`,
          top: `${rect.top + 24 - scrollY}px`,
          width: `${rect.width}px`,
          height: `${rect.height}px`,
          position: "absolute",
          pointerEvents: "auto",
          cursor: "pointer",
          outline: isSelected ? "2px solid #2563eb" : "none",
          outlineOffset: "-1px",
          zIndex: isSelected ? 5 : 1,
        } satisfies CSSProperties,
      };
    });
});

function onChartPointerDown(chart: XlsxChart, _event: PointerEvent) {
  props.controller.selectChart(chart.id);
}

function onChartElementPointerDown(selection: XlsxChartElementSelection, _event: PointerEvent) {
  props.controller.selectChartElement(selection);
}

function onChartElementDoubleClick(selection: XlsxChartElementSelection, _event: MouseEvent) {
  props.controller.selectChartElement(selection);
}
</script>

<style scoped>
.xlsx-chart-overlay {
  height: 100%; left: 0; pointer-events: none;
  position: absolute; top: 0; width: 100%;
}
.xlsx-chart-overlay__item { background: transparent; overflow: hidden; }
.xlsx-chart-overlay__item--selected { box-shadow: 0 0 0 1px #2563eb; }
.xlsx-chart-overlay__unsupported {
  align-items: center; display: flex; height: 100%;
  justify-content: center; opacity: 0.5; width: 100%;
}
</style>
