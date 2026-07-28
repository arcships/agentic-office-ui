<template>
  <div
    class="xlsx-chart-overlay"
    data-testid="xlsx-chart-overlay"
    :data-state="controller.isChartsLoading ? 'loading' : 'ready'"
    :data-chart-count="chartItems.length"
    :style="layerStyle"
  >
    <div
      v-for="pane in panes"
      :key="pane.key"
      class="xlsx-chart-overlay__pane"
      :data-grid-pane="pane.key"
      :style="pane.style"
    >
      <div
        v-for="item in pane.items"
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
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxChart, XlsxChartElementSelection, XlsxImageRect } from "@arcships/xlsx-core";
import { MemoChartSvg } from "../optional/lazy-renderers";
import { createGridPaneLayout, gridBodyLayerStyle, resolveAnchorRect, type GridPaneKey } from "./grid-metrics";
import type { ChartRendererPalette } from "../render/chart-types";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  scrollLeft?: number;
  scrollTop?: number;
}>();

const layerStyle = gridBodyLayerStyle;
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
  pane: GridPaneKey;
  style: CSSProperties;
}

const paneLayout = computed(() => createGridPaneLayout(
  props.controller.activeSheet,
  props.controller.zoomScale,
  props.scrollLeft ?? 0,
  props.scrollTop ?? 0,
));

const chartItems = computed<ChartItem[]>(() => {
  const layout = paneLayout.value;
  return (props.controller.charts as XlsxChart[])
    .filter((chart) => chart.anchor)
    .map((chart) => {
      const rect = resolveAnchorRect(chart.anchor, props.controller.activeSheet, props.controller.zoomScale);
      const isSelected = chart.id === selectedChartId.value;
      const placement = layout.place(rect);
      return {
        chart,
        rect,
        pane: placement.pane,
        style: {
          ...placement.style,
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

const panes = computed(() => paneLayout.value.panes.map((pane) => ({
  ...pane,
  items: chartItems.value.filter((item) => item.pane === pane.key),
})));

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
/* Geometry (inset past the headers) comes from `gridBodyLayerStyle`; each
   freeze pane clips its own charts so a scrolled-away chart never paints over
   frozen panes, the headers, or the host UI around the surface. */
.xlsx-chart-overlay {
  pointer-events: none; position: absolute;
}
.xlsx-chart-overlay__pane {
  overflow: hidden; pointer-events: none; position: absolute;
}
.xlsx-chart-overlay__item { background: transparent; overflow: hidden; }
.xlsx-chart-overlay__item--selected { box-shadow: 0 0 0 1px #2563eb; }
.xlsx-chart-overlay__unsupported {
  align-items: center; display: flex; height: 100%;
  justify-content: center; opacity: 0.5; width: 100%;
}
</style>
