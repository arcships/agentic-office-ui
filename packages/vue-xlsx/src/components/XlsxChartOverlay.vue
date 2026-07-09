<template>
  <div class="xlsx-chart-overlay">
    <div
      v-for="item in chartItems"
      :key="item.chart.id"
      class="xlsx-chart-overlay__item"
      :class="{
        'xlsx-chart-overlay__item--selected': item.chart.id === selectedChartId,
      }"
      :style="item.style"
      @pointerdown.stop="onChartPointerDown(item.chart, $event)"
      @dblclick.stop="onChartDoubleClick(item.chart, $event)"
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
import type { XlsxViewerController, XlsxChart, XlsxChartElementSelection, XlsxImageRect, XlsxImageAnchor } from "@extend-ai/xlsx-core";
import { emuToPixels } from "@extend-ai/xlsx-core";
import { MemoChartSvg } from "../render";
import type { ChartRendererPalette } from "../render";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
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

function resolveAnchorRect(anchor: XlsxImageAnchor): XlsxImageRect | null {
  if (anchor.kind === "absolute") {
    return {
      left: emuToPixels(anchor.positionEmu.x),
      top: emuToPixels(anchor.positionEmu.y),
      width: emuToPixels(anchor.sizeEmu.cx),
      height: emuToPixels(anchor.sizeEmu.cy),
    };
  }
  if (anchor.kind === "one-cell") {
    return {
      left: emuToPixels(anchor.from.colOffsetEmu),
      top: emuToPixels(anchor.from.rowOffsetEmu),
      width: emuToPixels(anchor.sizeEmu.cx),
      height: emuToPixels(anchor.sizeEmu.cy),
    };
  }
  return null;
}

const chartItems = computed<ChartItem[]>(() => {
  const charts = props.controller.charts as XlsxChart[];
  return charts
    .filter((chart) => chart.anchor)
    .map((chart) => {
      const rect = resolveAnchorRect(chart.anchor);
      if (!rect) {
        return {
          chart,
          rect: { left: 0, top: 0, width: 200, height: 200 },
          style: { display: "none" } as CSSProperties,
        };
      }
      const isSelected = chart.id === selectedChartId.value;
      return {
        chart,
        rect,
        style: {
          left: `${rect.left + 48}px`,
          top: `${rect.top + 24}px`,
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

function onChartDoubleClick(_chart: XlsxChart, _event: MouseEvent) {}

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
