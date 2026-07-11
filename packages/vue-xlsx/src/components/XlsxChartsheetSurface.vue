<template>
  <div class="xlsx-chartsheet" data-testid="xlsx-chartsheet" :data-state="controller.isChartsLoading ? 'loading' : 'ready'">
    <div class="xlsx-chartsheet__surface">
      <div
        v-for="chart in charts"
        :key="chart.id"
        class="xlsx-chartsheet__chart"
        data-testid="xlsx-chartsheet-chart"
        @pointerdown="controller.selectChart(chart.id)"
      >
        <div v-if="controller.isChartsLoading" class="xlsx-chartsheet__loading">Loading chart…</div>
        <MemoChartSvg
          v-else-if="chart.chartType !== 'Unsupported'"
          :chart="chart"
          :rect="chartRect"
          :palette="palette"
          :selected-chart-element="controller.selectedChartId === chart.id ? controller.selectedChartElement : null"
          :on-chart-element-pointer-down="onChartElementPointerDown"
          :on-chart-element-double-click="onChartElementDoubleClick"
        />
        <div v-else class="xlsx-chartsheet__empty">This chart type is not available yet.</div>
      </div>
      <div v-if="charts.length === 0 && !controller.isChartsLoading" class="xlsx-chartsheet__empty">
        This chart sheet does not expose an embedded chart payload.
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue";
import type { XlsxChart, XlsxChartElementSelection, XlsxViewerController } from "@arcships/xlsx-core";
import { MemoChartSvg } from "../optional/lazy-renderers";
import type { ChartRendererPalette } from "../render/chart-types";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
}>();

const chartRect = { height: 360, left: 0, top: 0, width: 640 };
const charts = computed(() => props.controller.charts as XlsxChart[]);
const palette = computed<ChartRendererPalette>(() => ({
  border: props.isDark ? "#52525b" : "#d4d4d8",
  mutedText: props.isDark ? "#a1a1aa" : "#71717a",
  surface: props.isDark ? "#27272a" : "#ffffff",
  text: props.isDark ? "#f4f4f5" : "#18181b",
}));

function onChartElementPointerDown(selection: XlsxChartElementSelection): void {
  props.controller.selectChartElement(selection);
}

function onChartElementDoubleClick(selection: XlsxChartElementSelection): void {
  props.controller.selectChartElement(selection);
}
</script>

<style scoped>
.xlsx-chartsheet {
  background: #f4f4f5;
  box-sizing: border-box;
  display: flex;
  flex: 1;
  min-height: 0;
  min-width: 0;
  overflow: auto;
  padding: 16px;
}
.xlsx-chartsheet__surface {
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 12px;
  box-sizing: border-box;
  display: grid;
  flex: 1;
  gap: 16px;
  grid-template-columns: repeat(auto-fit, minmax(min(100%, 360px), 1fr));
  min-width: 0;
  padding: 16px;
}
.xlsx-chartsheet__chart { min-height: 360px; min-width: 0; overflow: hidden; position: relative; }
.xlsx-chartsheet__chart :deep(svg) { height: 100%; max-height: 440px; width: 100%; }
.xlsx-chartsheet__loading,
.xlsx-chartsheet__empty {
  align-items: center;
  color: #71717a;
  display: flex;
  font-size: 13px;
  justify-content: center;
  min-height: 320px;
  text-align: center;
}
</style>
