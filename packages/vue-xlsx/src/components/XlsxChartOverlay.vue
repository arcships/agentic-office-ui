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
        tabindex="0"
        :class="{
          'xlsx-chart-overlay__item--selected': item.chart.id === selectedChartId,
          'xlsx-chart-overlay__item--gesture': interaction.isActive(item.chart.id),
        }"
        :style="item.style"
        @pointerdown.stop="onChartPointerDown(item, $event)"
        @contextmenu.prevent.stop="onChartContextMenu(item, $event)"
        @keydown="onChartKeydown(item.chart, $event)"
      >
        <MemoChartSvg
          v-if="item.chart.chartType !== 'Unsupported'"
          :chart="item.chart"
          :chart-labels="{
            ariaLabel: m.xlsx.chart.ariaLabel,
            unsupportedType: m.xlsx.chart.unsupportedType,
            other: m.xlsx.chart.other,
            seriesName: m.xlsx.chart.seriesName,
            surfaceSeries: m.xlsx.chart.surfaceSeries,
          }"
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
          📊 {{ m.xlsx.chart.unsupported }}
        </div>
        <div
          v-if="item.chart.id === selectedChartId"
          class="xlsx-chart-overlay__handles"
          data-testid="xlsx-chart-handles"
        >
          <div
            v-for="handle in DRAWING_RESIZE_HANDLES"
            :key="handle"
            class="xlsx-chart-overlay__handle"
            data-testid="xlsx-chart-resize-handle"
            :data-handle="handle"
            :style="handleStyle(handle)"
            @pointerdown.stop="onHandlePointerDown(item, handle, $event)"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxChart, XlsxChartElementSelection, XlsxImageRect, XlsxImageResizeHandlePosition } from "@arcships/xlsx-core";
import { MemoChartSvg } from "../optional/lazy-renderers";
import { createGridPaneLayout, gridBodyLayerStyle, resolveAnchorRect, type GridPaneKey } from "./grid-metrics";
import { DRAWING_RESIZE_HANDLES, drawingHandleStyle, useDrawingGestures } from "./drawing-interaction";
import type { ChartRendererPalette } from "../render/chart-types";
import { useXlsxMessages } from "../i18n";

const emit = defineEmits<{
  chartContextMenu: [payload: { chartId: string; clientX: number; clientY: number }];
}>();

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  scrollLeft?: number;
  scrollTop?: number;
  viewportWidth?: number;
  viewportHeight?: number;
}>();

const layerStyle = gridBodyLayerStyle;
const m = useXlsxMessages();
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

const interaction = useDrawingGestures({
  onSelect: (id) => props.controller.selectChart(id),
  onMove: (id, deltaX, deltaY) => props.controller.moveChartBy(id, deltaX, deltaY),
  onResize: (id, handle, deltaX, deltaY) => props.controller.resizeChartBy(id, handle, deltaX, deltaY),
});

function numericPx(value: string | number | undefined): number {
  return typeof value === "number" ? value : Number.parseFloat(value ?? "0") || 0;
}

function handleStyle(handle: XlsxImageResizeHandlePosition): CSSProperties {
  return drawingHandleStyle(handle);
}

const chartItems = computed<ChartItem[]>(() => {
  const layout = paneLayout.value;
  return (props.controller.charts as XlsxChart[])
    .filter((chart) => chart.anchor)
    .flatMap((chart) => {
      const rect = resolveAnchorRect(chart.anchor, props.controller.activeSheet, props.controller.zoomScale);
      const isSelected = chart.id === selectedChartId.value;
      const placement = layout.place(rect);
      if (!layout.isVisible(rect, placement.pane, props.viewportWidth ?? 0, props.viewportHeight ?? 0)) {
        return [];
      }
      const preview = interaction.previewFor(chart.id);
      const style: CSSProperties = {
        ...placement.style,
        position: "absolute",
        pointerEvents: "auto",
        cursor: isSelected ? "move" : "pointer",
        outline: isSelected ? "2px solid #2563eb" : "none",
        outlineOffset: "-1px",
        zIndex: isSelected ? 5 : 1,
      };
      if (preview) {
        style.left = `${numericPx(placement.style.left) + preview.left - rect.left}px`;
        style.top = `${numericPx(placement.style.top) + preview.top - rect.top}px`;
        style.width = `${preview.width}px`;
        style.height = `${preview.height}px`;
      }
      return [{ chart, rect, pane: placement.pane, style }];
    });
});

const panes = computed(() => paneLayout.value.panes.map((pane) => ({
  ...pane,
  items: chartItems.value.filter((item) => item.pane === pane.key),
})));

function onChartPointerDown(item: ChartItem, event: PointerEvent) {
  interaction.beginMove(item.chart.id, item.rect, event);
}

function onHandlePointerDown(item: ChartItem, handle: XlsxImageResizeHandlePosition, event: PointerEvent) {
  interaction.beginResize(item.chart.id, item.rect, handle, event);
}

function onChartContextMenu(item: ChartItem, event: MouseEvent) {
  props.controller.selectChart(item.chart.id);
  emit("chartContextMenu", {
    chartId: item.chart.id,
    clientX: event.clientX,
    clientY: event.clientY,
  });
}

function onChartElementPointerDown(selection: XlsxChartElementSelection, _event: PointerEvent) {
  props.controller.selectChartElement(selection);
}

function onChartElementDoubleClick(selection: XlsxChartElementSelection, _event: MouseEvent) {
  props.controller.selectChartElement(selection);
}

function onChartKeydown(chart: XlsxChart, event: KeyboardEvent) {
  if (props.controller.selectedChartId !== chart.id) {
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
    event.preventDefault();
    event.stopPropagation();
    void props.controller.copySelectionToClipboard?.();
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "x") {
    event.preventDefault();
    event.stopPropagation();
    props.controller.cutSelection?.();
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
    event.preventDefault();
    event.stopPropagation();
    void props.controller.pasteFromClipboard?.();
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
    event.preventDefault();
    event.stopPropagation();
    props.controller.duplicateDrawing?.();
    return;
  }
  const step = event.shiftKey ? 10 : 1;
  const nudge: Record<string, [number, number]> = {
    ArrowLeft: [-1, 0],
    ArrowRight: [1, 0],
    ArrowUp: [0, -1],
    ArrowDown: [0, 1],
  };
  const delta = nudge[event.key];
  if (delta) {
    event.preventDefault();
    event.stopPropagation();
    props.controller.moveChartBy(chart.id, delta[0] * step, delta[1] * step);
    return;
  }
  if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault();
    event.stopPropagation();
    props.controller.deleteChart(chart.id);
    return;
  }
  if (event.key === "Escape") {
    event.stopPropagation();
    if (props.controller.cancelCut?.()) return;
    props.controller.clearSelectedChart();
  }
}
</script>

<style scoped>
/* Geometry (inset past the headers) comes from `gridBodyLayerStyle`; each
   freeze pane clips its own charts so a scrolled-away chart never paints over
   frozen panes, the headers, or the host UI around the surface. The layer is
   its own stacking context so pane/object z-index values cannot escape it. */
.xlsx-chart-overlay {
  pointer-events: none; position: absolute;
  isolation: isolate; z-index: 1;
}
.xlsx-chart-overlay__pane {
  overflow: hidden; pointer-events: none; position: absolute;
}
.xlsx-chart-overlay__item { background: transparent; overflow: visible; }
.xlsx-chart-overlay__item--selected { box-shadow: 0 0 0 1px #2563eb; }
.xlsx-chart-overlay__item--gesture { cursor: move; }
.xlsx-chart-overlay__unsupported {
  align-items: center; display: flex; height: 100%;
  justify-content: center; opacity: 0.5; width: 100%;
}
.xlsx-chart-overlay__handles { inset: 0; pointer-events: none; position: absolute; }
.xlsx-chart-overlay__handle {
  background: #ffffff; border: 1px solid #2563eb; box-sizing: border-box;
  height: 8px; pointer-events: auto; position: absolute; width: 8px;
  z-index: 1000000;
}
</style>
