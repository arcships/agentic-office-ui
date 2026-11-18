<template>
  <div
    class="xlsx-range-highlight-layer"
    data-testid="xlsx-range-highlight-layer"
    aria-hidden="true"
    :data-highlight-count="highlightCount"
    :style="layerStyle"
  >
    <div
      v-for="pane in panes"
      :key="pane.key"
      class="xlsx-range-highlight-layer__pane"
      :data-grid-pane="pane.key"
      :style="pane.style"
    >
      <div
        v-for="item in pane.items"
        :key="item.key"
        class="xlsx-range-highlight-layer__range"
        data-testid="xlsx-range-highlight"
        :data-highlight-id="item.highlight.id"
        :data-range="item.address"
        :data-role="item.highlight.role"
        :data-sheet-index="item.highlight.sheetIndex"
        :style="item.style"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue";
import type { XlsxRangeHighlight, XlsxViewerController } from "@arcships/xlsx-core";
import { gridBodyLayerStyle } from "./grid-metrics";
import {
  createRangeHighlightLayout,
  rangeHighlightStyle,
  type RangeHighlightSegment,
} from "./range-highlight-geometry";
import { rangeToA1 } from "../composables/selection";

const props = defineProps<{
  controller: XlsxViewerController;
  scrollLeft?: number;
  scrollTop?: number;
  viewportWidth?: number;
  viewportHeight?: number;
}>();

const layerStyle = gridBodyLayerStyle;

const layout = computed(() => createRangeHighlightLayout(
  props.controller.activeSheet,
  props.controller.zoomScale,
  props.scrollLeft ?? 0,
  props.scrollTop ?? 0,
));

const activeSheetIndex = computed(() => props.controller.activeSheetIndex);

const activeHighlights = computed(() => {
  const highlights = props.controller.highlightedRanges ?? [];
  const sheetIndex = activeSheetIndex.value;
  return highlights.filter((highlight) => (highlight.sheetIndex ?? sheetIndex) === sheetIndex);
});

const highlightCount = computed(() => activeHighlights.value.length);

interface HighlightItem extends RangeHighlightSegment {
  address: string;
  highlight: XlsxRangeHighlight;
  key: string;
  style: ReturnType<typeof rangeHighlightStyle>;
}

const items = computed<HighlightItem[]>(() => {
  const viewport = (props.viewportWidth ?? 0) > 0 && (props.viewportHeight ?? 0) > 0
    ? { height: props.viewportHeight ?? 0, width: props.viewportWidth ?? 0 }
    : undefined;
  return activeHighlights.value.flatMap((highlight, highlightIndex) =>
    layout.value.resolve(highlight.range, viewport).map((segment, segmentIndex) => ({
      ...segment,
      address: rangeToA1(highlight.range),
      highlight,
      key: [
        highlight.id ?? `highlight-${highlightIndex}`,
        segment.pane,
        segmentIndex,
      ].join(":"),
      style: rangeHighlightStyle(segment.rect, highlight.color),
    })),
  );
});

const panes = computed(() => layout.value.panes.map((pane) => ({
  ...pane,
  items: items.value.filter((item) => item.pane === pane.key),
})));
</script>

<style scoped>
/* The layer tracks the grid body (headers excluded). Each freeze pane clips
   its own fragments so a highlighted range never paints over the headers,
   frozen cells, or host UI around the surface. It sits above the canvas but
   below charts/images in the surface stacking order. */
.xlsx-range-highlight-layer {
  isolation: isolate;
  pointer-events: none;
  position: absolute;
  z-index: 1;
}
.xlsx-range-highlight-layer__pane {
  overflow: hidden;
  pointer-events: none;
  position: absolute;
}
.xlsx-range-highlight-layer__range {
  border-radius: 1px;
}
</style>
