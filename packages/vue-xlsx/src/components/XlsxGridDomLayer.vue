<template>
  <div
    class="xlsx-grid__dom"
    data-testid="xlsx-grid-dom"
    data-render-mode="dom"
    :style="rootStyle"
  >
    <div
      class="xlsx-grid__dom-strip xlsx-grid__dom-strip--column"
      data-testid="xlsx-grid-dom-column-headers"
      aria-hidden="true"
      :style="columnHeaderStripStyle"
    >
      <div :style="columnHeaderScrollStyle">
        <span
          v-for="item in model.columnHeaders.scrolling"
          :key="item.key"
          class="xlsx-grid__dom-header"
          data-column-header
          :data-column="item.label"
          :style="item.style"
        >{{ item.label }}</span>
      </div>
      <div class="xlsx-grid__dom-strip-frozen">
        <span
          v-for="item in model.columnHeaders.frozen"
          :key="item.key"
          class="xlsx-grid__dom-header"
          data-column-header
          :data-column="item.label"
          :style="item.style"
        >{{ item.label }}</span>
      </div>
      <div
        v-if="model.columnHeaders.divider"
        class="xlsx-grid__dom-divider"
        aria-hidden="true"
        :style="model.columnHeaders.divider.style"
      />
    </div>
    <div
      class="xlsx-grid__dom-strip xlsx-grid__dom-strip--row"
      data-testid="xlsx-grid-dom-row-headers"
      aria-hidden="true"
      :style="rowHeaderStripStyle"
    >
      <div :style="rowHeaderScrollStyle">
        <span
          v-for="item in model.rowHeaders.scrolling"
          :key="item.key"
          class="xlsx-grid__dom-header"
          data-row-header
          :data-row="item.label"
          :style="item.style"
        >{{ item.label }}</span>
      </div>
      <div class="xlsx-grid__dom-strip-frozen">
        <span
          v-for="item in model.rowHeaders.frozen"
          :key="item.key"
          class="xlsx-grid__dom-header"
          data-row-header
          :data-row="item.label"
          :style="item.style"
        >{{ item.label }}</span>
      </div>
      <div
        v-if="model.rowHeaders.divider"
        class="xlsx-grid__dom-divider"
        aria-hidden="true"
        :style="model.rowHeaders.divider.style"
      />
    </div>
    <div
      class="xlsx-grid__dom-corner"
      data-testid="xlsx-grid-dom-corner"
      aria-hidden="true"
      :style="model.cornerStyle"
    />
    <div class="xlsx-grid__dom-body" :style="bodyStyle">
      <div
        v-for="pane in model.panes"
        :key="pane.key"
        class="xlsx-grid__dom-pane"
        :data-pane="pane.key"
        :style="pane.style"
      >
        <div class="xlsx-grid__dom-pane-content" :style="paneContentStyle(pane)">
          <div
            v-for="row in pane.rows"
            :key="row.key"
            role="row"
            class="xlsx-grid__dom-row"
            :aria-rowindex="row.ariaRowIndex"
            :style="row.style"
          >
            <div
              v-for="cell in row.cells"
              :id="cell.id"
              :key="cell.key"
              role="gridcell"
              class="xlsx-grid__dom-cell"
              data-testid="xlsx-grid-dom-cell"
              :data-cell="cell.address"
              :data-active="cell.active ? 'true' : undefined"
              :data-selected="cell.selected ? 'true' : undefined"
              :aria-rowindex="cell.ariaRowIndex"
              :aria-colindex="cell.ariaColIndex"
              :aria-selected="cell.selected ? 'true' : 'false'"
              :aria-label="cell.ariaLabel"
              :style="cell.style"
            >
              <span
                v-if="cell.dataBarStyle"
                class="xlsx-grid__dom-data-bar"
                data-testid="xlsx-grid-dom-data-bar"
                aria-hidden="true"
                :style="cell.dataBarStyle"
              />
              <span
                v-if="cell.dataBarAxisStyle"
                class="xlsx-grid__dom-data-bar-axis"
                data-testid="xlsx-grid-dom-data-bar-axis"
                aria-hidden="true"
                :style="cell.dataBarAxisStyle"
              />
              <svg
                v-if="cell.sparkline"
                class="xlsx-grid__dom-sparkline"
                data-testid="xlsx-grid-dom-sparkline"
                aria-hidden="true"
                :viewBox="cell.sparkline.viewBox"
                preserveAspectRatio="none"
              >
                <rect
                  v-for="bar in cell.sparkline.bars ?? []"
                  :key="bar.key"
                  :x="bar.x"
                  :y="bar.y"
                  :width="bar.width"
                  :height="bar.height"
                  :fill="cell.sparkline.color"
                />
                <polyline
                  v-if="cell.sparkline.line"
                  :points="cell.sparkline.line"
                  fill="none"
                  :stroke="cell.sparkline.color"
                  stroke-width="1.5"
                />
                <circle
                  v-for="marker in cell.sparkline.markers ?? []"
                  :key="marker.key"
                  :cx="marker.cx"
                  :cy="marker.cy"
                  r="2"
                  :fill="cell.sparkline.color"
                />
              </svg>
              <span
                v-if="cell.selectionStyle"
                class="xlsx-grid__dom-selection"
                aria-hidden="true"
                :style="cell.selectionStyle"
              />
              <span
                v-if="cell.icon"
                class="xlsx-grid__dom-icon"
                data-testid="xlsx-grid-dom-icon"
                aria-hidden="true"
                :style="cell.iconStyle"
              >{{ cell.icon }}</span>
              <span
                v-if="cell.text !== undefined"
                class="xlsx-grid__dom-text"
                :style="cell.textStyle"
              ><template v-if="cell.runs"><span
                  v-for="run in cell.runs"
                  :key="run.key"
                  :style="run.style"
                >{{ run.text }}</span></template><template v-else>{{ cell.text }}</template></span>
              <span
                v-if="cell.comment"
                class="xlsx-grid__dom-comment"
                aria-hidden="true"
              />
            </div>
          </div>
          <div
            v-for="outline in pane.outlines"
            :key="outline.key"
            class="xlsx-grid__dom-outline"
            aria-hidden="true"
            :style="outline.style"
          />
        </div>
      </div>
      <div
        v-for="divider in model.dividers"
        :key="divider.key"
        class="xlsx-grid__dom-divider"
        aria-hidden="true"
        :style="divider.style"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import {
  GRID_HEADER_HEIGHT_PX,
  GRID_ROW_HEADER_WIDTH_PX,
} from "./grid-metrics";
import type {
  GridDomLayerModel,
  GridDomPaneModel,
} from "./grid-dom";

/**
 * DOM fallback renderer for the grid body.
 *
 * Every visible cell is a real `role="gridcell"` with the text as a text node,
 * so browser find, native text selection and the accessibility tree work while
 * the canvas renderer is unavailable or debugging needs the structure. The grid
 * computes the model; this component only places it, using the same pane
 * geometry the canvas draws with: frozen strips stay pinned, the scrolling
 * panes are translated by the scroll offset, and the pane boxes clip.
 */
const props = defineProps<{
  model: GridDomLayerModel;
  scrollLeft: number;
  scrollTop: number;
}>();

const HEADER_HEIGHT = GRID_HEADER_HEIGHT_PX;
const ROW_HEADER_WIDTH = GRID_ROW_HEADER_WIDTH_PX;

const rootStyle: CSSProperties = {
  inset: "0",
  pointerEvents: "none",
  position: "absolute",
};

const columnHeaderStripStyle: CSSProperties = {
  height: `${HEADER_HEIGHT}px`,
  left: `${ROW_HEADER_WIDTH}px`,
  overflow: "hidden",
  position: "absolute",
  top: "0px",
  width: `calc(100% - ${ROW_HEADER_WIDTH}px)`,
};

const rowHeaderStripStyle: CSSProperties = {
  height: `calc(100% - ${HEADER_HEIGHT}px)`,
  left: "0px",
  overflow: "hidden",
  position: "absolute",
  top: `${HEADER_HEIGHT}px`,
  width: `${ROW_HEADER_WIDTH}px`,
};

const bodyStyle: CSSProperties = {
  height: `calc(100% - ${HEADER_HEIGHT}px)`,
  left: `${ROW_HEADER_WIDTH}px`,
  overflow: "hidden",
  position: "absolute",
  top: `${HEADER_HEIGHT}px`,
  width: `calc(100% - ${ROW_HEADER_WIDTH}px)`,
};

const columnHeaderScrollStyle = computed<CSSProperties>(() => ({
  inset: "0",
  position: "absolute",
  transform: `translateX(${-props.scrollLeft}px)`,
}));

const rowHeaderScrollStyle = computed<CSSProperties>(() => ({
  inset: "0",
  position: "absolute",
  transform: `translateY(${-props.scrollTop}px)`,
}));

function paneContentStyle(pane: GridDomPaneModel): CSSProperties {
  return {
    inset: "0",
    position: "absolute",
    transform: `translate(${pane.scrollX ? -props.scrollLeft : 0}px, ${pane.scrollY ? -props.scrollTop : 0}px)`,
  };
}
</script>

<style scoped>
.xlsx-grid__dom {
  /* The canvas lives at the same z-index; overlays (editor, popovers, fill
     handle) are later siblings with their own z-index. */
  z-index: 0;
}
.xlsx-grid__dom .xlsx-grid__dom-strip,
.xlsx-grid__dom .xlsx-grid__dom-body,
.xlsx-grid__dom .xlsx-grid__dom-pane {
  overflow: hidden;
}
.xlsx-grid__dom .xlsx-grid__dom-strip-frozen {
  inset: 0;
  pointer-events: none;
  position: absolute;
}
.xlsx-grid__dom .xlsx-grid__dom-header {
  align-items: center;
  box-sizing: border-box;
  display: flex;
  font-family: -apple-system, BlinkMacSystemFont, sans-serif;
  font-size: 11px;
  justify-content: center;
  line-height: 1;
  overflow: hidden;
  position: absolute;
  white-space: nowrap;
}
.xlsx-grid__dom .xlsx-grid__dom-pane {
  position: absolute;
}
.xlsx-grid__dom .xlsx-grid__dom-row {
  left: 0;
  position: absolute;
  width: 100%;
}
.xlsx-grid__dom .xlsx-grid__dom-cell {
  position: absolute;
  /* The canvas swallows pointer input; DOM cells own it so the browser can
     start a native text selection from the cell text. */
  pointer-events: auto;
  user-select: text;
  -webkit-user-select: text;
}
.xlsx-grid__dom .xlsx-grid__dom-text {
  pointer-events: auto;
  user-select: text;
  -webkit-user-select: text;
}
.xlsx-grid__dom .xlsx-grid__dom-data-bar,
.xlsx-grid__dom .xlsx-grid__dom-data-bar-axis,
.xlsx-grid__dom .xlsx-grid__dom-sparkline,
.xlsx-grid__dom .xlsx-grid__dom-icon,
.xlsx-grid__dom .xlsx-grid__dom-comment,
.xlsx-grid__dom .xlsx-grid__dom-outline,
.xlsx-grid__dom .xlsx-grid__dom-divider {
  pointer-events: none;
}
.xlsx-grid__dom .xlsx-grid__dom-sparkline {
  height: 100%;
  inset: 0;
  position: absolute;
  width: 100%;
  z-index: 0;
}
.xlsx-grid__dom .xlsx-grid__dom-icon {
  line-height: 1;
  position: absolute;
  text-align: center;
  z-index: 2;
}
.xlsx-grid__dom .xlsx-grid__dom-comment {
  border-left: 8px solid transparent;
  border-top: 8px solid #dc2626;
  height: 0;
  position: absolute;
  right: 1px;
  top: 1px;
  width: 0;
  z-index: 2;
}
.xlsx-grid__dom .xlsx-grid__dom-divider {
  position: absolute;
  z-index: 3;
}
</style>
