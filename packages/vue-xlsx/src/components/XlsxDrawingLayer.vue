<template>
  <div
    class="xlsx-drawing-layer"
    data-testid="xlsx-drawing-layer"
    :data-shape-count="shapeItems.length"
    :data-form-control-count="controlItems.length"
    :style="layerStyle"
  >
    <div
      v-for="pane in panes"
      :key="pane.key"
      class="xlsx-drawing-layer__pane"
      :data-grid-pane="pane.key"
      :style="pane.style"
    >
      <div
        v-for="item in pane.shapes"
        :key="item.shape.id"
        class="xlsx-drawing-layer__shape"
        data-testid="xlsx-shape"
        :data-shape-geometry="item.shape.geometry"
        :style="item.style"
      >
        <svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
          <ellipse v-if="/ellipse|oval/i.test(item.shape.geometry)" cx="50" cy="50" rx="48" ry="48" />
          <path v-else-if="item.shape.svgPath" :d="item.shape.svgPath" :viewBox="item.shape.svgViewBox ? `0 0 ${item.shape.svgViewBox.width} ${item.shape.svgViewBox.height}` : undefined" />
          <rect v-else x="1" y="1" width="98" height="98" rx="4" />
        </svg>
        <div class="xlsx-drawing-layer__text">{{ shapeText(item.shape) }}</div>
      </div>
      <div
        v-for="item in pane.controls"
        :key="item.control.id"
        class="xlsx-drawing-layer__control"
        data-testid="xlsx-form-control"
        :data-control-kind="item.control.kind"
        :style="item.style"
      >
        <span v-if="item.control.kind === 'checkbox'" class="xlsx-drawing-layer__box">{{ item.control.checked ? '✓' : '' }}</span>
        <span v-else-if="item.control.kind === 'radio'" class="xlsx-drawing-layer__radio">{{ item.control.checked ? '●' : '' }}</span>
        <button v-else-if="item.control.kind === 'button'" type="button" disabled>{{ item.control.label || item.control.name || '按钮' }}</button>
        <span v-else>{{ item.control.label || item.control.name || controlLabel(item.control.kind) }}</span>
        <span v-if="item.control.kind === 'checkbox' || item.control.kind === 'radio'">{{ item.control.label || item.control.name }}</span>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue"
import type { XlsxFormControl, XlsxImageAnchor, XlsxShape, XlsxViewerController } from "@arcships/xlsx-core"
import { createGridPaneLayout, gridBodyLayerStyle, resolveAnchorRect, type GridPaneKey } from "./grid-metrics"

const props = defineProps<{ controller: XlsxViewerController; scrollLeft?: number; scrollTop?: number }>()

const layerStyle = gridBodyLayerStyle

const paneLayout = computed(() => createGridPaneLayout(
  props.controller.activeSheet,
  props.controller.zoomScale,
  props.scrollLeft ?? 0,
  props.scrollTop ?? 0,
))

function itemPlacement(anchor: XlsxImageAnchor, zIndex: number, extra?: CSSProperties): { pane: GridPaneKey; style: CSSProperties } {
  const rect = resolveAnchorRect(anchor, props.controller.activeSheet, props.controller.zoomScale)
  const placement = paneLayout.value.place(rect)
  return {
    pane: placement.pane,
    style: { ...placement.style, zIndex: Math.max(1, zIndex), ...extra },
  }
}

const shapeItems = computed(() => (props.controller.shapes as XlsxShape[])
  .filter((shape) => !shape.hidden)
  .map((shape) => ({
    shape,
    ...itemPlacement(shape.anchor, shape.zIndex, {
      "--shape-fill": shape.fill?.none ? "transparent" : shape.fill?.color || "#dbeafe",
      "--shape-stroke": shape.stroke?.none ? "transparent" : shape.stroke?.color || "#2563eb",
      "--shape-stroke-width": `${shape.stroke?.widthPx || 1}px`,
      transform: shape.rotationDeg ? `rotate(${shape.rotationDeg}deg)` : undefined,
    } as CSSProperties),
  })))

const controlItems = computed(() => (props.controller.formControls as XlsxFormControl[])
  .filter((control) => !control.hidden)
  .map((control) => ({ control, ...itemPlacement(control.anchor, control.zIndex) })))

const panes = computed(() => paneLayout.value.panes.map((pane) => ({
  ...pane,
  shapes: shapeItems.value.filter((item) => item.pane === pane.key),
  controls: controlItems.value.filter((item) => item.pane === pane.key),
})))

function shapeText(shape: XlsxShape): string { return shape.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n") }
function controlLabel(kind: XlsxFormControl["kind"]): string { return ({ dropdown: "下拉框", editbox: "输入框", listbox: "列表框", spinner: "微调框", scrollbar: "滚动条", label: "标签", "group-box": "分组框", unknown: "控件" } as Record<string, string>)[kind] || kind }
</script>

<style scoped>
/* Geometry (inset past the headers) comes from `gridBodyLayerStyle`; each
   freeze pane clips its own drawings so a scrolled-away shape never paints
   over frozen panes, the headers, or the host UI around the surface. */
.xlsx-drawing-layer { pointer-events: none; position: absolute; }
.xlsx-drawing-layer__pane { overflow: hidden; pointer-events: none; position: absolute; }
.xlsx-drawing-layer__shape, .xlsx-drawing-layer__control { box-sizing: border-box; overflow: hidden; pointer-events: none; position: absolute; }
.xlsx-drawing-layer__shape svg { fill: var(--shape-fill); height: 100%; left: 0; position: absolute; stroke: var(--shape-stroke); stroke-width: var(--shape-stroke-width); top: 0; width: 100%; }
.xlsx-drawing-layer__text { align-items: center; display: flex; height: 100%; justify-content: center; padding: 6px; position: relative; text-align: center; white-space: pre-wrap; z-index: 1; }
.xlsx-drawing-layer__control { align-items: center; background: #f8fafc; border: 1px solid #64748b; color: #0f172a; display: flex; gap: 5px; padding: 3px 6px; }
.xlsx-drawing-layer__control button { height: 100%; width: 100%; }
.xlsx-drawing-layer__box, .xlsx-drawing-layer__radio { align-items: center; background: #fff; border: 1px solid #475569; display: inline-flex; flex: 0 0 14px; height: 14px; justify-content: center; }
.xlsx-drawing-layer__radio { border-radius: 50%; font-size: 9px; }
</style>
