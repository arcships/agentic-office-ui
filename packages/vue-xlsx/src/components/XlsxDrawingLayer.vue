<template>
  <div
    class="xlsx-drawing-layer"
    data-testid="xlsx-drawing-layer"
    :data-shape-count="shapeItems.length"
    :data-form-control-count="controlItems.length"
  >
    <div
      v-for="item in shapeItems"
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
      v-for="item in controlItems"
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
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue"
import type { XlsxFormControl, XlsxImageAnchor, XlsxShape, XlsxSheetData, XlsxViewerController } from "@arcships/xlsx-core"
import { anchorToAbsoluteRect, emuToPixels } from "@arcships/xlsx-core"

const props = defineProps<{ controller: XlsxViewerController; scrollLeft?: number; scrollTop?: number }>()

function frozenExtent(sheet: XlsxSheetData | null, axis: "column" | "row", zoomScale: number): number {
  if (!sheet?.freezePanes) return 0
  const limit = axis === "column" ? sheet.freezePanes.col : sheet.freezePanes.row
  const indices = axis === "column" ? sheet.visibleCols : sheet.visibleRows
  const sizes = axis === "column" ? sheet.colWidths : sheet.rowHeights
  const fallback = axis === "column" ? sheet.defaultColWidthPx : sheet.defaultRowHeightPx
  const zoom = Math.max(.1, zoomScale / 100)
  return indices.reduce((sum, actual, display) => actual < limit ? sum + (sizes[display] || fallback) * zoom : sum, 0)
}

function itemStyle(anchor: XlsxImageAnchor, zIndex: number, extra?: CSSProperties): CSSProperties {
  const sheet = props.controller.activeSheet
  const zoom = Math.max(.1, props.controller.zoomScale / 100)
  const rect = anchorToAbsoluteRect(anchor, sheet)
  const left = emuToPixels(rect.x) * zoom
  const top = emuToPixels(rect.y) * zoom
  const frozenWidth = frozenExtent(sheet, "column", props.controller.zoomScale)
  const frozenHeight = frozenExtent(sheet, "row", props.controller.zoomScale)
  return {
    left: `${left + 48 - (left < frozenWidth ? 0 : props.scrollLeft ?? 0)}px`,
    top: `${top + 24 - (top < frozenHeight ? 0 : props.scrollTop ?? 0)}px`,
    width: `${Math.max(1, emuToPixels(rect.cx) * zoom)}px`,
    height: `${Math.max(1, emuToPixels(rect.cy) * zoom)}px`,
    zIndex: Math.max(1, zIndex),
    ...extra,
  }
}

const shapeItems = computed(() => (props.controller.shapes as XlsxShape[])
  .filter((shape) => !shape.hidden)
  .map((shape) => ({
    shape,
    style: itemStyle(shape.anchor, shape.zIndex, {
      "--shape-fill": shape.fill?.none ? "transparent" : shape.fill?.color || "#dbeafe",
      "--shape-stroke": shape.stroke?.none ? "transparent" : shape.stroke?.color || "#2563eb",
      "--shape-stroke-width": `${shape.stroke?.widthPx || 1}px`,
      transform: shape.rotationDeg ? `rotate(${shape.rotationDeg}deg)` : undefined,
    } as CSSProperties),
  })))

const controlItems = computed(() => (props.controller.formControls as XlsxFormControl[])
  .filter((control) => !control.hidden)
  .map((control) => ({ control, style: itemStyle(control.anchor, control.zIndex) })))

function shapeText(shape: XlsxShape): string { return shape.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n") }
function controlLabel(kind: XlsxFormControl["kind"]): string { return ({ dropdown: "下拉框", editbox: "输入框", listbox: "列表框", spinner: "微调框", scrollbar: "滚动条", label: "标签", "group-box": "分组框", unknown: "控件" } as Record<string, string>)[kind] || kind }
</script>

<style scoped>
.xlsx-drawing-layer { height: 100%; left: 0; pointer-events: none; position: absolute; top: 0; width: 100%; }
.xlsx-drawing-layer__shape, .xlsx-drawing-layer__control { box-sizing: border-box; overflow: hidden; pointer-events: none; position: absolute; }
.xlsx-drawing-layer__shape svg { fill: var(--shape-fill); height: 100%; left: 0; position: absolute; stroke: var(--shape-stroke); stroke-width: var(--shape-stroke-width); top: 0; width: 100%; }
.xlsx-drawing-layer__text { align-items: center; display: flex; height: 100%; justify-content: center; padding: 6px; position: relative; text-align: center; white-space: pre-wrap; z-index: 1; }
.xlsx-drawing-layer__control { align-items: center; background: #f8fafc; border: 1px solid #64748b; color: #0f172a; display: flex; gap: 5px; padding: 3px 6px; }
.xlsx-drawing-layer__control button { height: 100%; width: 100%; }
.xlsx-drawing-layer__box, .xlsx-drawing-layer__radio { align-items: center; background: #fff; border: 1px solid #475569; display: inline-flex; flex: 0 0 14px; height: 14px; justify-content: center; }
.xlsx-drawing-layer__radio { border-radius: 50%; font-size: 9px; }
</style>
