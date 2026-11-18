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
        :data-shape-id="item.shape.id"
        :data-shape-geometry="item.shape.geometry"
        tabindex="0"
        :class="{
          'xlsx-drawing-layer__shape--selected': item.shape.id === selectedShapeId,
          'xlsx-drawing-layer__shape--gesture': interaction.isActive(item.shape.id),
        }"
        :style="item.style"
        @pointerdown.stop="onShapePointerDown(item, $event)"
        @contextmenu.prevent.stop="onShapeContextMenu(item, $event)"
        @keydown="onShapeKeydown(item.shape, $event)"
      >
        <svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
          <ellipse v-if="/ellipse|oval/i.test(item.shape.geometry)" cx="50" cy="50" rx="48" ry="48" />
          <path v-else-if="item.shape.svgPath" :d="item.shape.svgPath" />
          <rect v-else x="1" y="1" width="98" height="98" rx="4" />
        </svg>
        <div class="xlsx-drawing-layer__text" :style="shapeTextStyle(item.shape)">
          <p
            v-for="(paragraph, paragraphIndex) in item.shape.paragraphs"
            :key="paragraphIndex"
            class="xlsx-drawing-layer__paragraph"
            :style="shapeParagraphStyle(item.shape, paragraph)"
          >
            <span
              v-for="(run, runIndex) in paragraph.runs"
              :key="runIndex"
              :style="shapeRunStyle(run)"
            >{{ run.text }}</span>
          </p>
        </div>
        <div
          v-if="item.shape.id === selectedShapeId"
          class="xlsx-drawing-layer__handles"
          data-testid="xlsx-shape-handles"
        >
          <div
            v-for="handle in DRAWING_RESIZE_HANDLES"
            :key="handle"
            class="xlsx-drawing-layer__handle"
            data-testid="xlsx-shape-resize-handle"
            :data-handle="handle"
            :style="handleStyle(handle)"
            @pointerdown.stop="onShapeHandlePointerDown(item, handle, $event)"
          />
        </div>
      </div>
      <div
        v-for="item in pane.controls"
        :key="item.control.id"
        class="xlsx-drawing-layer__control"
        data-testid="xlsx-form-control"
        :data-control-kind="item.control.kind"
        :data-control-checked="isChecked(item.control) ? 'true' : 'false'"
        :style="[item.style, controlTextStyle(item.control)]"
        :role="controlRole(item.control)"
        :tabindex="controlRole(item.control) ? 0 : undefined"
        :aria-checked="item.control.kind === 'checkbox' || item.control.kind === 'radio' ? isChecked(item.control) : undefined"
        @click.stop="onControlActivate(item.control)"
        @keydown.enter.stop.prevent="onControlActivate(item.control)"
        @keydown.space.stop.prevent="onControlActivate(item.control)"
      >
        <span v-if="item.control.kind === 'checkbox'" class="xlsx-drawing-layer__box">{{ isChecked(item.control) ? '✓' : '' }}</span>
        <span v-else-if="item.control.kind === 'radio'" class="xlsx-drawing-layer__radio">{{ isChecked(item.control) ? '●' : '' }}</span>
        <button v-else-if="item.control.kind === 'button'" type="button">{{ item.control.label || item.control.name || m.xlsx.drawing.button }}</button>
        <span v-else>{{ item.control.label || item.control.name || controlLabel(item.control.kind) }}</span>
        <span v-if="item.control.kind === 'checkbox' || item.control.kind === 'radio'">{{ item.control.label || item.control.name }}</span>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, type CSSProperties } from "vue"
import {
  parseA1CellReference,
  type XlsxFormControl,
  type XlsxImageAnchor,
  type XlsxImageRect,
  type XlsxImageResizeHandlePosition,
  type XlsxShape,
  type XlsxShapeParagraph,
  type XlsxShapeTextRun,
  type XlsxViewerController,
} from "@arcships/xlsx-core"
import { createGridPaneLayout, gridBodyLayerStyle, resolveAnchorRect, type GridPaneKey } from "./grid-metrics"
import { DRAWING_RESIZE_HANDLES, drawingHandleStyle, useDrawingGestures } from "./drawing-interaction"
import { useXlsxMessages } from "../i18n"

const props = defineProps<{
  controller: XlsxViewerController
  scrollLeft?: number
  scrollTop?: number
  viewportWidth?: number
  viewportHeight?: number
}>()

const emit = defineEmits<{
  drawingContextMenu: [payload: { id: string; kind: "shape"; clientX: number; clientY: number }]
}>()

const layerStyle = gridBodyLayerStyle
const checkedOverrides = ref<Record<string, boolean>>({})
const selectedShapeId = computed(() => props.controller.selectedShapeId)
const m = useXlsxMessages()

const paneLayout = computed(() => createGridPaneLayout(
  props.controller.activeSheet,
  props.controller.zoomScale,
  props.scrollLeft ?? 0,
  props.scrollTop ?? 0,
))

const interaction = useDrawingGestures({
  onSelect: (id) => props.controller.selectShape(id),
  onMove: (id, deltaX, deltaY) => props.controller.moveShapeBy(id, deltaX, deltaY),
  onResize: (id, handle, deltaX, deltaY) => props.controller.resizeShapeBy(id, handle, deltaX, deltaY),
})

function numericPx(value: string | number | undefined): number {
  return typeof value === "number" ? value : Number.parseFloat(value ?? "0") || 0
}

function handleStyle(handle: XlsxImageResizeHandlePosition): CSSProperties {
  return drawingHandleStyle(handle)
}

function itemPlacement(anchor: XlsxImageAnchor, zIndex: number, extra?: CSSProperties): { pane: GridPaneKey; rect: XlsxImageRect; style: CSSProperties } {
  const rect = resolveAnchorRect(anchor, props.controller.activeSheet, props.controller.zoomScale)
  const placement = paneLayout.value.place(rect)
  return {
    pane: placement.pane,
    rect,
    style: { ...placement.style, zIndex: Math.max(0, zIndex), ...extra },
  }
}

function shapeTransform(shape: XlsxShape): string | undefined {
  const parts: string[] = []
  if (shape.rotationDeg) parts.push(`rotate(${shape.rotationDeg}deg)`)
  if (shape.flipH) parts.push("scaleX(-1)")
  if (shape.flipV) parts.push("scaleY(-1)")
  return parts.length ? parts.join(" ") : undefined
}

const shapeItems = computed(() => (props.controller.shapes as XlsxShape[])
  .filter((shape) => !shape.hidden)
  .flatMap((shape) => {
    const item = itemPlacement(shape.anchor, shape.zIndex, {
      "--shape-fill": shape.fill?.none ? "transparent" : shape.fill?.color || "#dbeafe",
      "--shape-stroke": shape.stroke?.none ? "transparent" : shape.stroke?.color || "#2563eb",
      "--shape-stroke-width": `${shape.stroke?.widthPx || 1}px`,
      transform: shapeTransform(shape),
    } as CSSProperties)
    if (!paneLayout.value.isVisible(item.rect, item.pane, props.viewportWidth ?? 0, props.viewportHeight ?? 0)) {
      return []
    }
    const isSelected = shape.id === selectedShapeId.value
    const preview = interaction.previewFor(shape.id)
    const style: CSSProperties = {
      ...item.style,
      cursor: isSelected ? "move" : "default",
      outline: isSelected ? "2px solid #2563eb" : "none",
      outlineOffset: "-1px",
      pointerEvents: "auto",
    }
    if (preview) {
      style.left = `${numericPx(item.style.left) + preview.left - item.rect.left}px`
      style.top = `${numericPx(item.style.top) + preview.top - item.rect.top}px`
      style.width = `${preview.width}px`
      style.height = `${preview.height}px`
    }
    return [{ shape, pane: item.pane, rect: item.rect, style }]
  }))

const controlItems = computed(() => (props.controller.formControls as XlsxFormControl[])
  .filter((control) => !control.hidden)
  .flatMap((control) => {
    const item = itemPlacement(control.anchor, control.zIndex)
    if (!paneLayout.value.isVisible(item.rect, item.pane, props.viewportWidth ?? 0, props.viewportHeight ?? 0)) {
      return []
    }
    return [{ control, pane: item.pane, rect: item.rect, style: item.style }]
  }))

const panes = computed(() => paneLayout.value.panes.map((pane) => ({
  ...pane,
  shapes: shapeItems.value.filter((item) => item.pane === pane.key),
  controls: controlItems.value.filter((item) => item.pane === pane.key),
})))

function onShapePointerDown(item: { shape: XlsxShape; rect: XlsxImageRect }, event: PointerEvent) {
  interaction.beginMove(item.shape.id, item.rect, event)
}

function onShapeContextMenu(item: { shape: XlsxShape; rect: XlsxImageRect }, event: MouseEvent) {
  props.controller.selectShape(item.shape.id)
  emit("drawingContextMenu", {
    id: item.shape.id,
    kind: "shape",
    clientX: event.clientX,
    clientY: event.clientY,
  })
}

function onShapeHandlePointerDown(
  item: { shape: XlsxShape; rect: XlsxImageRect },
  handle: XlsxImageResizeHandlePosition,
  event: PointerEvent,
) {
  interaction.beginResize(item.shape.id, item.rect, handle, event)
}

function onShapeKeydown(shape: XlsxShape, event: KeyboardEvent) {
  if (props.controller.selectedShapeId !== shape.id) {
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
    event.preventDefault()
    event.stopPropagation()
    void props.controller.copySelectionToClipboard?.()
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "x") {
    event.preventDefault()
    event.stopPropagation()
    props.controller.cutSelection?.()
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
    event.preventDefault()
    event.stopPropagation()
    void props.controller.pasteFromClipboard?.()
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
    event.preventDefault()
    event.stopPropagation()
    props.controller.duplicateDrawing?.()
    return
  }
  const step = event.shiftKey ? 10 : 1
  const nudge: Record<string, [number, number]> = {
    ArrowLeft: [-1, 0],
    ArrowRight: [1, 0],
    ArrowUp: [0, -1],
    ArrowDown: [0, 1],
  }
  const delta = nudge[event.key]
  if (delta) {
    event.preventDefault()
    event.stopPropagation()
    props.controller.moveShapeBy(shape.id, delta[0] * step, delta[1] * step)
    return
  }
  if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault()
    event.stopPropagation()
    props.controller.deleteShape(shape.id)
    return
  }
  if (event.key === "Escape") {
    event.stopPropagation()
    if (props.controller.cancelCut?.()) return
    props.controller.clearSelectedShape()
  }
}

function shapeTextStyle(shape: XlsxShape): CSSProperties {
  const insets = shape.textBox?.insetPx
  const align = shape.textBox?.verticalAlign
  return {
    justifyContent: align === "top" ? "flex-start" : align === "bottom" ? "flex-end" : "center",
    padding: insets
      ? `${insets.top}px ${insets.right}px ${insets.bottom}px ${insets.left}px`
      : "4.8px 9.6px",
  }
}

function shapeParagraphStyle(shape: XlsxShape, paragraph: XlsxShapeParagraph): CSSProperties {
  return {
    textAlign: paragraph.align ?? (shape.textBox?.horizontalAlign === "center" ? "center" : "left"),
  }
}

function shapeRunStyle(run: XlsxShapeTextRun): CSSProperties {
  return {
    color: run.color,
    fontFamily: run.fontFamily,
    fontSize: run.fontSizePt ? `${run.fontSizePt * (96 / 72)}px` : undefined,
    fontStyle: run.italic ? "italic" : undefined,
    fontWeight: run.bold ? "700" : undefined,
    textDecoration: run.underline ? "underline" : undefined,
  }
}

function isChecked(control: XlsxFormControl): boolean {
  return checkedOverrides.value[control.id] ?? Boolean(control.checked)
}

function controlRole(control: XlsxFormControl): string | undefined {
  if (control.kind === "checkbox") return "checkbox"
  if (control.kind === "radio") return "radio"
  if (control.kind === "button") return "button"
  return undefined
}

function controlTextStyle(control: XlsxFormControl): CSSProperties {
  return {
    color: control.textColor,
    fontFamily: control.fontFamily,
    fontSize: control.fontSizePt ? `${control.fontSizePt * (96 / 72)}px` : undefined,
    justifyContent: control.textAlign === "center"
      ? "center"
      : control.textAlign === "right"
        ? "flex-end"
        : undefined,
    textAlign: control.textAlign,
  }
}

/** Resolve `linkedCell` (e.g. `$B$4` or `Sheet1!$B$4`) and persist the toggle. */
function writeLinkedCell(control: XlsxFormControl, checked: boolean): void {
  const reference = control.linkedCell?.split("!").pop()?.trim().replace(/^=/, "")
  if (!reference) return
  const cell = parseA1CellReference(reference)
  if (!cell) return
  props.controller.setCellValue(cell, checked ? "TRUE" : "FALSE")
}

function onControlActivate(control: XlsxFormControl): void {
  if (props.controller.readOnly) return
  if (control.kind !== "checkbox" && control.kind !== "radio") return
  const next = control.kind === "radio" ? true : !isChecked(control)
  const overrides = { ...checkedOverrides.value }
  if (control.kind === "radio") {
    const groupKey = control.linkedCell ?? control.name
    for (const sibling of props.controller.formControls as XlsxFormControl[]) {
      if (
        sibling.id !== control.id
        && sibling.kind === "radio"
        && (sibling.linkedCell ?? sibling.name) === groupKey
      ) {
        overrides[sibling.id] = false
      }
    }
  }
  overrides[control.id] = next
  checkedOverrides.value = overrides
  writeLinkedCell(control, next)
}

function controlLabel(kind: XlsxFormControl["kind"]): string { return ({ dropdown: m.value.xlsx.drawing.dropdown, editbox: m.value.xlsx.drawing.editbox, listbox: m.value.xlsx.drawing.listbox, spinner: m.value.xlsx.drawing.spinner, scrollbar: m.value.xlsx.drawing.scrollbar, label: m.value.xlsx.drawing.label, "group-box": m.value.xlsx.drawing.groupBox, unknown: m.value.xlsx.drawing.unknown } as Record<string, string>)[kind] || kind }
</script>

<style scoped>
/* Geometry (inset past the headers) comes from `gridBodyLayerStyle`; each
   freeze pane clips its own drawings so a scrolled-away shape never paints
   over frozen panes, the headers, or the host UI around the surface. The layer
   is its own stacking context so pane/object z-index values cannot escape it. */
.xlsx-drawing-layer { pointer-events: none; position: absolute; isolation: isolate; z-index: 3; }
.xlsx-drawing-layer__pane { overflow: hidden; pointer-events: none; position: absolute; }
.xlsx-drawing-layer__shape {
  box-sizing: border-box; overflow: visible; pointer-events: auto; position: absolute;
  transform-origin: center;
}
.xlsx-drawing-layer__shape--gesture { cursor: move; }
.xlsx-drawing-layer__shape svg { fill: var(--shape-fill); height: 100%; left: 0; position: absolute; stroke: var(--shape-stroke); stroke-width: var(--shape-stroke-width); top: 0; width: 100%; }
.xlsx-drawing-layer__text {
  box-sizing: border-box; display: flex; flex-direction: column;
  height: 100%; overflow: visible; position: relative; white-space: pre-wrap; z-index: 1;
}
.xlsx-drawing-layer__paragraph { margin: 0; }
.xlsx-drawing-layer__control {
  align-items: center; background: #f8fafc; border: 1px solid #64748b; box-sizing: border-box;
  color: #0f172a; cursor: pointer; display: flex; gap: 5px; overflow: hidden;
  padding: 3px 6px; pointer-events: auto; position: absolute;
}
.xlsx-drawing-layer__control button { background: transparent; border: 0; color: inherit; cursor: pointer; font: inherit; height: 100%; width: 100%; }
.xlsx-drawing-layer__box, .xlsx-drawing-layer__radio { align-items: center; background: #fff; border: 1px solid #475569; display: inline-flex; flex: 0 0 14px; height: 14px; justify-content: center; }
.xlsx-drawing-layer__radio { border-radius: 50%; font-size: 9px; }
.xlsx-drawing-layer__handles { inset: 0; pointer-events: none; position: absolute; }
.xlsx-drawing-layer__handle {
  background: #ffffff; border: 1px solid #2563eb; box-sizing: border-box;
  height: 8px; pointer-events: auto; position: absolute; width: 8px;
  z-index: 1000000;
}
</style>
