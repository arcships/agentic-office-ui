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
        <button v-else-if="item.control.kind === 'button'" type="button">{{ item.control.label || item.control.name || '按钮' }}</button>
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
  type XlsxShape,
  type XlsxShapeParagraph,
  type XlsxShapeTextRun,
  type XlsxViewerController,
} from "@arcships/xlsx-core"
import { createGridPaneLayout, gridBodyLayerStyle, resolveAnchorRect, type GridPaneKey } from "./grid-metrics"

const props = defineProps<{ controller: XlsxViewerController; scrollLeft?: number; scrollTop?: number }>()

const layerStyle = gridBodyLayerStyle
const checkedOverrides = ref<Record<string, boolean>>({})

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
  .map((shape) => ({
    shape,
    ...itemPlacement(shape.anchor, shape.zIndex, {
      "--shape-fill": shape.fill?.none ? "transparent" : shape.fill?.color || "#dbeafe",
      "--shape-stroke": shape.stroke?.none ? "transparent" : shape.stroke?.color || "#2563eb",
      "--shape-stroke-width": `${shape.stroke?.widthPx || 1}px`,
      transform: shapeTransform(shape),
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

function controlLabel(kind: XlsxFormControl["kind"]): string { return ({ dropdown: "下拉框", editbox: "输入框", listbox: "列表框", spinner: "微调框", scrollbar: "滚动条", label: "标签", "group-box": "分组框", unknown: "控件" } as Record<string, string>)[kind] || kind }
</script>

<style scoped>
/* Geometry (inset past the headers) comes from `gridBodyLayerStyle`; each
   freeze pane clips its own drawings so a scrolled-away shape never paints
   over frozen panes, the headers, or the host UI around the surface. The layer
   is its own stacking context so pane/object z-index values cannot escape it. */
.xlsx-drawing-layer { pointer-events: none; position: absolute; isolation: isolate; z-index: 3; }
.xlsx-drawing-layer__pane { overflow: hidden; pointer-events: none; position: absolute; }
.xlsx-drawing-layer__shape {
  box-sizing: border-box; overflow: visible; pointer-events: none; position: absolute;
  transform-origin: center;
}
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
</style>
