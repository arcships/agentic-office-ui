<template>
  <div
    class="xlsx-sheet-surface"
    data-testid="xlsx-sheet-surface"
    @keydown="onKeydown"
    @contextmenu="onContextMenu"
    tabindex="0"
  >
    <XlsxChartsheetSurface
      v-if="controller.activeTab?.kind === 'chartsheet'"
      :controller="controller"
      :is-dark="isDark ?? false"
    />
    <template v-else-if="controller.activeSheet">
      <XlsxGrid
        ref="gridRef"
        :controller="controller"
        :get-cell-style="getCellStyle"
        :is-dark="isDark ?? false"
        :read-only="effectiveReadOnly"
        :selection-color="selectionColor"
        :selection-fill-color="selectionFillColor"
        @cell-double-click="emit('cellDoubleClick', $event)"
        @viewport-change="gridViewport = $event"
      />
      <XlsxChartOverlay
        :controller="controller"
        :is-dark="isDark ?? false"
        :scroll-left="gridViewport.scrollLeft"
        :scroll-top="gridViewport.scrollTop"
      />
      <XlsxImageLayer
        :controller="controller"
        :show-images="showImages ?? true"
        :scroll-left="gridViewport.scrollLeft"
        :scroll-top="gridViewport.scrollTop"
      />
      <XlsxDrawingLayer
        :controller="controller"
        :scroll-left="gridViewport.scrollLeft"
        :scroll-top="gridViewport.scrollTop"
      />
      <XlsxSelectionOverlay
        :controller="controller"
        :get-cell-style="getCellStyle"
        :selection-color="selectionColor"
        :selection-fill-color="selectionFillColor"
      />
      <XlsxContextMenu
        :controller="controller"
        :target-element="gridElement"
      />
    </template>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch, type ComponentPublicInstance, type CSSProperties } from "vue"
import type { XlsxViewerController, XlsxCellAddress, XlsxCellStyleContext } from "@arcships/xlsx-core"
import XlsxGrid from "./XlsxGrid.vue"
import XlsxChartOverlay from "./XlsxChartOverlay.vue"
import XlsxImageLayer from "./XlsxImageLayer.vue"
import XlsxDrawingLayer from "./XlsxDrawingLayer.vue"
import XlsxSelectionOverlay from "./XlsxSelectionOverlay.vue"
import XlsxContextMenu from "./XlsxContextMenu.vue"
import XlsxChartsheetSurface from "./XlsxChartsheetSurface.vue"

// ── Props ────────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    /** Viewer controller from useXlsxViewerController. Required. */
    controller: XlsxViewerController
    /** Optional per-cell style override. */
    getCellStyle?: ((cell: XlsxCellAddress, context?: XlsxCellStyleContext) => Partial<CSSProperties> | undefined) | null
    /** Dark mode for the sheet surface. */
    isDark?: boolean | null
    /** Read-only mode — disables editing and context-menu actions. */
    readOnly?: boolean
    /** Selection border color. */
    selectionColor?: string
    /** Selection fill overlay color. */
    selectionFillColor?: string
    /** Show floating images imported from the workbook. */
    showImages?: boolean
  }>(),
  {
    getCellStyle: null,
    isDark: false,
    readOnly: false,
    selectionColor: undefined,
    selectionFillColor: undefined,
    showImages: true,
  },
)

const emit = defineEmits<{
  cellDoubleClick: [cell: XlsxCellAddress]
  contextMenu: [ctx: {
    clientX: number; clientY: number
    containerX: number; containerY: number
    sheetName?: string
    selection?: { start: { row: number; col: number }; end: { row: number; col: number } }
    activeCell?: { row: number; col: number }
  }]
  selectionChange: [sel: { kind: string; range?: { start: { row: number; col: number }; end: { row: number; col: number } }; value?: string }]
  objectClick: [obj: { kind: "chart" | "image" | "shape"; id: string }]
}>()

// ── State ────────────────────────────────────────────────────────────
const gridRef = ref<ComponentPublicInstance | null>(null)
const gridViewport = ref({ scrollLeft: 0, scrollTop: 0 })

const effectiveReadOnly = computed(() => props.controller.readOnly || props.readOnly)

const gridElement = computed<HTMLElement | null>(() => {
  const element = gridRef.value?.$el
  return typeof HTMLElement !== "undefined" && element instanceof HTMLElement ? element : null
})

function onContextMenu(event: MouseEvent): void {
  const sel = props.controller.selection
  const active = props.controller.activeCell
  const container = event.currentTarget as HTMLElement
  const rect = container.getBoundingClientRect()
  emit("contextMenu", {
    clientX: event.clientX,
    clientY: event.clientY,
    containerX: event.clientX - rect.left,
    containerY: event.clientY - rect.top,
    sheetName: props.controller.activeSheet?.name,
    selection: sel ? { start: { row: sel.start.row, col: sel.start.col }, end: { row: sel.end.row, col: sel.end.col } } : undefined,
    activeCell: active ? { row: active.row, col: active.col } : undefined,
  })
}

// ── Keyboard shortcuts ───────────────────────────────────────────────
function onKeydown(event: KeyboardEvent) {
  if (!props.controller || effectiveReadOnly.value) return
  if ((event.ctrlKey || event.metaKey) && event.key === "z") {
    event.preventDefault()
    props.controller.undo()
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key === "y") {
    event.preventDefault()
    props.controller.redo()
  }
}

// ── Selection tracking ───────────────────────────────────────────────
watch(
  () => props.controller.selection,
  (sel) => {
    if (!sel) { emit("selectionChange", { kind: "none" }); return }
    const range = { start: { row: sel.start.row, col: sel.start.col }, end: { row: sel.end.row, col: sel.end.col } }
    emit("selectionChange", { kind: "range", range })
  },
  { immediate: true },
)
</script>

<style scoped>
.xlsx-sheet-surface {
  display: flex;
  flex: 1;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  font-size: 12px;
  min-height: 0;
  min-width: 0;
  outline: none;
  position: relative;
  background: var(--xlsx-surface-bg, #ffffff);
}
</style>
