<template>
  <div
    class="xlsx-sheet-surface"
    data-testid="xlsx-sheet-surface"
    @keydown="onKeydown"
    @contextmenu.prevent="onContextMenu"
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
import { computed, ref, type ComponentPublicInstance, type CSSProperties } from "vue"
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
  contextMenu: [ctx: { clientX: number; clientY: number; sheetName?: string }]
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
  emit("contextMenu", {
    clientX: event.clientX,
    clientY: event.clientY,
    sheetName: props.controller.activeSheet?.name,
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
