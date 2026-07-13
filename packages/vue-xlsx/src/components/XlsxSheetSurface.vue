<template>
  <div
    ref="surfaceRef"
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
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch, type CSSProperties } from "vue"
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom"
import type { XlsxViewerController, XlsxCellAddress, XlsxCellStyleContext } from "@arcships/xlsx-core"
import XlsxGrid from "./XlsxGrid.vue"
import XlsxChartOverlay from "./XlsxChartOverlay.vue"
import XlsxImageLayer from "./XlsxImageLayer.vue"
import XlsxDrawingLayer from "./XlsxDrawingLayer.vue"
import XlsxSelectionOverlay from "./XlsxSelectionOverlay.vue"
import XlsxContextMenu from "./XlsxContextMenu.vue"
import XlsxChartsheetSurface from "./XlsxChartsheetSurface.vue"

const props = withDefaults(
  defineProps<{
    controller: XlsxViewerController
    getCellStyle?: ((cell: XlsxCellAddress, context?: XlsxCellStyleContext) => Partial<CSSProperties> | undefined) | null
    isDark?: boolean | null
    readOnly?: boolean
    selectionColor?: string
    selectionFillColor?: string
    showImages?: boolean
    /** Controlled zoom factor. 1 = 100%. */
    zoom?: number
    enableGestureZoom?: boolean
  }>(),
  {
    getCellStyle: null,
    isDark: false,
    readOnly: false,
    selectionColor: undefined,
    selectionFillColor: undefined,
    showImages: true,
    enableGestureZoom: true,
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
  "update:zoom": [zoom: number]
}>()

type GridInstance = InstanceType<typeof XlsxGrid>
type GridZoomAnchor = ReturnType<GridInstance["captureZoomAnchor"]>
type WebKitGestureEvent = Event & { clientX?: number; clientY?: number; scale?: number }

const surfaceRef = ref<HTMLElement | null>(null)
const gridRef = ref<GridInstance | null>(null)
const gridViewport = ref({ scrollLeft: 0, scrollTop: 0 })
const effectiveReadOnly = computed(() => props.controller.readOnly || props.readOnly)
const gridElement = computed<HTMLElement | null>(() => gridRef.value?.scrollContainer ?? null)

let pendingAnchor: { anchor: NonNullable<GridZoomAnchor>; requestedZoom: number; token: number } | undefined
let pendingZoom: number | undefined
let gestureStartZoom = 1
let gestureToken = 0
let webkitGestureActive = false

function controlledZoom(): number | undefined {
  return typeof props.zoom === "number" && Number.isFinite(props.zoom)
    ? clampSurfaceZoom(props.zoom)
    : undefined
}

function requestGestureZoom(nextZoom: number, clientX: number, clientY: number): void {
  const current = controlledZoom()
  if (current === undefined || nextZoom === (pendingZoom ?? current)) return
  const token = ++gestureToken
  const anchor = gridRef.value?.captureZoomAnchor(clientX, clientY)
  pendingZoom = nextZoom
  pendingAnchor = anchor ? { anchor, requestedZoom: nextZoom, token } : undefined
  emit("update:zoom", nextZoom)
}

function onWheel(event: WheelEvent): void {
  const current = controlledZoom()
  if (props.enableGestureZoom === false || current === undefined || !event.ctrlKey) return
  event.preventDefault()
  if (webkitGestureActive) return
  requestGestureZoom(
    nextSurfaceZoom(pendingZoom ?? current, event.deltaY, event.deltaMode),
    event.clientX,
    event.clientY,
  )
}

function onGestureStart(event: Event): void {
  const current = controlledZoom()
  if (props.enableGestureZoom === false || current === undefined) return
  event.preventDefault()
  webkitGestureActive = true
  gestureStartZoom = current
  pendingZoom = current
}

function onGestureChange(event: WebKitGestureEvent): void {
  if (!webkitGestureActive) return
  event.preventDefault()
  const rect = gridRef.value?.scrollContainer?.getBoundingClientRect()
  requestGestureZoom(
    clampSurfaceZoom(gestureStartZoom * (event.scale ?? 1)),
    event.clientX ?? (rect ? rect.left + rect.width / 2 : 0),
    event.clientY ?? (rect ? rect.top + rect.height / 2 : 0),
  )
}

function onGestureEnd(): void {
  webkitGestureActive = false
}

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

watch(() => props.zoom, async (next) => {
  if (typeof next !== "number" || !Number.isFinite(next)) return
  const normalized = clampSurfaceZoom(next)
  props.controller.setZoomScale(normalized * 100)
  const restore = pendingAnchor?.requestedZoom === normalized ? pendingAnchor : undefined
  pendingAnchor = undefined
  pendingZoom = undefined
  if (!restore) {
    gestureToken += 1
    return
  }
  await nextTick()
  requestAnimationFrame(() => {
    if (restore.token === gestureToken) gridRef.value?.restoreZoomAnchor(restore.anchor)
  })
}, { immediate: true })

watch(() => props.controller.activeTabIndex, () => {
  gestureToken += 1
  pendingAnchor = undefined
  pendingZoom = undefined
  const current = controlledZoom()
  if (current !== undefined) props.controller.setZoomScale(current * 100)
})

watch(
  () => props.controller.selection,
  (sel) => {
    if (!sel) { emit("selectionChange", { kind: "none" }); return }
    const range = { start: { row: sel.start.row, col: sel.start.col }, end: { row: sel.end.row, col: sel.end.col } }
    emit("selectionChange", { kind: "range", range })
  },
  { immediate: true },
)

onMounted(() => {
  const element = surfaceRef.value
  if (!element) return
  element.addEventListener("wheel", onWheel, { passive: false })
  element.addEventListener("gesturestart", onGestureStart, { passive: false })
  element.addEventListener("gesturechange", onGestureChange as EventListener, { passive: false })
  element.addEventListener("gestureend", onGestureEnd)
})

onBeforeUnmount(() => {
  gestureToken += 1
  const element = surfaceRef.value
  element?.removeEventListener("wheel", onWheel)
  element?.removeEventListener("gesturestart", onGestureStart)
  element?.removeEventListener("gesturechange", onGestureChange as EventListener)
  element?.removeEventListener("gestureend", onGestureEnd)
})
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
