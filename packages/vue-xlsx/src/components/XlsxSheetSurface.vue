<template>
  <div
    ref="surfaceRef"
    class="xlsx-sheet-surface"
    data-testid="xlsx-sheet-surface"
    @keydown="onKeydown"
    @contextmenu="onContextMenu"
    @pointermove="onReferencePointerMove"
    @click="onReferenceClick"
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
    <div
      v-if="selectionMode === 'region'"
      class="xlsx-sheet-surface__region-layer"
      aria-label="选择工作表区域"
      tabindex="0"
      @pointerdown="onRegionPointerDown"
      @pointermove="onRegionPointerMove"
      @pointerup="onRegionPointerUp"
      @pointercancel="onRegionPointerCancel"
      @keydown.esc="onRegionKeyboardCancel"
    >
      <div v-if="regionFrame" class="xlsx-sheet-surface__region-frame" :style="regionFrame" />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch, type CSSProperties } from "vue"
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom"
import type { XlsxViewerController, XlsxCellAddress, XlsxCellStyleContext, XlsxOfficeReference, XlsxReferenceContext, XlsxReferenceSheet } from "@arcships/xlsx-core"
import {
  createXlsxCellReferenceDraft,
  createXlsxColumnReferenceDraft,
  createXlsxRangeReferenceDraft,
  createXlsxRegionReferenceDraft,
  createXlsxRowReferenceDraft,
  createXlsxWorksheetReferenceDraft,
  describeXlsxReference,
  resolveXlsxReference,
} from "@arcships/xlsx-core"
import {
  confirmOfficeReferenceDraft,
  createOfficeReferenceId,
  type OfficeDocumentRevision,
  type OfficeObjectReference,
  type OfficeReferenceCandidatePreview,
  type OfficeReferenceError,
  type OfficeSelectionMode,
  type SheetRegionPoint,
  type ResolveReferenceResult,
} from "@arcships/office-interaction"
import XlsxGrid from "./XlsxGrid.vue"
import XlsxChartOverlay from "./XlsxChartOverlay.vue"
import XlsxImageLayer from "./XlsxImageLayer.vue"
import XlsxDrawingLayer from "./XlsxDrawingLayer.vue"
import XlsxSelectionOverlay from "./XlsxSelectionOverlay.vue"
import XlsxContextMenu from "./XlsxContextMenu.vue"
import XlsxChartsheetSurface from "./XlsxChartsheetSurface.vue"
import { useXlsxSearch, type XlsxSearchState } from "../composables/useXlsxSearch"

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
    documentId?: string
    selectionMode?: OfficeSelectionMode
    emitReferenceCandidates?: boolean
  }>(),
  {
    getCellStyle: null,
    isDark: false,
    readOnly: false,
    selectionColor: undefined,
    selectionFillColor: undefined,
    showImages: true,
    enableGestureZoom: true,
    selectionMode: "content",
    emitReferenceCandidates: false,
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
  searchStateChange: [state: XlsxSearchState]
  documentRevisionChange: [revision: OfficeDocumentRevision]
  referenceCandidateChange: [change: { candidates: readonly OfficeReferenceCandidatePreview[]; activeCandidateId?: string }]
  referenceConfirm: [event: ReturnType<typeof confirmOfficeReferenceDraft>]
  regionDraftChange: [event: { phase: "start" | "change"; region: { space: "sheet"; sheetId: string; start: SheetRegionPoint; end: SheetRegionPoint } }]
  selectionCancel: [event: { mode: OfficeSelectionMode; reason: "escape" | "pointer-cancel" | "programmatic" }]
  referenceResolve: [event: { referenceId: string; result: ResolveReferenceResult }]
  referenceError: [error: OfficeReferenceError]
}>()

type GridInstance = InstanceType<typeof XlsxGrid>
type GridZoomAnchor = ReturnType<GridInstance["captureZoomAnchor"]>
type WebKitGestureEvent = Event & { clientX?: number; clientY?: number; scale?: number }

const surfaceRef = ref<HTMLElement | null>(null)
const gridRef = ref<GridInstance | null>(null)
const gridViewport = ref({ scrollLeft: 0, scrollTop: 0 })
const effectiveReadOnly = computed(() => props.controller.readOnly || props.readOnly)
const selectionMode = computed(() => props.selectionMode)
const gridElement = computed<HTMLElement | null>(() => gridRef.value?.scrollContainer ?? null)
const generatedDocumentId = `xlsx-surface-${typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2)}`
const documentRevision = computed<OfficeDocumentRevision & { format: "xlsx" }>(() => ({
  format: "xlsx",
  documentId: props.documentId?.trim() || generatedDocumentId,
  revision: `${generatedDocumentId}:${props.controller.revision}`,
}))

function referenceSheets(): XlsxReferenceSheet[] {
  return props.controller.sheets.map((sheet, index) => ({
    index,
    name: sheet.name,
    workbookSheetIndex: sheet.workbookSheetIndex,
  }))
}

function activeReferenceSheet(): XlsxReferenceSheet | undefined {
  return referenceContext().sheets[props.controller.activeSheetIndex]
}

let cachedReferenceContext: { key: string; context: XlsxReferenceContext } | undefined

function referenceContext(): XlsxReferenceContext {
  const key = `${documentRevision.value.documentId}:${documentRevision.value.revision}`
  if (cachedReferenceContext?.key === key) return cachedReferenceContext.context
  const context: XlsxReferenceContext = {
    revision: documentRevision.value,
    sheets: referenceSheets(),
    charts: props.controller.charts,
    getCellSnapshot(_sheet, cell) {
      return {
        displayValue: props.controller.getCellDisplayValue(cell),
        formula: props.controller.getCellFormula(cell),
      }
    },
  }
  cachedReferenceContext = { key, context }
  return context
}

function emitReferenceError(operation: OfficeReferenceError["operation"], code: OfficeReferenceError["code"], message: string, referenceId?: string): void {
  emit("referenceError", { code, operation, format: "xlsx", recoverable: true, ...(referenceId ? { referenceId } : {}), message })
}

function confirmDraft(
  draft: Parameters<typeof confirmOfficeReferenceDraft>[0],
  trigger: "pointer" | "keyboard" | "touch" | "programmatic",
  snapshot?: Parameters<typeof confirmOfficeReferenceDraft>[1]["snapshot"],
  additiveRequested = false,
): void {
  try {
    emit("referenceConfirm", confirmOfficeReferenceDraft(draft, {
      referenceId: createOfficeReferenceId(), trigger, additiveRequested, ...(snapshot ? { snapshot } : {}),
    }))
  } catch (reason) {
    emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : "Unable to confirm XLSX reference.")
  }
}

function worksheetPath(sheet: XlsxReferenceSheet) {
  return [{ kind: "workbook" as const, label: "Workbook" }, { kind: "worksheet" as const, label: sheet.name }]
}

function columnLabel(columnIndex: number): string {
  let label = ""
  let current = columnIndex
  while (current >= 0) {
    label = String.fromCharCode(65 + current % 26) + label
    current = Math.floor(current / 26) - 1
  }
  return label
}

function candidateForPoint(point: { clientX: number; clientY: number }): OfficeReferenceCandidatePreview[] {
  const sheet = activeReferenceSheet()
  if (!sheet) return []
  const axis = gridRef.value?.hitTestAxis(point.clientX, point.clientY)
  if (axis?.kind === "row") {
    const draft = createXlsxRowReferenceDraft(referenceContext(), sheet, axis.actualIndex)
    return [{ candidateId: `xlsx:row:${axis.actualIndex}`, draft, preview: { label: `Row ${axis.actualIndex + 1}`, path: [...worksheetPath(sheet), { kind: "row", label: `Row ${axis.actualIndex + 1}` }] }, hit: "direct", depth: 2 }]
  }
  if (axis?.kind === "column") {
    const label = columnLabel(axis.actualIndex)
    const draft = createXlsxColumnReferenceDraft(referenceContext(), sheet, axis.actualIndex)
    return [{ candidateId: `xlsx:column:${axis.actualIndex}`, draft, preview: { label: `Column ${label}`, path: [...worksheetPath(sheet), { kind: "column", label: `Column ${label}` }] }, hit: "direct", depth: 2 }]
  }
  if (axis?.kind === "corner") {
    const draft = createXlsxWorksheetReferenceDraft(referenceContext(), sheet)
    return [{ candidateId: `xlsx:worksheet:${sheet.index}`, draft, preview: { label: sheet.name, path: worksheetPath(sheet) }, hit: "direct", depth: 1 }]
  }
  const cell = gridRef.value?.hitTestCell(point.clientX, point.clientY)
  if (!cell) return []
  const draft = createXlsxCellReferenceDraft(referenceContext(), sheet, cell)
  const label = `${sheet.name}!R${cell.row + 1}C${cell.col + 1}`
  return [{ candidateId: `xlsx:cell:${cell.row}:${cell.col}`, draft, preview: { label, path: [...worksheetPath(sheet), { kind: "cell", label }] }, hit: "direct", depth: 2 }]
}

function hitTest(point: { clientX: number; clientY: number }): readonly OfficeReferenceCandidatePreview[] {
  try { return candidateForPoint(point) } catch (reason) {
    emitReferenceError("hit-test", "HIT_TEST_FAILED", reason instanceof Error ? reason.message : "XLSX hit test failed.")
    return []
  }
}

function onReferencePointerMove(event: PointerEvent): void {
  if (selectionMode.value !== "object" || !props.emitReferenceCandidates) return
  const candidates = hitTest(event)
  emit("referenceCandidateChange", { candidates, ...(candidates[0] ? { activeCandidateId: candidates[0].candidateId } : {}) })
}

function onReferenceClick(event: MouseEvent): void {
  if (selectionMode.value !== "object") return
  const candidate = hitTest(event)[0]
  if (candidate) confirmDraft(candidate.draft, event.detail === 0 ? "keyboard" : "pointer", candidate.preview, event.shiftKey)
}

type SheetRegion = { space: "sheet"; sheetId: string; start: SheetRegionPoint; end: SheetRegionPoint }
type RegionGesture = { pointerId: number; start: SheetRegionPoint; bounds: DOMRect; sheet: XlsxReferenceSheet }
const regionGesture = ref<RegionGesture | null>(null)
const regionDraft = ref<SheetRegion | null>(null)
const regionVisual = ref<{ startX: number; startY: number; endX: number; endY: number } | null>(null)
const regionFrame = computed(() => {
  const visual = regionVisual.value
  const bounds = surfaceRef.value?.getBoundingClientRect()
  if (!visual || !bounds || bounds.width <= 0 || bounds.height <= 0) return undefined
  const x1 = Math.min(visual.startX, visual.endX) - bounds.left
  const y1 = Math.min(visual.startY, visual.endY) - bounds.top
  return { left: `${x1}px`, top: `${y1}px`, width: `${Math.abs(visual.endX - visual.startX)}px`, height: `${Math.abs(visual.endY - visual.startY)}px` }
})

function orderedSheetRegion(sheet: XlsxReferenceSheet, left: SheetRegionPoint, right: SheetRegionPoint): SheetRegion {
  const before = left.row < right.row || left.row === right.row && left.col < right.col
    || left.row === right.row && left.col === right.col && (left.yOffset < right.yOffset || left.yOffset === right.yOffset && left.xOffset <= right.xOffset)
  const start = before ? left : right
  const end = before ? right : left
  return { space: "sheet", sheetId: sheet.id ?? sheet.name, start, end }
}

function onRegionPointerDown(event: PointerEvent): void {
  if (event.button !== 0) return
  const sheet = activeReferenceSheet()
  const start = gridRef.value?.hitTestSheetPoint(event.clientX, event.clientY)
  const bounds = surfaceRef.value?.getBoundingClientRect()
  if (!sheet || !start || !bounds) return
  regionGesture.value = { pointerId: event.pointerId, start, bounds, sheet }
  regionDraft.value = orderedSheetRegion(sheet, start, { ...start, xOffset: Math.min(1, start.xOffset + 0.0001), yOffset: Math.min(1, start.yOffset + 0.0001) })
  regionVisual.value = { startX: event.clientX, startY: event.clientY, endX: event.clientX, endY: event.clientY }
  ;(event.currentTarget as HTMLElement).setPointerCapture?.(event.pointerId)
  emit("regionDraftChange", { phase: "start", region: regionDraft.value })
  event.preventDefault()
}

function onRegionPointerMove(event: PointerEvent): void {
  const gesture = regionGesture.value
  if (!gesture || gesture.pointerId !== event.pointerId) return
  const point = gridRef.value?.hitTestSheetPoint(event.clientX, event.clientY)
  if (!point) return
  regionDraft.value = orderedSheetRegion(gesture.sheet, gesture.start, point)
  if (regionVisual.value) regionVisual.value = { ...regionVisual.value, endX: event.clientX, endY: event.clientY }
  emit("regionDraftChange", { phase: "change", region: regionDraft.value })
}

function finishRegion(event: PointerEvent, cancelled: boolean): void {
  const gesture = regionGesture.value
  if (!gesture || gesture.pointerId !== event.pointerId) return
  const target = event.currentTarget as HTMLElement
  if (target.hasPointerCapture?.(event.pointerId)) target.releasePointerCapture?.(event.pointerId)
  const region = regionDraft.value
  regionGesture.value = null
  regionDraft.value = null
  regionVisual.value = null
  if (cancelled || !region) {
    emit("selectionCancel", { mode: "region", reason: cancelled ? "pointer-cancel" : "programmatic" })
    return
  }
  confirmDraft(createXlsxRegionReferenceDraft(referenceContext(), gesture.sheet, region), event.pointerType === "touch" ? "touch" : "pointer")
}

function onRegionPointerUp(event: PointerEvent): void { finishRegion(event, false) }
function onRegionPointerCancel(event: PointerEvent): void { finishRegion(event, true) }
function onRegionKeyboardCancel(): void {
  regionGesture.value = null
  regionDraft.value = null
  regionVisual.value = null
  emit("selectionCancel", { mode: "region", reason: "escape" })
}
const surfaceSearch = useXlsxSearch(
  () => props.controller,
  {
    scrollToCell: async (cell) => {
      await nextTick()
      gridRef.value?.scrollToCell(cell)
    },
  },
)
const surfaceMounted = ref(false)

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
  if (event.key === "Escape" && selectionMode.value !== "content") {
    emit("selectionCancel", { mode: selectionMode.value, reason: "escape" })
  }
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
    if (!surfaceMounted.value || selectionMode.value !== "content") return
    const sheet = activeReferenceSheet()
    if (!sheet) return
    const single = range.start.row === range.end.row && range.start.col === range.end.col
    const draft = single
      ? createXlsxCellReferenceDraft(referenceContext(), sheet, range.start)
      : createXlsxRangeReferenceDraft(referenceContext(), sheet, range)
    const label = `${sheet.name}!${props.controller.selectedRangeAddress ?? "selection"}`
    confirmDraft(draft, "programmatic", {
      label,
      path: [...worksheetPath(sheet), { kind: single ? "cell" : "cell-range", label }],
      content: single ? { value: props.controller.selectedValue, formula: props.controller.selectedFormula } : undefined,
    })
  },
  { immediate: true },
)

watch(
  () => [props.controller.revision, props.documentId] as const,
  () => {
    cachedReferenceContext = undefined
    emit("documentRevisionChange", documentRevision.value)
  },
  { immediate: true },
)

watch(surfaceSearch.searchState, (next) => {
  emit("searchStateChange", next)
}, { immediate: true })

onMounted(() => {
  surfaceMounted.value = true
  const element = surfaceRef.value
  if (!element) return
  element.addEventListener("wheel", onWheel, { passive: false })
  element.addEventListener("gesturestart", onGestureStart, { passive: false })
  element.addEventListener("gesturechange", onGestureChange as EventListener, { passive: false })
  element.addEventListener("gestureend", onGestureEnd)
})

onBeforeUnmount(() => {
  surfaceMounted.value = false
  gestureToken += 1
  const element = surfaceRef.value
  element?.removeEventListener("wheel", onWheel)
  element?.removeEventListener("gesturestart", onGestureStart)
  element?.removeEventListener("gesturechange", onGestureChange as EventListener)
  element?.removeEventListener("gestureend", onGestureEnd)
})

defineExpose({
  scrollToCell: async (cell: XlsxCellAddress) => {
    await nextTick()
    gridRef.value?.scrollToCell(cell)
  },
  search: surfaceSearch.search,
  activateSearchMatch: surfaceSearch.activateSearchMatch,
  searchNext: surfaceSearch.searchNext,
  searchPrevious: surfaceSearch.searchPrevious,
  clearSearch: surfaceSearch.clearSearch,
  getSearchState: surfaceSearch.getSearchState,
  get scrollContainer() {
    return gridRef.value?.scrollContainer ?? null
  },
  getDocumentRevision: (): OfficeDocumentRevision => documentRevision.value,
  hitTest,
  async describeReference(reference: OfficeObjectReference, signal?: AbortSignal) {
    if (signal?.aborted) throw new DOMException("XLSX reference description aborted.", "AbortError")
    const descriptor = describeXlsxReference(referenceContext(), reference)
    if (!descriptor) throw new Error("XLSX reference is not resolvable in the current revision.")
    return descriptor
  },
  async resolveReference(reference: OfficeObjectReference): Promise<ResolveReferenceResult> {
    const result = resolveXlsxReference(referenceContext(), reference)
    emit("referenceResolve", { referenceId: reference.referenceId, result })
    return result
  },
  async scrollToReference(reference: OfficeObjectReference): Promise<void> {
    const result = resolveXlsxReference(referenceContext(), reference)
    if (result.status !== "exact" && result.status !== "relocated") return
    const locator = (result.reference as XlsxOfficeReference).locator
    if (locator.type === "manual-region") {
      await nextTick(); gridRef.value?.scrollToCell(locator.value.start)
    } else if (locator.value.kind === "range") {
      const match = /^\$?([A-Z]+)\$?([1-9][0-9]*)/u.exec(locator.value.a1)
      if (match) {
        let col = 0
        for (const character of match[1]!) col = col * 26 + character.charCodeAt(0) - 64
        await nextTick(); gridRef.value?.scrollToCell({ row: Number(match[2]) - 1, col: col - 1 })
      }
    } else if (locator.value.kind === "row") {
      await nextTick(); gridRef.value?.scrollToCell({ row: locator.value.start, col: 0 })
    } else if (locator.value.kind === "column") {
      await nextTick(); gridRef.value?.scrollToCell({ row: 0, col: locator.value.start })
    }
  },
  async captureReferencePreview(reference: OfficeObjectReference): Promise<Blob> {
    emitReferenceError("capture", "CAPTURE_UNSUPPORTED", "XLSX preview capture is not provided by the Surface; use the host's capture pipeline.", reference.referenceId)
    throw new Error("XLSX reference preview capture is unsupported")
  },
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

.xlsx-sheet-surface__region-layer {
  cursor: crosshair;
  inset: 0;
  position: absolute;
  touch-action: none;
  z-index: 30;
}

.xlsx-sheet-surface__region-frame {
  background: rgb(37 99 235 / 0.08);
  border: 2px solid #2563eb;
  box-sizing: border-box;
  pointer-events: none;
  position: absolute;
}
</style>
