<template>
  <div
    ref="surfaceRef"
    class="xlsx-sheet-surface"
    :class="{ 'xlsx-sheet-surface--print-preview': controller.isPrintPreviewOpen }"
    data-testid="xlsx-sheet-surface"
    @keydown="onKeydown"
    @contextmenu="onContextMenu"
    @pointermove="onReferencePointerMove"
    @pointerleave="onReferencePointerLeave"
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
        :render-mode="renderMode"
        @cell-double-click="emit('cellDoubleClick', $event)"
        @viewport-change="gridViewport = $event"
      >
        <template
          v-if="$slots['table-header-menu']"
          #table-header-menu="headerMenuProps"
        >
          <slot name="table-header-menu" v-bind="headerMenuProps" />
        </template>
      </XlsxGrid>
      <XlsxRangeHighlightLayer
        :controller="controller"
        :scroll-left="gridViewport.scrollLeft"
        :scroll-top="gridViewport.scrollTop"
        :viewport-width="gridViewport.width"
        :viewport-height="gridViewport.height"
      />
      <div v-if="objectOutline" class="xlsx-sheet-surface__object-outline" :style="objectOutline.style">
        <span class="xlsx-sheet-surface__object-outline-label">{{ objectOutline.label }}</span>
      </div>
      <XlsxChartOverlay
        :controller="controller"
        @chart-context-menu="onChartContextMenu"
        :is-dark="isDark ?? false"
        :scroll-left="gridViewport.scrollLeft"
        :scroll-top="gridViewport.scrollTop"
        :viewport-width="gridViewport.width"
        :viewport-height="gridViewport.height"
      />
      <XlsxImageLayer
        :controller="controller"
        :show-images="showImages ?? true"
        :scroll-left="gridViewport.scrollLeft"
        :scroll-top="gridViewport.scrollTop"
        :viewport-width="gridViewport.width"
        :viewport-height="gridViewport.height"
        @drawing-context-menu="onDrawingContextMenu"
      />
      <XlsxDrawingLayer
        :controller="controller"
        :scroll-left="gridViewport.scrollLeft"
        :scroll-top="gridViewport.scrollTop"
        :viewport-width="gridViewport.width"
        :viewport-height="gridViewport.height"
        @drawing-context-menu="onDrawingContextMenu"
      />
      <XlsxSelectionOverlay
        :controller="controller"
        :get-cell-style="getCellStyle"
        :selection-color="selectionColor"
        :selection-fill-color="selectionFillColor"
      />
      <XlsxContextMenu
        ref="contextMenuRef"
        :controller="controller"
        :target-element="gridElement"
        :suppress-context-menu="touchGestures.shouldSuppressContextMenu"
        @authoring="openAuthoring"
      />
    </template>
    <XlsxAuthoringDialog
      v-if="authoringKind"
      :controller="controller"
      :kind="authoringKind"
      :is-dark="isDark ?? false"
      @close="authoringKind = null"
    />
    <XlsxPrintPreview
      v-if="controller.isPrintPreviewOpen"
      :controller="controller"
      :is-dark="isDark ?? false"
      @close="controller.closePrintPreview()"
    />
    <div
      v-if="selectionMode === 'region'"
      class="xlsx-sheet-surface__region-layer"
      :aria-label="m.xlsx.surface.region"
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
import { createTouchGestureTracker, nextDoubleTapZoom } from "@arcships/office-runtime/touch-gestures"
import type {
  XlsxBehaviorRecord,
  XlsxCellAddress,
  XlsxCellStyleContext,
  XlsxOfficeReference,
  XlsxReferenceComment,
  XlsxReferenceContext,
  XlsxReferenceDrawing,
  XlsxReferenceSheet,
  XlsxReferenceTable,
  XlsxViewerController,
} from "@arcships/xlsx-core"
import {
  createXlsxBehaviorRecords,
  createXlsxCellReferenceDraft,
  createXlsxChartElementReferenceDraft,
  createXlsxChartReferenceDraft,
  createXlsxCommentReferenceDraft,
  createXlsxColumnReferenceDraft,
  createXlsxDrawingReferenceDraft,
  createXlsxRangeReferenceDraft,
  createXlsxRegionReferenceDraft,
  createXlsxRowReferenceDraft,
  createXlsxTableReferenceDraft,
  createXlsxWorksheetReferenceDraft,
  describeXlsxReference,
  parseA1RangeReference,
  resolveXlsxReference,
} from "@arcships/xlsx-core"
import {
  confirmOfficeReferenceDraft,
  createOfficeReferenceId,
  type OfficeDocumentRevision,
  type OfficeObjectReference,
  type OfficeReferenceCandidatePreview,
  type OfficeReferenceError,
  type OfficeRegionDraftChange,
  type OfficeSelectionMode,
  type SheetRegionPoint,
  type ResolveReferenceResult,
  type ManualRegionLocator,
} from "@arcships/office-interaction"
import XlsxGrid from "./XlsxGrid.vue"
import XlsxRangeHighlightLayer from "./XlsxRangeHighlightLayer.vue"
import XlsxChartOverlay from "./XlsxChartOverlay.vue"
import XlsxImageLayer from "./XlsxImageLayer.vue"
import XlsxDrawingLayer from "./XlsxDrawingLayer.vue"
import XlsxSelectionOverlay from "./XlsxSelectionOverlay.vue"
import XlsxContextMenu from "./XlsxContextMenu.vue"
import XlsxAuthoringDialog from "./XlsxAuthoringDialog.vue"
import XlsxPrintPreview from "./XlsxPrintPreview.vue"
import type { XlsxAuthoringDialogKind } from "./authoring-dialog"
import XlsxChartsheetSurface from "./XlsxChartsheetSurface.vue"
import { useXlsxSearch, type XlsxSearchState } from "../composables/useXlsxSearch"
import { provideXlsxMessages } from "../i18n"
import { formatOfficeMessage, localizeOfficeObjectDescriptor, localizeOfficeResolveResult } from "@arcships/office-interaction"
import type {
  OfficeRegionCaptureRequest,
  OfficeRegionCaptureSource,
} from "@arcships/vue-ui"

const props = withDefaults(
  defineProps<{
    controller: XlsxViewerController
    getCellStyle?: ((cell: XlsxCellAddress, context?: XlsxCellStyleContext) => Partial<CSSProperties> | undefined) | null
    isDark?: boolean | null
    readOnly?: boolean
    selectionColor?: string
    selectionFillColor?: string
    /** Grid body renderer: canvas (default) or the DOM fallback. */
    renderMode?: import("./grid-dom").XlsxGridRenderMode
    showImages?: boolean
    /** Controlled zoom factor. 1 = 100%. */
    zoom?: number
    enableGestureZoom?: boolean
    documentId?: string
    selectionMode?: OfficeSelectionMode
    emitReferenceCandidates?: boolean
    /**
     * Locale for the built-in UI copy and, unless `documentLocale` is set, for
     * number and date formats that name no locale of their own.
     */
    locale?: import("@arcships/office-interaction").OfficeLocale
    /** Overrides the number/date locale; defaults to `locale`. */
    documentLocale?: string
    messages?: import("@arcships/office-interaction").OfficeMessagesOverride
  }>(),
  {
    getCellStyle: null,
    isDark: false,
    readOnly: false,
    selectionColor: undefined,
    selectionFillColor: undefined,
    renderMode: "canvas",
    showImages: true,
    enableGestureZoom: true,
    selectionMode: "content",
    emitReferenceCandidates: false,
    locale: "en",
    documentLocale: undefined,
    messages: undefined,
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
  regionDraftChange: [event: OfficeRegionDraftChange]
  selectionCancel: [event: { mode: OfficeSelectionMode; reason: "escape" | "pointer-cancel" | "programmatic" }]
  referenceResolve: [event: { referenceId: string; result: ResolveReferenceResult }]
  referenceError: [error: OfficeReferenceError]
}>()

type GridInstance = InstanceType<typeof XlsxGrid>
type GridZoomAnchor = ReturnType<GridInstance["captureZoomAnchor"]>
type WebKitGestureEvent = Event & { clientX?: number; clientY?: number; scale?: number }

const surfaceRef = ref<HTMLElement | null>(null)
const m = provideXlsxMessages(() => props.locale, () => props.messages)

// Number/date formats without a `[$-…]` section follow the surface locale. A
// default `"en"` leaves a controller created with its own locale untouched.
let hasAppliedDocumentLocale = false
const effectiveDocumentLocale = computed(() => props.documentLocale ?? (props.locale === "en" ? undefined : props.locale))
watch(
  effectiveDocumentLocale,
  (locale) => {
    if (locale === undefined && !hasAppliedDocumentLocale) return
    hasAppliedDocumentLocale = locale !== undefined
    props.controller.setDocumentLocale?.(locale)
  },
  { immediate: true },
)
const descriptor = (
  key: keyof typeof m.value.common.descriptors,
  values: Record<string, string | number> = {},
) => formatOfficeMessage(m.value.common.descriptors[key], values)
const confirmFailed = () => formatOfficeMessage(m.value.common.reference.confirmFailed, { format: "XLSX" })
const gridRef = ref<GridInstance | null>(null)
const contextMenuRef = ref<{
  openAt(clientX: number, clientY: number, target?: unknown): boolean
  openForChart(chartId: string, clientX: number, clientY: number): boolean
  openForDrawing(kind: "image" | "shape", id: string, clientX: number, clientY: number): boolean
} | null>(null)
const gridViewport = ref({ height: 0, scrollLeft: 0, scrollTop: 0, width: 0 })
const authoringKind = ref<XlsxAuthoringDialogKind | null>(null)

function openAuthoring(kind: XlsxAuthoringDialogKind) {
  authoringKind.value = kind
}
const effectiveReadOnly = computed(() => props.readOnly || props.controller.isEditingLocked === true || props.controller.readOnly)
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

function a1ForCell(cell: XlsxCellAddress): string {
  return `${columnLabel(cell.col)}${cell.row + 1}`
}

function referenceTables(sheet: XlsxReferenceSheet | undefined): XlsxReferenceTable[] {
  if (!sheet) return []
  return (props.controller.tables ?? []).map((table) => ({
    name: table.name,
    ...(table.displayName && table.displayName !== table.name ? { displayName: table.displayName } : {}),
    sheet,
    a1: table.reference || `${a1ForCell(table.start)}:${a1ForCell(table.end)}`,
  }))
}

function sheetDrawings(sheet: XlsxReferenceSheet): XlsxReferenceDrawing[] {
  const sheetIndex = props.controller.activeSheetIndex
  const drawings: XlsxReferenceDrawing[] = []
  for (const image of props.controller.getSheetImages?.(sheetIndex) ?? []) {
    drawings.push({ id: image.id, objectType: "image", sheet, ...(image.name ? { name: image.name } : {}), ...(image.description ? { alt: image.description } : {}) })
  }
  for (const shape of props.controller.getSheetShapes?.(sheetIndex) ?? []) {
    drawings.push({ id: shape.id, objectType: "shape", sheet, ...(shape.name ? { name: shape.name } : {}) })
  }
  for (const control of props.controller.getSheetFormControls?.(sheetIndex) ?? []) {
    drawings.push({ id: control.id, objectType: "form-control", sheet, ...(control.name ? { name: control.name } : {}) })
  }
  return drawings
}

function sheetComments(sheet: XlsxReferenceSheet): XlsxReferenceComment[] {
  return (props.controller.activeSheet?.comments ?? []).map((comment) => ({
    sheet,
    cellA1: a1ForCell(comment.cell),
    ...(comment.text ? { text: comment.text } : {}),
    ...(comment.author ? { author: comment.author } : {}),
  }))
}

function sheetBehaviors(sheet: XlsxReferenceSheet): XlsxBehaviorRecord[] {
  const activeSheet = props.controller.activeSheet
  if (!activeSheet) return []
  const behaviors = createXlsxBehaviorRecords({
    sheet: { index: sheet.index, name: sheet.name, ...(sheet.id ? { id: sheet.id } : {}) },
    conditionalFormats: (activeSheet.conditionalFormatRules ?? []).flatMap((rule, index) => {
      const range = rule.ranges?.[0]
      if (!range) return []
      const start = a1ForCell(range.start)
      const end = a1ForCell(range.end)
      return [{
        appliesToA1: start === end ? start : `${start}:${end}`,
        index,
        ruleType: rule.kind,
      }]
    }),
    dataValidations: (activeSheet.dataValidations ?? []).flatMap((validation, index) => {
      const range = validation.ranges?.[0]
      if (!range) return []
      const start = a1ForCell(range.start)
      const end = a1ForCell(range.end)
      return [{
        appliesToA1: start === end ? start : `${start}:${end}`,
        index,
        ...(validation.validationType ? { validationType: validation.validationType } : {}),
        ...(validation.value1 ? { formula1: validation.value1 } : {}),
      }]
    }),
    definedNames: (activeSheet.definedNames ?? []).map((name) => ({
      name: name.name,
      refersTo: name.range,
      ...(name.scope === "sheet" ? { sheet: { index: sheet.index, name: sheet.name, ...(sheet.id ? { id: sheet.id } : {}) } } : {}),
    })),
    ...(activeSheet.protection?.enabled
      ? { protection: { enabled: true } }
      : {}),
  })
  return [...behaviors]
}

function referenceContext(): XlsxReferenceContext {
  const key = `${documentRevision.value.documentId}:${documentRevision.value.revision}:${props.controller.activeSheetIndex}`
  if (cachedReferenceContext?.key === key) return cachedReferenceContext.context
  const sheets = referenceSheets()
  const sheet = sheets[props.controller.activeSheetIndex]
  const context: XlsxReferenceContext = {
    revision: documentRevision.value,
    sheets,
    charts: props.controller.charts,
    tables: referenceTables(sheet),
    drawings: sheet ? sheetDrawings(sheet) : [],
    comments: sheet ? sheetComments(sheet) : [],
    behaviors: sheet ? sheetBehaviors(sheet) : [],
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
    emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : confirmFailed())
  }
}

function worksheetPath(sheet: XlsxReferenceSheet) {
  return [{ kind: "workbook" as const, label: descriptor("workbook") }, { kind: "worksheet" as const, label: sheet.name }]
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

const objectOutline = ref<{ style: Record<string, string>; label: string } | null>(null)
let candidateElements = new Map<string, Element>()
let candidateSheetRanges = new Map<string, { sheet: XlsxReferenceSheet; a1: string }>()

function clearObjectOutline(): void {
  objectOutline.value = null
  candidateElements = new Map()
  candidateSheetRanges = new Map()
}

function gridBodyRect(): DOMRect | undefined {
  const canvas = surfaceRef.value?.querySelector<HTMLCanvasElement>(".xlsx-grid__body")
  return canvas?.getBoundingClientRect?.()
}

function updateObjectOutline(candidates: readonly OfficeReferenceCandidatePreview[]): void {
  const active = candidates[0]
  if (!active) {
    objectOutline.value = null
    return
  }
  const root = surfaceRef.value?.getBoundingClientRect?.()
  if (!root) {
    objectOutline.value = null
    return
  }
  const element = candidateElements.get(active.candidateId)
  const elementRect = element && typeof element.getBoundingClientRect === "function"
    ? element.getBoundingClientRect()
    : undefined
  if (elementRect && elementRect.width > 0 && elementRect.height > 0) {
    objectOutline.value = {
      style: {
        left: `${elementRect.left - root.left}px`,
        top: `${elementRect.top - root.top}px`,
        width: `${elementRect.width}px`,
        height: `${elementRect.height}px`,
      },
      label: active.preview?.label ?? "",
    }
    return
  }
  const rangeInfo = candidateSheetRanges.get(active.candidateId)
  const body = gridBodyRect()
  const parsed = rangeInfo ? parseA1RangeReference(rangeInfo.a1) ?? undefined : undefined
  const viewport = rangeInfo && parsed
    ? gridRef.value?.sheetRegionViewportRect({
        space: "sheet",
        sheetId: rangeInfo.sheet.id ?? rangeInfo.sheet.name,
        start: { row: parsed.start.row, col: parsed.start.col, xOffset: 0, yOffset: 0 },
        end: { row: parsed.end.row, col: parsed.end.col, xOffset: 1, yOffset: 1 },
      })
    : undefined
  if (!viewport || !body) {
    objectOutline.value = null
    return
  }
  objectOutline.value = {
    style: {
      left: `${body.left - root.left + viewport.x * body.width}px`,
      top: `${body.top - root.top + viewport.y * body.height}px`,
      width: `${viewport.width * body.width}px`,
      height: `${viewport.height * body.height}px`,
    },
    label: active.preview?.label ?? "",
  }
}

function chartElementCandidateForPoint(point: { clientX: number; clientY: number }): OfficeReferenceCandidatePreview[] {
  const target = surfaceRef.value?.ownerDocument?.elementFromPoint?.(point.clientX, point.clientY)
  if (!target) return []
  const sheet = activeReferenceSheet()
  if (!sheet) return []
  const chartElement = target.closest<Element>("[data-xlsx-chart-series-index]")
  const chartHost = (chartElement ?? target).closest<HTMLElement>("[data-chart-id]")
  const chartId = chartHost?.dataset.chartId
  if (!chartId) return []
  const chart = props.controller.getChartById?.(chartId)
  if (!chart) return []
  if (chartElement) {
    const seriesIndex = Number(chartElement.getAttribute("data-xlsx-chart-series-index"))
    const pointAttribute = chartElement.getAttribute("data-xlsx-chart-point-index")
    const pointIndex = pointAttribute === null ? undefined : Number(pointAttribute)
    try {
      const target_ = Number.isSafeInteger(pointIndex) && pointIndex !== undefined
        ? { element: "point" as const, seriesIndex, pointIndex: pointIndex as number }
        : { element: "series" as const, seriesIndex }
      const draft = createXlsxChartElementReferenceDraft(referenceContext(), sheet, chartId, target_)
      const series = chart.series[seriesIndex]
      const label = target_.element === "series"
        ? `Series: ${series?.name?.trim() || `Series ${seriesIndex + 1}`}`
        : `Point ${(target_ as { pointIndex: number }).pointIndex + 1}`
      const candidateId = `xlsx:chart-element:${chartId}:${target_.element}:${seriesIndex}:${"pointIndex" in target_ ? target_.pointIndex : ""}`
      candidateElements.set(candidateId, chartElement)
      return [{
        candidateId,
        draft,
        preview: {
          label,
          path: [...worksheetPath(sheet), { kind: "chart", label: chart.name?.trim() || chart.title?.trim() || descriptor("chart") }, { kind: draft.kind, label }],
        },
        hit: "direct",
        depth: 4,
      }]
    } catch {
      return []
    }
  }
  try {
    const draft = createXlsxChartReferenceDraft(referenceContext(), sheet, chart)
    const label = chart.name?.trim() || chart.title?.trim() || descriptor("chart")
    const candidateId = `xlsx:chart:${chartId}`
    candidateElements.set(candidateId, chartHost!)
    return [{
      candidateId,
      draft,
      preview: { label, path: [...worksheetPath(sheet), { kind: "chart", label }] },
      hit: "direct",
      depth: 3,
    }]
  } catch {
    return []
  }
}

function drawingCandidateForPoint(point: { clientX: number; clientY: number }): OfficeReferenceCandidatePreview[] {
  const target = surfaceRef.value?.ownerDocument?.elementFromPoint?.(point.clientX, point.clientY)
  const drawingElement = target?.closest?.<HTMLElement>("[data-drawing-id]")
  const drawingId = drawingElement?.dataset.drawingId
  const objectType = drawingElement?.dataset.drawingKind
  const sheet = activeReferenceSheet()
  if (!drawingElement || !drawingId || !sheet) return []
  if (objectType !== "image" && objectType !== "shape" && objectType !== "form-control") return []
  const drawing = sheetDrawings(sheet).find((candidate) => candidate.id === drawingId && candidate.objectType === objectType)
  if (!drawing) return []
  try {
    const draft = createXlsxDrawingReferenceDraft(referenceContext(), drawing)
    const label = drawing.name?.trim() || drawing.alt?.trim() || objectType
    const candidateId = `xlsx:drawing:${objectType}:${drawingId}`
    candidateElements.set(candidateId, drawingElement)
    return [{
      candidateId,
      draft,
      preview: { label, path: [...worksheetPath(sheet), { kind: draft.kind, label }] },
      hit: "direct",
      depth: 3,
    }]
  } catch {
    return []
  }
}

function tableCandidateForCell(cell: XlsxCellAddress): OfficeReferenceCandidatePreview[] {
  const sheet = activeReferenceSheet()
  if (!sheet) return []
  const table = referenceTables(sheet).find((candidate) => {
    const range = parseA1RangeReference(candidate.a1)
    return range !== null && range !== undefined
      && cell.row >= range.start.row && cell.row <= range.end.row
      && cell.col >= range.start.col && cell.col <= range.end.col
  })
  if (!table) return []
  try {
    const draft = createXlsxTableReferenceDraft(referenceContext(), table)
    const label = table.displayName?.trim() || table.name
    const candidateId = `xlsx:table:${table.name}`
    candidateSheetRanges.set(candidateId, { sheet, a1: table.a1 })
    return [{
      candidateId,
      draft,
      preview: { label, path: [...worksheetPath(sheet), { kind: "table", label }] },
      hit: "direct",
      depth: 2,
    }]
  } catch {
    return []
  }
}

function commentCandidateForCell(cell: XlsxCellAddress): OfficeReferenceCandidatePreview[] {
  const sheet = activeReferenceSheet()
  if (!sheet) return []
  const a1 = a1ForCell(cell)
  const comment = sheetComments(sheet).find((candidate) => candidate.cellA1 === a1)
  if (!comment) return []
  try {
    const draft = createXlsxCommentReferenceDraft(referenceContext(), comment)
    const label = comment.text ? `Comment: ${comment.text.replace(/\s+/gu, " ").trim().slice(0, 80)}` : `Comment ${a1}`
    const candidateId = `xlsx:comment:${a1}`
    candidateSheetRanges.set(candidateId, { sheet, a1 })
    return [{
      candidateId,
      draft,
      preview: { label, path: [...worksheetPath(sheet), { kind: "cell", label: a1 }, { kind: "comment", label }] },
      hit: "direct",
      depth: 3,
    }]
  } catch {
    return []
  }
}

function candidateForPoint(point: { clientX: number; clientY: number }): OfficeReferenceCandidatePreview[] {
  const sheet = activeReferenceSheet()
  if (!sheet) return []
  candidateElements = new Map()
  candidateSheetRanges = new Map()
  const objectCandidates = [
    ...chartElementCandidateForPoint(point),
    ...drawingCandidateForPoint(point),
  ]
  const cell = gridRef.value?.hitTestCell(point.clientX, point.clientY)
  if (cell) objectCandidates.push(...tableCandidateForCell(cell), ...commentCandidateForCell(cell))
  if (objectCandidates.length > 0) return objectCandidates
  const axis = gridRef.value?.hitTestAxis(point.clientX, point.clientY)
  if (axis?.kind === "row") {
    const draft = createXlsxRowReferenceDraft(referenceContext(), sheet, axis.actualIndex)
    return [{ candidateId: `xlsx:row:${axis.actualIndex}`, draft, preview: { label: descriptor("row", { number: axis.actualIndex + 1 }), path: [...worksheetPath(sheet), { kind: "row", label: descriptor("row", { number: axis.actualIndex + 1 }) }] }, hit: "direct", depth: 2 }]
  }
  if (axis?.kind === "column") {
    const label = columnLabel(axis.actualIndex)
    const draft = createXlsxColumnReferenceDraft(referenceContext(), sheet, axis.actualIndex)
    return [{ candidateId: `xlsx:column:${axis.actualIndex}`, draft, preview: { label: descriptor("column", { column: label }), path: [...worksheetPath(sheet), { kind: "column", label: descriptor("column", { column: label }) }] }, hit: "direct", depth: 2 }]
  }
  if (axis?.kind === "corner") {
    const draft = createXlsxWorksheetReferenceDraft(referenceContext(), sheet)
    return [{ candidateId: `xlsx:worksheet:${sheet.index}`, draft, preview: { label: sheet.name, path: worksheetPath(sheet) }, hit: "direct", depth: 1 }]
  }
  if (!cell) return []
  const draft = createXlsxCellReferenceDraft(referenceContext(), sheet, cell)
  const label = `${sheet.name}!R${cell.row + 1}C${cell.col + 1}`
  return [{ candidateId: `xlsx:cell:${cell.row}:${cell.col}`, draft, preview: { label, path: [...worksheetPath(sheet), { kind: "cell", label }] }, hit: "direct", depth: 2 }]
}

function hitTest(point: { clientX: number; clientY: number }): readonly OfficeReferenceCandidatePreview[] {
  try { return candidateForPoint(point) } catch (reason) {
    emitReferenceError("hit-test", "HIT_TEST_FAILED", reason instanceof Error ? reason.message : formatOfficeMessage(m.value.common.reference.hitTestFailed, { format: "XLSX" }))
    return []
  }
}

function onReferencePointerMove(event: PointerEvent): void {
  if (selectionMode.value !== "object") return
  const candidates = hitTest(event)
  updateObjectOutline(candidates)
  if (!props.emitReferenceCandidates) return
  emit("referenceCandidateChange", { candidates, ...(candidates[0] ? { activeCandidateId: candidates[0].candidateId } : {}) })
}

function onReferenceClick(event: MouseEvent): void {
  if (selectionMode.value !== "object") return
  const candidate = hitTest(event)[0]
  if (candidate) confirmDraft(candidate.draft, event.detail === 0 ? "keyboard" : "pointer", candidate.preview, event.shiftKey)
}

function onReferencePointerLeave(): void {
  clearObjectOutline()
  if (selectionMode.value !== "object" || !props.emitReferenceCandidates) return
  emit("referenceCandidateChange", { candidates: [] })
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
  emit("regionDraftChange", {
    phase: "start",
    point: { space: "sheet", sheetId: regionDraft.value.sheetId, start },
  })
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

function onChartContextMenu(payload: { chartId: string; clientX: number; clientY: number }): void {
  contextMenuRef.value?.openForChart(payload.chartId, payload.clientX, payload.clientY);
}

function onDrawingContextMenu(payload: {
  id: string;
  kind: "image" | "shape";
  clientX: number;
  clientY: number;
}): void {
  contextMenuRef.value?.openForDrawing(payload.kind, payload.id, payload.clientX, payload.clientY);
}

function emitContextMenuAt(clientX: number, clientY: number): void {
  const sel = props.controller.selection
  const active = props.controller.activeCell
  const container = surfaceRef.value
  if (!container) return
  const rect = container.getBoundingClientRect()
  emit("contextMenu", {
    clientX,
    clientY,
    containerX: clientX - rect.left,
    containerY: clientY - rect.top,
    sheetName: props.controller.activeSheet?.name,
    selection: sel ? { start: { row: sel.start.row, col: sel.start.col }, end: { row: sel.end.row, col: sel.end.col } } : undefined,
    activeCell: active ? { row: active.row, col: active.col } : undefined,
  })
}

function onContextMenu(event: MouseEvent): void {
  if (touchGestures.shouldSuppressContextMenu(event.clientX, event.clientY)) {
    event.preventDefault()
    return
  }
  emitContextMenuAt(event.clientX, event.clientY)
}

const touchGestures = createTouchGestureTracker({
  isEnabled: () => selectionMode.value === "content",
  isDoubleTapEnabled: () => effectiveReadOnly.value && controlledZoom() !== undefined,
  getZoom: controlledZoom,
  onZoom: (zoom, center) => requestGestureZoom(zoom, center.clientX, center.clientY),
  onDoubleTap: (pointer) => {
    const current = controlledZoom()
    if (current === undefined) return
    requestGestureZoom(nextDoubleTapZoom(current), pointer.clientX, pointer.clientY)
  },
  onLongPress: (pointer) => {
    if (!contextMenuRef.value?.openAt(pointer.clientX, pointer.clientY, pointer.target)) return
    emitContextMenuAt(pointer.clientX, pointer.clientY)
  },
})

function onTouchPointerDown(event: PointerEvent): void { touchGestures.pointerDown(event) }
function onTouchPointerMove(event: PointerEvent): void { touchGestures.pointerMove(event) }
function onTouchPointerUp(event: PointerEvent): void { touchGestures.pointerUp(event) }
function onTouchPointerCancel(event: PointerEvent): void { touchGestures.pointerCancel(event) }

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
  element.addEventListener("pointerdown", onTouchPointerDown, { passive: true })
  element.addEventListener("pointermove", onTouchPointerMove, { passive: true })
  element.addEventListener("pointerup", onTouchPointerUp)
  element.addEventListener("pointercancel", onTouchPointerCancel)
})

onBeforeUnmount(() => {
  surfaceMounted.value = false
  gestureToken += 1
  touchGestures.reset()
  const element = surfaceRef.value
  element?.removeEventListener("wheel", onWheel)
  element?.removeEventListener("gesturestart", onGestureStart)
  element?.removeEventListener("gesturechange", onGestureChange as EventListener)
  element?.removeEventListener("gestureend", onGestureEnd)
  element?.removeEventListener("pointerdown", onTouchPointerDown)
  element?.removeEventListener("pointermove", onTouchPointerMove)
  element?.removeEventListener("pointerup", onTouchPointerUp)
  element?.removeEventListener("pointercancel", onTouchPointerCancel)
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
    return localizeOfficeObjectDescriptor(descriptor, m.value.common)
  },
  async resolveReference(reference: OfficeObjectReference): Promise<ResolveReferenceResult> {
    const result = resolveXlsxReference(referenceContext(), reference)
    const localized = localizeOfficeResolveResult(result, m.value.common)
    emit("referenceResolve", { referenceId: reference.referenceId, result: localized })
    return localized
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
    emitReferenceError("capture", "CAPTURE_UNSUPPORTED", formatOfficeMessage(m.value.common.reference.captureUnsupported, { format: "XLSX" }), reference.referenceId)
    throw new Error("XLSX reference preview capture is unsupported")
  },
  /** Region-capture hook for `captureRegion()` in `@arcships/vue-ui`. */
  getCaptureSurface(request: OfficeRegionCaptureRequest): OfficeRegionCaptureSource | null {
    const locator = request.region as { space?: string; sheetId?: string }
    const sheetId = request.target?.sheetId
      ?? (locator.space === "sheet" ? locator.sheetId : undefined)
      ?? props.controller.activeTab?.id
      ?? props.controller.activeSheet?.name
    const root = surfaceRef.value
    const body = root?.querySelector<HTMLCanvasElement>(".xlsx-grid__body")
    const resolved = locator.space === "sheet"
      ? gridRef.value?.sheetRegionViewportRect(request.region as Extract<ManualRegionLocator, { space: "sheet" }>) ?? undefined
      : undefined
    if (body) {
      return {
        element: body,
        raster: { source: body, width: body.width, height: body.height },
        ...(resolved ? { region: resolved } : {}),
        format: "xlsx",
        space: "sheet",
        ...(sheetId ? { sheetId } : {}),
        documentId: documentRevision.value.documentId,
        revision: documentRevision.value.revision,
      }
    }
    const grid = root?.querySelector<HTMLElement>(".xlsx-grid") ?? root
    if (!grid) return null
    return {
      element: grid,
      ...(resolved ? { region: resolved } : {}),
      format: "xlsx",
      space: "sheet",
      ...(sheetId ? { sheetId } : {}),
      documentId: documentRevision.value.documentId,
      revision: documentRevision.value.revision,
    }
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
  touch-action: pan-x pan-y;
  background: var(--xlsx-surface-bg, #ffffff);
}

.xlsx-sheet-surface__object-outline {
  border: 2px solid #2563eb;
  box-shadow: 0 0 0 1px rgb(255 255 255 / 0.75);
  box-sizing: border-box;
  pointer-events: none;
  position: absolute;
  z-index: 25;
}

.xlsx-sheet-surface__object-outline-label {
  background: #2563eb;
  border-radius: 2px;
  color: #fff;
  font-size: 11px;
  left: -2px;
  max-width: 280px;
  overflow: hidden;
  padding: 1px 4px;
  position: absolute;
  text-overflow: ellipsis;
  top: -20px;
  white-space: nowrap;
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

@media print {
  .xlsx-sheet-surface--print-preview > :not(.xlsx-print-preview) { display: none !important; }
  .xlsx-sheet-surface { overflow: visible !important; }
}
</style>
