<template>
  <div
    ref="scrollRef"
    class="pdf-surface"
    :class="{ 'pdf-surface--dark': props.isDark }"
    :style="minHeightStyle"
    data-testid="pdf-surface"
    tabindex="0"
    @scroll.passive="onScroll"
    @contextmenu.prevent="onContextMenu"
    @copy="onCopy"
    @keydown.esc="onSurfaceEscape"
  >
    <div v-if="!effectiveSource" class="pdf-surface__empty">{{ m.pdf.surface.noPdf }}</div>
    <div v-else-if="loading" class="pdf-surface__empty">{{ m.pdf.surface.opening }}</div>
    <div v-else-if="loadError" class="pdf-surface__error" :data-error-code="loadError.code">
      {{ loadError.code }}: {{ loadErrorMessage }}
    </div>
    <div v-else-if="renderDocument" class="pdf-surface__pages" :style="pageStackStyle">
      <div
        v-for="pageIndex in Array.from({ length: renderDocument.pageCount }, (_, i) => i)"
        :key="`page-${pageIndex}`"
        class="pdf-surface__page-slot"
        :data-page-index="pageIndex"
        :style="pageSlotStyle(pageIndex)"
      >
        <div v-if="pageSlots[pageIndex]?.state === 'loading'" class="pdf-surface__page-loading">
          {{ formatOfficeMessage(m.pdf.surface.renderingPage, { page: pageIndex + 1 }) }}
        </div>
        <div v-else-if="pageSlots[pageIndex]?.state === 'error'" class="pdf-surface__page-error">
          {{ formatOfficeMessage(m.pdf.surface.pageRenderFailed, { page: pageIndex + 1 }) }}
        </div>
        <img
          v-else-if="pageSlots[pageIndex]?.url"
          :src="pageSlots[pageIndex].url"
          :alt="formatOfficeMessage(m.pdf.surface.pageAlt, { page: pageIndex + 1 })"
          :style="pageImgStyle(pageIndex)"
          class="pdf-surface__page-img"
          draggable="false"
        />
        <div v-else class="pdf-surface__page-placeholder" />
        <div
          v-if="textLayerFor(pageIndex)"
          class="pdf-surface__text-layer"
          :data-text-page="pageIndex"
          :data-text-transform="textLayerTransform(pageIndex)"
          :style="textLayerStyle(pageIndex)"
        >
          <span
            v-for="glyph in textLayerFor(pageIndex)?.glyphs ?? []"
            :key="`glyph-${glyph.charIndex}`"
            class="pdf-surface__text-glyph"
            :class="{ 'pdf-surface__text-glyph--unmapped': glyph.unmapped }"
            :data-char-index="glyph.charIndex"
            :aria-hidden="glyph.unmapped ? 'true' : undefined"
            :style="textGlyphStyle(glyph)"
          >{{ glyph.text }}</span>
        </div>
        <div class="pdf-surface__overlay" aria-hidden="true">
          <div
            v-for="entry in searchRectsForPage(pageIndex)"
            :key="entry.key"
            class="pdf-surface__search-hit"
            :class="{ 'pdf-surface__search-hit--active': entry.active }"
            :data-search-result-index="entry.resultIndex"
            :style="displayRectStyle(pageIndex, entry.rect)"
          />
          <div
            v-for="entry in selectionRectsForPage(pageIndex)"
            :key="entry.key"
            class="pdf-surface__selection-hit"
            :style="displayRectStyle(pageIndex, entry.rect)"
          />
          <div
            v-for="annotation in annotationHitsForPage(pageIndex)"
            :key="`annotation-${annotation.id}`"
            class="pdf-surface__annotation-hit"
            :class="`pdf-surface__annotation-hit--${annotation.kind}`"
            :data-annotation-kind="annotation.kind"
            :data-annotation-page="pageIndex"
            :style="displayRectStyle(pageIndex, annotation.rect)"
          />
        </div>
        <div
          class="pdf-surface__interaction"
          :data-selection-mode="selectionMode"
          :aria-label="selectionMode === 'region' ? m.pdf.surface.selectionRegion : m.pdf.surface.selectionText"
          :tabindex="selectionMode === 'region' ? 0 : undefined"
          @pointerdown="onSurfacePointerDown($event, pageIndex)"
          @pointermove="onSurfacePointerMove($event, pageIndex)"
          @pointerup="onSurfacePointerUp($event, pageIndex)"
          @pointercancel="onSurfacePointerCancel($event)"
          @click="onSurfaceClick($event, pageIndex)"
          @keydown.esc="onRegionKeyboardCancel"
        >
          <div
            v-if="selectionMode === 'region' && regionPageIndex === pageIndex && regionFrame"
            class="pdf-surface__region-frame"
            :style="regionFrame"
          />
        </div>
        <div
          v-if="selectionMode === 'content' && formWritesAvailable && formControlsForPage(pageIndex).length > 0"
          class="pdf-surface__form-layer"
          :data-form-page="pageIndex"
        >
          <template v-for="annotation in formControlsForPage(pageIndex)" :key="`form-${annotation.id}`">
            <select
              v-if="annotation.fieldType === 'combobox' || annotation.fieldType === 'listbox'"
              class="pdf-surface__form-control pdf-surface__form-control--select"
              :data-form-field-id="annotation.id"
              :data-form-field-type="annotation.fieldType"
              :aria-label="annotation.fieldName || annotation.id"
              :disabled="annotation.readOnly"
              :value="annotation.fieldValue ?? ''"
              :style="displayRectStyle(pageIndex, annotation.rect)"
              @change="onFormSelectChange($event, pageIndex, annotation)"
              @keydown.stop
            >
              <option
                v-for="option in annotation.options ?? []"
                :key="option.label"
                :value="option.label"
                :selected="option.selected || option.label === annotation.fieldValue"
              >{{ option.label }}</option>
            </select>
            <input
              v-else
              class="pdf-surface__form-control"
              :class="`pdf-surface__form-control--${annotation.fieldType}`"
              :data-form-field-id="annotation.id"
              :data-form-field-type="annotation.fieldType"
              :aria-label="annotation.fieldName || annotation.id"
              :type="annotation.fieldType === 'text' ? 'text' : annotation.fieldType === 'radio' ? 'radio' : 'checkbox'"
              :name="annotation.fieldName || annotation.id"
              :checked="annotation.fieldType === 'text' ? undefined : isFormFieldChecked(annotation)"
              :value="annotation.fieldType === 'text' ? annotation.fieldValue ?? '' : annotation.exportValue ?? annotation.fieldValue ?? 'on'"
              :disabled="annotation.readOnly"
              :style="displayRectStyle(pageIndex, annotation.rect)"
              @change="annotation.fieldType === 'text' ? onFormTextChange($event, pageIndex, annotation) : onFormToggleChange($event, pageIndex, annotation)"
              @keydown.stop
            />
          </template>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref, watch } from "vue"
import type { PageTextSlice, PdfPageGeometry, Position } from "@embedpdf/models"
import { expandToWordBoundary, glyphAt, rectsWithinSlice } from "@embedpdf/plugin-selection"
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom"
import { createTouchGestureTracker, nextDoubleTapZoom } from "@arcships/office-runtime/touch-gestures"
import {
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
  createSurfaceSearchSession,
  resolveOfficeUrl,
  type OfficeSource,
} from "@arcships/office-runtime"
import {
  DEFAULT_PDF_MAX_FILE_SIZE,
  loadVerifiedPdfSource,
  localizePdfErrorMessage,
  toPdfLoadError,
  type PdfDiagnostic,
  type PdfLoadError,
  type PdfSource,
  type PdfUrlPolicy,
} from "../pdf-url-policy"
import {
  createPdfRenderRuntime,
  type PdfAnnotationAppearancePatch,
  type PdfFormFieldType,
  type PdfFormFieldValue,
  type PdfPageTextLayer,
  type PdfPageTextLayerGlyph,
  type PdfPasswordHandler,
  type PdfRenderAnnotation,
  type PdfRenderDocument,
  type PdfRenderRect,
  type PdfRenderRuntime,
  type PdfRotation,
  type PdfSearchHit,
  type PdfSearchState,
  type PdfSelectionSegment,
  type PdfSelectionState,
} from "../pdf/pdf-render-runtime"
import {
  pdfCanonicalRectToDisplay,
  pdfCanonicalToDisplayTransform,
  pdfDisplayPointToCanonical,
  pdfDisplaySize,
} from "../pdf/pdf-coordinate-transform"
import { currentRenderDpr, watchDevicePixelRatio } from "../pdf/pdf-device-pixel-ratio"
import {
  DEFAULT_PDF_BITMAP_BUDGET_BYTES,
  createPdfBitmapCache,
  type PdfBitmapCache,
} from "../pdf/pdf-bitmap-cache"
import {
  DEFAULT_MAX_PAGE_RASTER_PIXELS,
  planPageRaster,
  type PdfPageRasterPlan,
} from "../pdf/pdf-page-raster"
import {
  DEFAULT_PAGE_OVERSCAN,
  computePageOffsets,
  pageIndexAtPosition,
  pageRangeEquals,
  planPageRenders,
  visiblePageRange,
  type PdfPageBitmapState,
  type PdfPageRange,
  type PdfPageSlotSnapshot,
} from "../pdf/pdf-page-viewport"
import {
  hasUnmappedPdfText,
  joinSelectionTexts,
  pdfPageCharCount,
  selectionPageIndexes,
  selectionPageRanges,
  type PdfSelectionPoint,
} from "../pdf/pdf-selection-range"
import {
  nearestGlyphWithinContainingRun,
  wordRangeFromCharacterTexts,
} from "../pdf/pdf-selection-interaction"
import {
  createPdfPageReferenceDraft,
  createPdfRegionReferenceDraft,
  createPdfTextReferenceDraft,
  describePdfReference,
  resolvePdfReference,
  type PdfOfficeReference,
  type PdfReferenceContext,
} from "../pdf/pdf-reference-adapter"
import {
  confirmOfficeReferenceDraft,
  createOfficeReferenceId,
  type NormalizedRect,
  type OfficeDocumentRevision,
  type OfficeObjectReference,
  type OfficeReferenceCandidatePreview,
  type OfficeReferenceError,
  type OfficeRegionDraftChange,
  type OfficeSelectionMode,
  type ResolveReferenceResult,
} from "@arcships/office-interaction"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { providePdfMessages } from "../i18n"

// ── Types ────────────────────────────────────────────────────────────
interface PageSlot {
  url: string | null
  state: PdfPageBitmapState
  /** Zoom/rotation/DPR key of the bitmap in `url`; an empty key is stale. */
  renderKey: string
  /** Decoded RGBA footprint of the bitmap in `url`, for the bitmap budget. */
  bitmapBytes: number
}

interface PageSelectionRects {
  pageIndex: number
  rects: readonly PdfRenderRect[]
}

// ── Props ────────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    src?: string
    source?: PdfSource
    urlPolicy?: PdfUrlPolicy
    runtime?: PdfRenderRuntime
    pdfiumWasmUrl?: string
    maxFileSize?: number
    password?: string
    onPassword?: PdfPasswordHandler
    className?: string
    /** Controlled zoom factor. 1 = 100%. */
    zoom?: number
    defaultZoom?: number
    enableGestureZoom?: boolean
    /** When true, zoom auto-fits the widest page to container width. */
    fitWidth?: boolean
    documentId?: string
    selectionMode?: OfficeSelectionMode
    emitReferenceCandidates?: boolean
    /** Optional minimum height; the surface itself does not impose one. */
    minHeight?: string | number
    /** Dark viewer chrome and page surround. Defaults to false. */
    isDark?: boolean
    locale?: import("@arcships/office-interaction").OfficeLocale
    messages?: import("@arcships/office-interaction").OfficeMessagesOverride
    /**
     * Byte budget for page bitmaps kept outside the viewport. Pages scroll back
     * from this cache instead of rasterizing again; the least recently used
     * bitmap is released once the budget is exceeded. Defaults to 64 MiB.
     */
    maxBitmapBytes?: number
    /**
     * Per-page raster pixel cap (device px). Above it the page renders at a
     * reduced scale and is downscaled for display. Defaults to 8 MP.
     */
    maxPagePixels?: number
  }>(),
  {
    maxFileSize: DEFAULT_PDF_MAX_FILE_SIZE,
    defaultZoom: 1,
    enableGestureZoom: true,
    selectionMode: "content",
    emitReferenceCandidates: false,
    minHeight: undefined,
    isDark: false,
    locale: "en",
    messages: undefined,
    maxBitmapBytes: DEFAULT_PDF_BITMAP_BUDGET_BYTES,
    maxPagePixels: DEFAULT_MAX_PAGE_RASTER_PIXELS,
  },
)

const emit = defineEmits<{
  "document-load-success": [numPages: number]
  "document-load-error": [error: PdfLoadError]
  "visible-page-change": [pageIndex: number]
  contextMenu: [ctx: { pageIndex: number; clientX: number; clientY: number; containerX: number; containerY: number }]
  selectionChange: [selection: PdfSelectionState]
  searchStateChange: [state: PdfSearchState]
  diagnostic: [event: PdfDiagnostic]
  "update:zoom": [zoom: number]
  documentRevisionChange: [revision: OfficeDocumentRevision]
  referenceCandidateChange: [change: { candidates: readonly OfficeReferenceCandidatePreview[]; activeCandidateId?: string }]
  referenceConfirm: [event: ReturnType<typeof confirmOfficeReferenceDraft>]
  regionDraftChange: [event: OfficeRegionDraftChange]
  selectionCancel: [event: { mode: OfficeSelectionMode; reason: "escape" | "pointer-cancel" | "programmatic" }]
  linkActivate: [event: {
    pageIndex: number
    annotationId: string
    clientX: number
    clientY: number
    url?: string
    targetPageIndex?: number
    blocked?: boolean
  }]
  annotationActivate: [event: {
    pageIndex: number
    annotationId: string
    kind: "widget" | "annotation"
    clientX: number
    clientY: number
    fieldName?: string
    fieldValue?: string
    contents?: string
  }]
  formFieldChange: [event: {
    pageIndex: number
    annotationId: string
    annotation: PdfRenderAnnotation
    fieldName?: string
    fieldType?: PdfFormFieldType
    value?: string
  }]
  formFieldError: [event: { pageIndex: number; annotationId: string; message: string }]
  annotationChange: [event: {
    pageIndex: number
    annotationId: string
    annotation: PdfRenderAnnotation
    patch: PdfAnnotationAppearancePatch
  }]
  annotationError: [event: { pageIndex: number; annotationId: string; message: string }]
  referenceResolve: [event: { referenceId: string; result: ResolveReferenceResult }]
  referenceError: [error: OfficeReferenceError]
}>()

// ── Runtime & loading ────────────────────────────────────────────────
const scrollRef = ref<HTMLElement | null>(null)
const m = providePdfMessages(() => props.locale, () => props.messages)
const descriptor = (
  key: keyof typeof m.value.common.descriptors,
  values: Record<string, string | number> = {},
) => formatOfficeMessage(m.value.common.descriptors[key], values)
const confirmFailed = () => formatOfficeMessage(m.value.common.reference.confirmFailed, { format: "PDF" })
const internalZoom = ref(clampSurfaceZoom(props.defaultZoom))
const zoom = computed({
  get: () => props.zoom === undefined ? internalZoom.value : clampSurfaceZoom(props.zoom),
  set: (value: number) => {
    const next = clampSurfaceZoom(value)
    if (props.zoom === undefined) internalZoom.value = next
    else emit("update:zoom", next)
  },
})
const rotation = ref<PdfRotation>(0)
const minHeightStyle = computed(() => {
  const value = props.minHeight
  if (value === undefined) return undefined
  return { minHeight: typeof value === "number" ? `${value}px` : value }
})
/** Pixel ratio the current bitmaps were rendered at; re-renders on display change. */
const renderDpr = ref(currentRenderDpr())
const loading = ref(false)
const loadError = ref<PdfLoadError | null>(null)
const loadErrorMessage = computed(() => loadError.value
  ? localizePdfErrorMessage(loadError.value, m.value.pdf.errors)
  : "")
const renderDocument = ref<PdfRenderDocument | null>(null)
const printMode = ref(false)
const verifiedBlob = ref<Blob | null>(null)
const selectionMode = computed(() => props.selectionMode)
const revisionCounter = ref(0)
const generatedDocumentId = `pdf-surface-${typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2)}`
const documentRevision = computed<OfficeDocumentRevision & { format: "pdf" }>(() => ({
  format: "pdf",
  documentId: props.documentId?.trim() || generatedDocumentId,
  revision: `${generatedDocumentId}:${revisionCounter.value}`,
}))

const effectiveSource = computed<PdfSource | undefined>(() => {
  if (props.source) return props.source
  if (props.src) return { kind: "url", url: props.src }
  return undefined
})

let ownedRuntime: PdfRenderRuntime | undefined
let ownedRuntimeWasmUrl: string | undefined
let runtimeForDocument: PdfRenderRuntime | undefined
let latestRequestId = 0
let renderGeneration = 0
let renderAbort: AbortController | undefined
let printMediaQuery: MediaQueryList | undefined
let printPageStyle: HTMLStyleElement | undefined
let stopPrintDetection: (() => void) | undefined
let unmounted = false
let stopDprWatch: (() => void) | undefined

const searchState = ref<PdfSearchState>({
  status: "idle",
  query: "",
  matches: [],
  activeIndex: -1,
})
const selectionState = ref<PdfSelectionState>({ kind: "none" })
const selectionVisual = ref<{ ranges: readonly PageSelectionRects[] } | null>(null)
const pageAnnotations = reactive<Record<number, readonly PdfRenderAnnotation[]>>({})
const textLayers = reactive<Record<number, PdfPageTextLayer>>({})
const formWritesAvailable = ref(false)
const annotationRequests = new Map<number, Promise<readonly PdfRenderAnnotation[]>>()
const textLayerRequests = new Map<number, Promise<PdfPageTextLayer | undefined>>()
const geometryCache = new Map<number, PdfPageGeometry>()
const geometryRequests = new Map<number, Promise<PdfPageGeometry>>()
const characterTextCache = new Map<number, Map<number, string>>()
let selectionRequest = 0
const PDF_SELECTION_DRAG_THRESHOLD_PX = 4
type ClientPoint = { x: number; y: number }
type PdfWordSegmenterConstructor = new (
  locales?: string | readonly string[],
  options?: { granularity: "word" },
) => {
  segment(input: string): Iterable<{ index: number; segment: string; isWordLike?: boolean }>
}
let selectionGesture: {
  request: number
  pageIndex: number
  pointerId: number
  layer: HTMLElement
  down: ClientPoint
  latest: ClientPoint
  pointerActive: boolean
  thresholdExceeded: boolean
  started: boolean
  geometries: Map<number, PdfPageGeometry>
  pendingGeometries: Set<number>
  anchor?: PdfSelectionPoint
  focus?: PdfSelectionPoint
  visualKey?: string
} | null = null
let selectionFrame: number | undefined
let selectionClickCandidate = false

function referenceContext(): PdfReferenceContext {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) throw new Error("PDF reference context is unavailable before the document loads.")
  return { revision: documentRevision.value, document, runtime }
}

function ensurePageContentHash(pageIndex: number): Promise<void> | undefined {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime || pageIndex < 0 || pageIndex >= document.pageCount) return undefined
  const page = document.pages[pageIndex]
  if (!page || page.contentHash) return undefined
  if (typeof runtime.getPageContentHash !== "function") return undefined
  return runtime.getPageContentHash(document, pageIndex)
    .then(() => undefined)
    .catch(() => {
      // A content digest is optional metadata; fall back to the geometry-only
      // page signature when text extraction is unavailable.
      return undefined
    })
}

function emitReferenceError(operation: OfficeReferenceError["operation"], code: OfficeReferenceError["code"], message: string, referenceId?: string): void {
  emit("referenceError", { code, operation, format: "pdf", recoverable: true, ...(referenceId ? { referenceId } : {}), message })
}

function confirmDraft(
  draft: Parameters<typeof confirmOfficeReferenceDraft>[0],
  trigger: "pointer" | "keyboard" | "touch" | "programmatic",
  snapshot?: Parameters<typeof confirmOfficeReferenceDraft>[1]["snapshot"],
  additiveRequested = false,
): void {
  try {
    emit("referenceConfirm", confirmOfficeReferenceDraft(draft, { referenceId: createOfficeReferenceId(), trigger, additiveRequested, ...(snapshot ? { snapshot } : {}) }))
  } catch (reason) {
    emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : confirmFailed())
  }
}

function pageCandidate(pageIndex: number): OfficeReferenceCandidatePreview {
  const draft = createPdfPageReferenceDraft(referenceContext(), pageIndex)
  return {
    candidateId: `pdf:page:${pageIndex}`,
    draft,
    preview: { label: descriptor("page", { number: pageIndex + 1 }), path: [{ kind: "document", label: descriptor("pdf") }, { kind: "page", label: descriptor("page", { number: pageIndex + 1 }) }] },
    hit: "inside",
    depth: 1,
  }
}

function hitTest(point: { clientX: number; clientY: number }): readonly OfficeReferenceCandidatePreview[] {
  try {
    const element = scrollRef.value?.ownerDocument?.elementFromPoint?.(point.clientX, point.clientY)
    const page = element?.closest<HTMLElement>("[data-page-index]")
    const pageIndex = Number(page?.dataset.pageIndex)
    return Number.isSafeInteger(pageIndex) && pageIndex >= 0 ? [pageCandidate(pageIndex)] : []
  } catch (reason) {
    emitReferenceError("hit-test", "HIT_TEST_FAILED", reason instanceof Error ? reason.message : formatOfficeMessage(m.value.common.reference.hitTestFailed, { format: "PDF" }))
    return []
  }
}

type RegionGesture = { pointerId: number; pageIndex: number; layer: HTMLElement; startX: number; startY: number }
const regionGesture = ref<RegionGesture | null>(null)
const regionRect = ref<NormalizedRect | null>(null)
const regionPageIndex = ref<number | null>(null)
const regionFrame = computed(() => regionRect.value ? {
  left: `${regionRect.value.x * 100}%`, top: `${regionRect.value.y * 100}%`,
  width: `${regionRect.value.width * 100}%`, height: `${regionRect.value.height * 100}%`,
} : undefined)

function clampUnit(value: number): number { return Math.max(0, Math.min(1, value)) }
function pageRegionRect(gesture: RegionGesture, clientX: number, clientY: number): NormalizedRect {
  const bounds = gesture.layer.getBoundingClientRect()
  const x1 = clampUnit((Math.min(gesture.startX, clientX) - bounds.left) / bounds.width)
  const y1 = clampUnit((Math.min(gesture.startY, clientY) - bounds.top) / bounds.height)
  const x2 = clampUnit((Math.max(gesture.startX, clientX) - bounds.left) / bounds.width)
  const y2 = clampUnit((Math.max(gesture.startY, clientY) - bounds.top) / bounds.height)
  return { x: x1, y: y1, width: x2 - x1, height: y2 - y1 }
}

function beginRegion(event: PointerEvent, pageIndex: number): void {
  if (event.button !== 0) return
  const layer = event.currentTarget as HTMLElement
  const bounds = layer.getBoundingClientRect()
  const x = bounds.width > 0 ? clampUnit((event.clientX - bounds.left) / bounds.width) : 0
  const y = bounds.height > 0 ? clampUnit((event.clientY - bounds.top) / bounds.height) : 0
  regionGesture.value = { pointerId: event.pointerId, pageIndex, layer, startX: event.clientX, startY: event.clientY }
  regionPageIndex.value = pageIndex
  regionRect.value = { x, y, width: 0, height: 0 }
  layer.setPointerCapture?.(event.pointerId)
  emit("regionDraftChange", { phase: "start", point: { space: "page", pageIndex, x, y } })
  event.preventDefault()
}

function updateRegion(event: PointerEvent): void {
  const gesture = regionGesture.value
  if (!gesture || gesture.pointerId !== event.pointerId) return
  regionRect.value = pageRegionRect(gesture, event.clientX, event.clientY)
  emit("regionDraftChange", { phase: "change", region: { space: "page", pageIndex: gesture.pageIndex, rect: regionRect.value } })
}

function finishRegion(event: PointerEvent, cancelled: boolean): void {
  const gesture = regionGesture.value
  if (!gesture || gesture.pointerId !== event.pointerId) return
  if (gesture.layer.hasPointerCapture?.(event.pointerId)) gesture.layer.releasePointerCapture?.(event.pointerId)
  const rect = regionRect.value
  regionGesture.value = null
  regionRect.value = null
  regionPageIndex.value = null
  if (cancelled || !rect || rect.width < 0.005 || rect.height < 0.005) {
    emit("selectionCancel", { mode: "region", reason: cancelled ? "pointer-cancel" : "programmatic" })
    return
  }
  const context = referenceContext()
  const pending = ensurePageContentHash(gesture.pageIndex)
  const trigger = event.pointerType === "touch" ? "touch" : "pointer"
  const confirmRegion = () => {
    if (renderDocument.value?.id !== context.document.id) return
    confirmDraft(createPdfRegionReferenceDraft(context, { space: "page", pageIndex: gesture.pageIndex, rect }), trigger)
  }
  if (pending) void pending.then(confirmRegion)
  else confirmRegion()
}

function onSurfacePointerDown(event: PointerEvent, pageIndex: number): void {
  if (selectionMode.value === "content") onSelectionPointerDown(event, pageIndex)
  else if (selectionMode.value === "region") beginRegion(event, pageIndex)
}
function onSurfacePointerMove(event: PointerEvent, pageIndex: number): void {
  if (selectionMode.value === "content") onSelectionPointerMove(event, pageIndex)
  else if (selectionMode.value === "region") updateRegion(event)
  else if (props.emitReferenceCandidates) {
    const candidates = [pageCandidate(pageIndex)]
    emit("referenceCandidateChange", { candidates, activeCandidateId: candidates[0].candidateId })
  }
}
function onSurfacePointerUp(event: PointerEvent, pageIndex: number): void {
  if (selectionMode.value === "content") onSelectionPointerUp(event, pageIndex)
  else if (selectionMode.value === "region") finishRegion(event, false)
}
function onSurfacePointerCancel(event: PointerEvent): void {
  if (selectionMode.value === "content") onSelectionPointerCancel(event)
  else if (selectionMode.value === "region") finishRegion(event, true)
}
function onSurfaceClick(event: MouseEvent, pageIndex: number): void {
  if (selectionMode.value === "content") onSelectionClick(event, pageIndex)
  else if (selectionMode.value === "object") {
    const context = referenceContext()
    const trigger = event.detail === 0 ? "keyboard" : "pointer"
    const additive = event.shiftKey
    const confirmPage = () => {
      if (renderDocument.value?.id !== context.document.id) return
      const candidate = pageCandidate(pageIndex)
      confirmDraft(candidate.draft, trigger, candidate.preview, additive)
    }
    const pending = ensurePageContentHash(pageIndex)
    if (pending) void pending.then(confirmPage)
    else confirmPage()
  }
}
function onRegionKeyboardCancel(event?: KeyboardEvent): void {
  if (selectionMode.value !== "region") return
  // The Surface also listens for Escape; stop the bubble so the host receives
  // exactly one `selectionCancel` per key press.
  event?.stopPropagation?.()
  event?.preventDefault?.()
  regionGesture.value = null; regionRect.value = null; regionPageIndex.value = null
  emit("selectionCancel", { mode: "region", reason: "escape" })
}

const surfaceSearch = createSurfaceSearchSession<PdfSearchHit>({
  async search(query, { signal }) {
    const doc = renderDocument.value
    const rt = runtimeForDocument
    if (!doc || !rt) return []
    return rt.search(doc, query, { signal })
  },
  activate(match, index, { signal }) {
    return activatePdfSearchHit(match, index, signal)
  },
})
surfaceSearch.subscribe((next) => {
  searchState.value = next
  emit("searchStateChange", next)
})

const pdfRuntimeId = `pdf-surface-${
  typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}`
const pdfTaskSequence = createOfficeTaskSequence(pdfRuntimeId)
const pdfLoadCoordinator = createLatestTaskCoordinator(pdfTaskSequence)

// ── Page slots & viewport-driven rasterization ───────────────────────
// Only the pages near the viewport keep a rasterized bitmap. Everything else
// stays a correctly sized placeholder until scrolling reaches it, so zoom,
// rotation and resize re-rasterize the visible window instead of every page.
const PDF_PAGE_GAP_PX = 8
const PDF_PAGES_PADDING_TOP_PX = 16
const PDF_FALLBACK_PAGE_HEIGHT_PX = 200
/** Pages kept rasterized above and below the viewport. */
const PDF_PAGE_OVERSCAN = DEFAULT_PAGE_OVERSCAN
/** Trailing debounce so a fast scroll rasterizes only where it stops. */
const PDF_SCROLL_RENDER_DEBOUNCE_MS = 100
/** Debounce so a resize drag re-measures and re-renders once, not every frame. */
const PDF_RESIZE_RENDER_DEBOUNCE_MS = 120
const pageStackStyle = {
  gap: `${PDF_PAGE_GAP_PX}px`,
  paddingTop: `${PDF_PAGES_PADDING_TOP_PX}px`,
}

const pageSlots = reactive<Record<number, PageSlot>>({})
const viewportHeightPx = ref(0)
/**
 * Off-viewport bitmaps under the `maxBitmapBytes` budget. Pages scroll back
 * from here instead of rasterizing again; `bitmapCache` owns no URLs, the slot
 * keeps the object URL and the cache only records its byte footprint.
 */
let bitmapCache: PdfBitmapCache = createPdfBitmapCache({ maxBytes: props.maxBitmapBytes })
let pageLayoutCache: { key: string; offsets: number[]; heights: number[] } | undefined
let renderedRange: PdfPageRange = { first: 0, last: -1 }
let visiblePageIndex = -1
let renderTimer: ReturnType<typeof setTimeout> | undefined
let resizeTimer: ReturnType<typeof setTimeout> | undefined
let scrollFrame: number | undefined
let resizeObserver: ResizeObserver | undefined

function initPageSlots(): void {
  for (const key of Object.keys(pageSlots)) delete pageSlots[Number(key)]
  bitmapCache.clear()
  if (!renderDocument.value) return
  for (let i = 0; i < renderDocument.value.pageCount; i++) {
    pageSlots[i] = { url: null, state: "pending", renderKey: "", bitmapBytes: 0 }
  }
  renderedRange = { first: 0, last: -1 }
  visiblePageIndex = -1
}

/**
 * Zoom, rotation and display scale moved: every rasterized bitmap is stale, but
 * the window keeps showing it until the replacement arrives, so the page stack
 * never collapses to empty placeholders mid-gesture. Cached off-viewport
 * bitmaps cannot be reused at the new scale, so they are released right away.
 */
function markPageBitmapsStale(): void {
  for (const slot of Object.values(pageSlots)) {
    if (slot.url) slot.renderKey = ""
  }
  releasePageBitmaps(bitmapCache.clear())
  renderedRange = { first: 0, last: -1 }
  visiblePageIndex = -1
}

function invalidatePageLayout(): void {
  pageLayoutCache = undefined
}

function pageLayoutKey(): string {
  const doc = renderDocument.value
  if (!doc) return "none"
  return `${doc.id}:${doc.pageCount}:${zoom.value}:${rotation.value}`
}

/** Bitmaps are only valid for the zoom/rotation/DPR they were rasterized at. */
function pageRenderKey(): string {
  return `${zoom.value}:${rotation.value}:${renderDpr.value}`
}

/**
 * Cached page offsets for the current document / zoom / rotation.
 *
 * Slot widths and heights are inline styles and the stack gap/padding are part
 * of this component's own layout, so offsets are derived arithmetically once
 * per layout change instead of reading one DOM node per page on every scroll.
 */
function pageLayout(): { offsets: number[]; heights: number[] } {
  const key = pageLayoutKey()
  if (pageLayoutCache?.key === key) return pageLayoutCache
  const pageCount = renderDocument.value?.pageCount ?? 0
  const heights: number[] = []
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
    heights.push(pageDisplaySize(pageIndex)?.height ?? PDF_FALLBACK_PAGE_HEIGHT_PX)
  }
  pageLayoutCache = {
    key,
    offsets: computePageOffsets(heights, {
      gap: PDF_PAGE_GAP_PX,
      paddingTop: PDF_PAGES_PADDING_TOP_PX,
    }),
    heights,
  }
  return pageLayoutCache
}

function measureViewportHeight(): void {
  const height = scrollRef.value?.clientHeight
  if (typeof height === "number" && Number.isFinite(height) && height > 0) viewportHeightPx.value = height
}

function scrollTopPx(): number {
  const top = scrollRef.value?.scrollTop
  return typeof top === "number" && Number.isFinite(top) ? Math.max(0, top) : 0
}

function currentRenderRange(): PdfPageRange {
  const doc = renderDocument.value
  if (!doc || doc.pageCount <= 0) return { first: 0, last: -1 }
  // Print must rasterize every page: placeholders are not acceptable in the
  // print snapshot, so print mode replaces the lazy viewport window entirely.
  if (printMode.value) return { first: 0, last: doc.pageCount - 1 }
  const { offsets, heights } = pageLayout()
  return visiblePageRange(offsets, heights, {
    scrollTop: scrollTopPx(),
    viewportHeight: viewportHeightPx.value,
    overscanPages: PDF_PAGE_OVERSCAN,
  })
}

/** Revoke a slot's bitmap and drop it back to a placeholder. */
function releasePageBitmaps(pageIndexes: Iterable<number>): void {
  for (const pageIndex of pageIndexes) {
    const slot = pageSlots[pageIndex]
    if (!slot) continue
    if (slot.url) {
      URL.revokeObjectURL(slot.url)
      slot.url = null
    }
    slot.renderKey = ""
    slot.bitmapBytes = 0
    if (slot.state !== "pending") slot.state = "pending"
  }
}

/**
 * Pages left the render window: keep usable bitmaps in the byte-budgeted LRU so
 * scrolling back is instant, and release whatever the budget evicted.
 */
function retirePageBitmaps(pageIndexes: Iterable<number>): void {
  const renderKey = pageRenderKey()
  const evicted: number[] = []
  for (const pageIndex of pageIndexes) {
    const slot = pageSlots[pageIndex]
    if (!slot) continue
    if (!slot.url) {
      if (slot.state !== "pending") slot.state = "pending"
      continue
    }
    // Stale pixels can never be shown again and unmeasured bitmaps would sit in
    // the cache outside the budget, so both are released instead of retained.
    if (slot.renderKey !== renderKey || slot.bitmapBytes <= 0) {
      evicted.push(pageIndex)
      continue
    }
    // Already cached: keep its original recency so the LRU reflects when the
    // page left the window, not the order pages happen to be scanned in.
    if (bitmapCache.has(pageIndex)) continue
    evicted.push(...bitmapCache.retain(pageIndex, slot.bitmapBytes))
  }
  releasePageBitmaps(evicted)
}

/** The window owns its bitmaps while it shows them; only off-window pages count. */
function activatePageBitmaps(range: PdfPageRange): void {
  if (range.last < range.first) return
  for (let pageIndex = range.first; pageIndex <= range.last; pageIndex++) {
    bitmapCache.drop(pageIndex)
  }
}

/** Per-page raster plan for the current zoom, rotation and display scale. */
function pageRaster(pageIndex: number): PdfPageRasterPlan {
  const size = pageDisplaySize(pageIndex)
  return planPageRaster({
    displayWidth: size?.width ?? 0,
    displayHeight: size?.height ?? 0,
    zoom: zoom.value,
    dpr: renderDpr.value,
    maxPixels: props.maxPagePixels,
  })
}

function pageSlotSnapshots(): readonly PdfPageSlotSnapshot[] {
  const pageCount = renderDocument.value?.pageCount ?? 0
  const renderKey = pageRenderKey()
  const snapshots: PdfPageSlotSnapshot[] = []
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
    const slot = pageSlots[pageIndex]
    snapshots.push({
      state: slot?.state ?? "pending",
      hasBitmap: !!slot?.url,
      stale: !!slot?.url && slot.renderKey !== renderKey,
    })
  }
  return snapshots
}

function cancelScheduledRender(): void {
  if (renderTimer === undefined) return
  clearTimeout(renderTimer)
  renderTimer = undefined
}

function scheduleVisibleRender(delay = PDF_SCROLL_RENDER_DEBOUNCE_MS): void {
  cancelScheduledRender()
  renderTimer = setTimeout(() => {
    renderTimer = undefined
    void renderVisiblePages()
  }, delay)
}

// ── Runtime helpers ──────────────────────────────────────────────────
async function selectRuntime(): Promise<PdfRenderRuntime> {
  if (props.runtime) {
    if (ownedRuntime) {
      const rt = ownedRuntime
      ownedRuntime = undefined
      ownedRuntimeWasmUrl = undefined
      await rt.dispose()
    }
    return props.runtime
  }
  const wasmUrl = props.pdfiumWasmUrl?.trim() || undefined
  if (!ownedRuntime || ownedRuntimeWasmUrl !== wasmUrl) {
    if (ownedRuntime) await ownedRuntime.dispose()
    ownedRuntime = createPdfRenderRuntime({
      wasmUrl,
      password: props.password,
      onPassword: props.onPassword,
    })
    ownedRuntimeWasmUrl = wasmUrl
  }
  return ownedRuntime
}

function cancelRender(): void {
  renderGeneration++
  renderAbort?.abort()
  renderAbort = undefined
  cancelScheduledRender()
}

function syncPrintPageSize(): void {
  if (typeof document === "undefined" || typeof document.createElement !== "function") return
  const page = renderDocument.value?.pages[0]
  if (!page || page.width <= 0 || page.height <= 0) {
    printPageStyle?.remove?.()
    printPageStyle = undefined
    return
  }
  const quarterTurns = ((page.rotation + rotation.value) / 90) % 4
  const width = quarterTurns % 2 ? page.height : page.width
  const height = quarterTurns % 2 ? page.width : page.height
  if (!printPageStyle) {
    printPageStyle = document.createElement("style")
    if (printPageStyle.dataset) printPageStyle.dataset.officePrintPage = "pdf-surface"
    document.head?.append(printPageStyle)
  }
  printPageStyle.textContent = `@page { size: ${Math.round(width)}px ${Math.round(height)}px; margin: 0; }`
}

function setPrintMode(enabled: boolean): void {
  if (printMode.value === enabled) {
    if (enabled) scheduleVisibleRender(0)
    return
  }
  printMode.value = enabled
  if (enabled) scheduleVisibleRender(0)
  else scheduleVisibleRender(0)
}

function onPrintMediaChange(event: MediaQueryListEvent): void {
  setPrintMode(event.matches)
}

function onBeforePrint(): void {
  setPrintMode(true)
}

function onAfterPrint(): void {
  setPrintMode(false)
}

function setupPrintDetection(): void {
  if (typeof window === "undefined") return
  printMediaQuery = typeof window.matchMedia === "function" ? window.matchMedia("print") : undefined
  if (printMediaQuery) {
    if (typeof printMediaQuery.addEventListener === "function") {
      printMediaQuery.addEventListener("change", onPrintMediaChange)
    } else {
      printMediaQuery.addListener?.(onPrintMediaChange)
    }
    if (printMediaQuery.matches) setPrintMode(true)
  }
  if (typeof window.addEventListener === "function") {
    window.addEventListener("beforeprint", onBeforePrint)
    window.addEventListener("afterprint", onAfterPrint)
  }
  stopPrintDetection = () => {
    if (printMediaQuery) {
      if (typeof printMediaQuery.removeEventListener === "function") {
        printMediaQuery.removeEventListener("change", onPrintMediaChange)
      } else {
        printMediaQuery.removeListener?.(onPrintMediaChange)
      }
    }
    printMediaQuery = undefined
    if (typeof window.removeEventListener === "function") {
      window.removeEventListener("beforeprint", onBeforePrint)
      window.removeEventListener("afterprint", onAfterPrint)
    }
    printPageStyle?.remove?.()
    printPageStyle = undefined
  }
}

async function releaseDocument(): Promise<void> {
  cancelRender()
  surfaceSearch.clearSearch()
  clearSelection()
  geometryCache.clear()
  geometryRequests.clear()
  characterTextCache.clear()
  annotationRequests.clear()
  textLayerRequests.clear()
  for (const key of Object.keys(textLayers)) delete textLayers[Number(key)]
  formWritesAvailable.value = false
  for (const key of Object.keys(pageAnnotations)) delete pageAnnotations[Number(key)]
  for (const slot of Object.values(pageSlots)) {
    if (slot.url) URL.revokeObjectURL(slot.url)
  }
  bitmapCache.clear()
  for (const key of Object.keys(pageSlots)) delete pageSlots[Number(key)]
  invalidatePageLayout()
  renderedRange = { first: 0, last: -1 }
  visiblePageIndex = -1
  const doc = renderDocument.value
  const rt = runtimeForDocument
  renderDocument.value = null
  runtimeForDocument = undefined
  if (doc && rt) await rt.closeDocument(doc)
}

function createLinkedAbort(parent?: AbortSignal) {
  const controller = new AbortController()
  const abort = () => controller.abort()
  if (parent?.aborted) controller.abort()
  else parent?.addEventListener("abort", abort, { once: true })
  return { controller, unlink: () => parent?.removeEventListener("abort", abort) }
}

function toOfficeSource(source: PdfSource): OfficeSource {
  if (source.kind === "url") return { kind: "url", url: source.url }
  return {
    kind: "file",
    name: (source.blob as File).name ?? "document.pdf",
    file: {
      name: (source.blob as File).name ?? "document.pdf",
      size: source.blob.size,
      type: source.blob.type,
      arrayBuffer: () => source.blob.arrayBuffer(),
    },
  }
}

// ── Render the pages near the viewport sequentially ──────────────────
async function renderVisiblePages(): Promise<void> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt) return
  const gen = ++renderGeneration
  renderAbort?.abort()
  const { controller, unlink } = createLinkedAbort()
  renderAbort = controller
  const range = currentRenderRange()
  renderedRange = range
  const renderKey = pageRenderKey()
  const plan = planPageRenders(pageSlotSnapshots(), range)
  // The window takes its bitmaps out of the LRU (they are active), and the
  // pages that left it move into the byte-budgeted cache instead of vanishing.
  activatePageBitmaps(range)
  retirePageBitmaps(plan.release)
  // Text layers follow the same window so a long document cannot accumulate
  // glyph spans for every page ever visited; the runtime keeps the extracted
  // data cached for when the page returns.
  for (const key of Object.keys(textLayers)) {
    const pageIndex = Number(key)
    if (pageIndex >= range.first && pageIndex <= range.last) continue
    delete textLayers[pageIndex]
    textLayerRequests.delete(pageIndex)
  }

  try {
    for (const pageIndex of plan.render) {
      if (controller.signal.aborted || gen !== renderGeneration) return
      if (renderDocument.value !== doc || runtimeForDocument !== rt) return
      const slot = pageSlots[pageIndex]
      if (!slot || (slot.state === "ready" && slot.renderKey === renderKey)) continue
      // A stale bitmap stays mounted (and keeps its slot record) until the
      // replacement arrives; only pages without pixels show a loading state.
      if (!slot.url) slot.state = "loading"
      try {
        const raster = pageRaster(pageIndex)
        const blob = await rt.renderPage(doc, pageIndex, {
          scale: raster.scale,
          rotation: rotation.value,
          dpr: raster.dpr,
          signal: controller.signal,
        })
        if (controller.signal.aborted || gen !== renderGeneration) return
        if (renderDocument.value !== doc || runtimeForDocument !== rt) return
        if (pageSlots[pageIndex] !== slot) continue
        const url = URL.createObjectURL(blob)
        if (slot.url) URL.revokeObjectURL(slot.url)
        slot.url = url
        slot.state = "ready"
        slot.renderKey = renderKey
        slot.bitmapBytes = raster.bytes
        void ensurePageContentHash(pageIndex)
      } catch (err) {
        if ((err as Error)?.name === "AbortError" || controller.signal.aborted) return
        if (gen === renderGeneration && !slot.url) slot.state = "error"
      }
    }
    // Scrolling during a render pass can move the window past the pages just
    // rasterized; schedule one follow-up pass instead of leaving blanks.
    if (!controller.signal.aborted && gen === renderGeneration) {
      if (!pageRangeEquals(currentRenderRange(), renderedRange)) scheduleVisibleRender(0)
    }
  } finally {
    unlink()
    if (renderAbort === controller) renderAbort = undefined
  }
}

// ── Main load flow ───────────────────────────────────────────────────
async function load(): Promise<void> {
  const source = effectiveSource.value
  const requestId = ++latestRequestId
  pdfLoadCoordinator.cancel()
  loading.value = !!source
  loadError.value = null
  await releaseDocument()
  if (unmounted || requestId !== latestRequestId) return
  if (!source) { loading.value = false; return }

  const task = pdfLoadCoordinator.start(toOfficeSource(source), {
    urlPolicy: props.urlPolicy,
  })
  let openedDocument: PdfRenderDocument | undefined
  let selectedRuntime: PdfRenderRuntime | undefined
  try {
    const verified = await loadVerifiedPdfSource(source, props.urlPolicy, {
      signal: task.signal,
      maxFileSize: props.maxFileSize,
    })
    if (unmounted || requestId !== latestRequestId) return
    task.assertCurrent()
    selectedRuntime = await selectRuntime()
    if (unmounted || requestId !== latestRequestId) return
    task.assertCurrent()
    const content = verified.bytes.buffer.slice(
      verified.bytes.byteOffset,
      verified.bytes.byteOffset + verified.bytes.byteLength,
    ) as ArrayBuffer
    openedDocument = await selectedRuntime.openDocument(content, {
      signal: task.signal,
      password: props.password,
      onPassword: props.onPassword,
    })
    if (unmounted || requestId !== latestRequestId) {
      await selectedRuntime.closeDocument(openedDocument)
      openedDocument = undefined
      return
    }
    task.assertCurrent()
    if (openedDocument.pageCount <= 0) throw new Error(m.value.pdf.surface.noPdf)
    runtimeForDocument = selectedRuntime
    formWritesAvailable.value = typeof selectedRuntime.setFormFieldValue === "function"
    renderDocument.value = openedDocument
    revisionCounter.value += 1
    emit("documentRevisionChange", documentRevision.value)
    verifiedBlob.value = verified.blob
    initPageSlots()
    emit("document-load-success", openedDocument.pageCount)
    loading.value = false
    invalidatePageLayout()
    measureViewportHeight()
    scheduleVisibleRender(0)
    prefetchPageGeometry(0)
    prefetchPageAnnotations(0)
    prefetchPageTextLayer(0)
  } catch (cause) {
    if (openedDocument && selectedRuntime) {
      await selectedRuntime.closeDocument(openedDocument)
      openedDocument = undefined
    }
    if (unmounted || !task.isCurrent() || requestId !== latestRequestId || (cause as Error)?.name === "AbortError") return
    const sourceKind = source.kind
    const normalized = toPdfLoadError(cause, sourceKind, sourceKind === "url" ? source.url : undefined)
    const localizedError = {
      ...normalized,
      message: localizePdfErrorMessage(normalized, m.value.pdf.errors),
    }
    loadError.value = localizedError
    loading.value = false
    emit("document-load-error", localizedError)
  } finally {
    task.finish()
  }
}

// ── Watchers ─────────────────────────────────────────────────────────
watch([zoom, rotation, renderDpr], () => {
  if (!renderDocument.value) return
  cancelRender()
  // Every bitmap is stale at the new scale, rotation or display ratio: keep the
  // rasterized window on screen and let the next pass swap each bitmap in place.
  // Bitmaps cached outside the window cannot be reused at the new scale and are
  // released, so the cache never holds pixels from two scales at once.
  markPageBitmapsStale()
  invalidatePageLayout()
  scheduleVisibleRender(0)
})
watch([renderDocument, rotation], syncPrintPageSize, { immediate: true })

watch(
  [effectiveSource, () => props.urlPolicy, () => props.maxFileSize, () => props.runtime, () => props.pdfiumWasmUrl],
  () => {
    void load()
  },
  { immediate: true },
)

watch(() => props.documentId, () => emit("documentRevisionChange", documentRevision.value))

// A host that lowers the bitmap budget expects the excess to be released now.
watch(() => props.maxBitmapBytes, (next) => {
  releasePageBitmaps(bitmapCache.clear())
  bitmapCache = createPdfBitmapCache({ maxBytes: next })
  if (renderDocument.value) scheduleVisibleRender(0)
})

// ── Page geometry ────────────────────────────────────────────────────
function pageDisplaySize(pageIndex: number): { width: number; height: number } | undefined {
  const doc = renderDocument.value
  if (!doc) return undefined
  const page = doc.pages[pageIndex]
  if (!page) return undefined
  const size = pdfDisplaySize(page, rotation.value, zoom.value)
  return {
    width: Math.max(1, Math.round(size.width)),
    height: Math.max(1, Math.round(size.height)),
  }
}

function pageSlotStyle(pageIndex: number): Record<string, string> | undefined {
  const size = pageDisplaySize(pageIndex)
  if (!size) return undefined
  return {
    width: `${size.width}px`,
    height: `${size.height}px`,
  }
}

function pageImgStyle(pageIndex: number): Record<string, string> | undefined {
  const size = pageDisplaySize(pageIndex)
  if (!size) return undefined
  return {
    width: `${size.width}px`,
    height: `${size.height}px`,
    maxWidth: "100%",
    display: "block",
  }
}

function displayRect(pageIndex: number, rect: PdfRenderRect): PdfRenderRect {
  const page = renderDocument.value?.pages[pageIndex]
  if (!page) return rect
  return pdfCanonicalRectToDisplay(page, rect, rotation.value, zoom.value)
}

function displayRectStyle(pageIndex: number, rect: PdfRenderRect): Record<string, string> {
  const displayed = displayRect(pageIndex, rect)
  return {
    left: `${displayed.x}px`,
    top: `${displayed.y}px`,
    width: `${displayed.width}px`,
    height: `${displayed.height}px`,
  }
}

function searchRectsForPage(pageIndex: number): Array<{
  key: string
  resultIndex: number
  active: boolean
  rect: PdfRenderRect
}> {
  return searchState.value.matches.flatMap((match, resultIndex) =>
    match.pageIndex === pageIndex
      ? match.rects.map((rect, rectIndex) => ({
          key: `${resultIndex}:${rectIndex}`,
          resultIndex,
          active: resultIndex === searchState.value.activeIndex,
          rect,
        }))
      : [],
  )
}

function selectionRectsForPage(pageIndex: number): Array<{ key: string; rect: PdfRenderRect }> {
  const page = selectionVisual.value?.ranges.find((entry) => entry.pageIndex === pageIndex)
  return page ? page.rects.map((rect, rectIndex) => ({ key: `selection-${rectIndex}`, rect })) : []
}

function annotationHitsForPage(pageIndex: number): readonly PdfRenderAnnotation[] {
  return pageAnnotations[pageIndex] ?? []
}

async function pageAnnotationsFor(pageIndex: number): Promise<readonly PdfRenderAnnotation[]> {
  const cached = pageAnnotations[pageIndex]
  if (cached) return cached
  const pending = annotationRequests.get(pageIndex)
  if (pending) return pending
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt || typeof rt.getPageAnnotations !== "function") return []
  const read = rt.getPageAnnotations.bind(rt)
  let request: Promise<readonly PdfRenderAnnotation[]>
  request = read(doc, pageIndex)
    .then((annotations) => {
      if (renderDocument.value === doc && runtimeForDocument === rt) pageAnnotations[pageIndex] = annotations
      return annotations
    })
    .catch(() => [] as readonly PdfRenderAnnotation[])
    .finally(() => {
      if (annotationRequests.get(pageIndex) === request) annotationRequests.delete(pageIndex)
    })
  annotationRequests.set(pageIndex, request)
  return request
}

function pageSupportsAnnotations(): boolean {
  return typeof runtimeForDocument?.getPageAnnotations === "function"
}
function prefetchPageAnnotations(centerPageIndex: number): void {
  const doc = renderDocument.value
  if (!doc || !pageSupportsAnnotations()) return
  const first = Math.max(0, centerPageIndex - 1)
  const last = Math.min(doc.pageCount - 1, centerPageIndex + 1)
  for (let pageIndex = first; pageIndex <= last; pageIndex++) {
    if (!pageAnnotations[pageIndex]) void pageAnnotationsFor(pageIndex)
  }
}

// ── Accessible text layer ────────────────────────────────────────────
// One transparent, selectable DOM span per PDFium glyph. Spans stay in canonical
// page coordinates and the layer applies one CSS matrix for zoom/rotation, so the
// browser never re-shapes text to place it and screen readers, browser find and
// CJK extraction see the real characters instead of an aria-hidden bitmap.

function textLayerFor(pageIndex: number): PdfPageTextLayer | undefined {
  return textLayers[pageIndex]
}

function textLayerTransform(pageIndex: number): string {
  const page = renderDocument.value?.pages[pageIndex]
  if (!page) return ""
  return pdfCanonicalToDisplayTransform(page, rotation.value, zoom.value)
}

function textLayerStyle(pageIndex: number): Record<string, string> | undefined {
  const layer = textLayers[pageIndex]
  if (!layer) return undefined
  return {
    width: `${layer.width}px`,
    height: `${layer.height}px`,
    transform: textLayerTransform(pageIndex),
  }
}

function textGlyphStyle(glyph: PdfPageTextLayerGlyph): Record<string, string> {
  return {
    left: `${glyph.rect.x}px`,
    top: `${glyph.rect.y}px`,
    width: `${glyph.rect.width}px`,
    height: `${glyph.rect.height}px`,
    fontSize: `${glyph.rect.height}px`,
  }
}

function pageSupportsTextLayer(): boolean {
  return typeof runtimeForDocument?.getPageTextLayer === "function"
}

async function pageTextLayer(pageIndex: number): Promise<PdfPageTextLayer | undefined> {
  const cached = textLayers[pageIndex]
  if (cached) return cached
  const pending = textLayerRequests.get(pageIndex)
  if (pending) return pending
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt || !pageSupportsTextLayer()) return undefined
  const read = rt.getPageTextLayer!.bind(rt)
  let request: Promise<PdfPageTextLayer | undefined>
  request = read(doc, pageIndex)
    .then((layer) => {
      if (renderDocument.value === doc && runtimeForDocument === rt) textLayers[pageIndex] = layer
      return layer
    })
    .catch(() => undefined)
    .finally(() => {
      if (textLayerRequests.get(pageIndex) === request) textLayerRequests.delete(pageIndex)
    })
  textLayerRequests.set(pageIndex, request)
  return request
}

function prefetchPageTextLayer(centerPageIndex: number): void {
  const doc = renderDocument.value
  if (!doc || !pageSupportsTextLayer()) return
  const first = Math.max(0, centerPageIndex - 1)
  const last = Math.min(doc.pageCount - 1, centerPageIndex + 1)
  for (let pageIndex = first; pageIndex <= last; pageIndex++) {
    if (!textLayers[pageIndex]) void pageTextLayer(pageIndex)
  }
}

function invalidatePageTextLayer(pageIndex: number): void {
  delete textLayers[pageIndex]
  textLayerRequests.delete(pageIndex)
  if (pageSupportsTextLayer()) void pageTextLayer(pageIndex)
}

/** Widget kinds the Surface renders a native control for. */
const FORM_CONTROL_FIELD_TYPES: ReadonlySet<PdfFormFieldType> = new Set<PdfFormFieldType>([
  "text",
  "checkbox",
  "radio",
  "combobox",
  "listbox",
])

function formControlsForPage(pageIndex: number): readonly PdfRenderAnnotation[] {
  return annotationHitsForPage(pageIndex).filter(
    (annotation) => annotation.kind === "widget"
      && annotation.fieldType !== undefined
      && FORM_CONTROL_FIELD_TYPES.has(annotation.fieldType),
  )
}

function isFormFieldChecked(annotation: PdfRenderAnnotation): boolean {
  if (annotation.exportValue) return annotation.fieldValue === annotation.exportValue
  return Boolean(annotation.fieldValue)
}

/** Replace one annotation in the page cache after a write. */
function replacePageAnnotation(pageIndex: number, annotation: PdfRenderAnnotation): void {
  const current = pageAnnotations[pageIndex] ?? []
  pageAnnotations[pageIndex] = current.map((candidate) =>
    candidate.id === annotation.id ? annotation : candidate,
  )
}

/** Re-rasterize one page after its content changed, keeping the old bitmap visible. */
function refreshPageBitmap(pageIndex: number): void {
  const slot = pageSlots[pageIndex]
  if (!slot) return
  // The mounted bitmap stays until the replacement lands; an empty render key
  // marks it stale so the next plan re-renders this page.
  slot.renderKey = ""
  scheduleVisibleRender(0)
}

function formEventValue(event: Event): string {
  const target = event.target as { value?: unknown } | null
  return typeof target?.value === "string" ? target.value : ""
}

function formEventChecked(event: Event): boolean {
  const target = event.target as { checked?: unknown; value?: unknown } | null
  if (typeof target?.checked === "boolean") return target.checked
  return formEventValue(event) !== ""
}

async function commitFormFieldValue(
  pageIndex: number,
  annotation: PdfRenderAnnotation,
  value: PdfFormFieldValue,
): Promise<PdfRenderAnnotation | undefined> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt || typeof rt.setFormFieldValue !== "function") return undefined
  try {
    const updated = await rt.setFormFieldValue(doc, pageIndex, annotation.id, value)
    if (renderDocument.value !== doc || runtimeForDocument !== rt) return undefined
    replacePageAnnotation(pageIndex, updated)
    refreshPageBitmap(pageIndex)
    invalidatePageTextLayer(pageIndex)
    emit("formFieldChange", {
      pageIndex,
      annotationId: updated.id,
      annotation: updated,
      ...(updated.fieldName ? { fieldName: updated.fieldName } : {}),
      ...(updated.fieldType ? { fieldType: updated.fieldType } : {}),
      ...(updated.fieldValue !== undefined ? { value: updated.fieldValue } : {}),
    })
    return updated
  } catch (reason) {
    emit("formFieldError", {
      pageIndex,
      annotationId: annotation.id,
      message: reason instanceof Error ? reason.message : String(reason),
    })
    throw reason
  }
}

async function onFormTextChange(event: Event, pageIndex: number, annotation: PdfRenderAnnotation): Promise<void> {
  try {
    await commitFormFieldValue(pageIndex, annotation, formEventValue(event))
  } catch {
    // The failure was reported through `formFieldError`; keep the control editable.
  }
}

async function onFormSelectChange(event: Event, pageIndex: number, annotation: PdfRenderAnnotation): Promise<void> {
  try {
    await commitFormFieldValue(pageIndex, annotation, formEventValue(event))
  } catch {
    // The failure was reported through `formFieldError`.
  }
}

async function onFormToggleChange(event: Event, pageIndex: number, annotation: PdfRenderAnnotation): Promise<void> {
  try {
    await commitFormFieldValue(pageIndex, annotation, formEventChecked(event))
  } catch {
    // The failure was reported through `formFieldError`.
  }
}

async function setFormFieldValue(
  pageIndex: number,
  annotationId: string,
  value: PdfFormFieldValue,
): Promise<PdfRenderAnnotation | undefined> {
  const annotations = pageAnnotations[pageIndex] ?? await pageAnnotationsFor(pageIndex)
  const annotation = annotations.find((candidate) => candidate.id === annotationId)
  if (!annotation) {
    throw new Error(`PDF annotation ${annotationId} is not loaded for page ${pageIndex + 1}.`)
  }
  return commitFormFieldValue(pageIndex, annotation, value)
}

async function updateAnnotationAppearance(
  pageIndex: number,
  annotationId: string,
  patch: PdfAnnotationAppearancePatch,
): Promise<PdfRenderAnnotation | undefined> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt || typeof rt.updateAnnotationAppearance !== "function") {
    throw new Error("PDF runtime does not support annotation appearance updates.")
  }
  try {
    const updated = await rt.updateAnnotationAppearance(doc, pageIndex, annotationId, patch)
    if (renderDocument.value !== doc || runtimeForDocument !== rt) return undefined
    replacePageAnnotation(pageIndex, updated)
    refreshPageBitmap(pageIndex)
    emit("annotationChange", { pageIndex, annotationId, annotation: updated, patch })
    return updated
  } catch (reason) {
    emit("annotationError", {
      pageIndex,
      annotationId,
      message: reason instanceof Error ? reason.message : String(reason),
    })
    throw reason
  }
}

async function activatePdfSearchHit(
  match: PdfSearchHit,
  index: number,
  signal: AbortSignal,
): Promise<void> {
  await nextTick()
  if (signal.aborted) return
  const container = scrollRef.value
  const slot = container?.querySelector<HTMLElement>(`[data-page-index="${match.pageIndex}"]`)
  if (!container || !slot) throw new Error(`PDF page ${match.pageIndex + 1} is not mounted.`)
  const hit = slot.querySelector<HTMLElement>(`[data-search-result-index="${index}"]`)
  ;(hit ?? slot).scrollIntoView({ behavior: "auto", block: "center", inline: "nearest" })
}

async function pageGeometry(pageIndex: number): Promise<PdfPageGeometry> {
  const cached = geometryCache.get(pageIndex)
  if (cached) return cached
  const pending = geometryRequests.get(pageIndex)
  if (pending) return pending
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt) throw new Error("PDF document is not ready.")
  let request: Promise<PdfPageGeometry>
  request = rt.getPageTextGeometry(doc, pageIndex)
    .then((geometry) => {
      if (renderDocument.value === doc && runtimeForDocument === rt) {
        geometryCache.set(pageIndex, geometry)
      }
      return geometry
    })
    .finally(() => {
      if (geometryRequests.get(pageIndex) === request) geometryRequests.delete(pageIndex)
    })
  geometryRequests.set(pageIndex, request)
  return request
}

function prefetchPageGeometry(centerPageIndex: number): void {
  const doc = renderDocument.value
  if (!doc) return
  const first = Math.max(0, centerPageIndex - 1)
  const last = Math.min(doc.pageCount - 1, centerPageIndex + 1)
  for (let pageIndex = first; pageIndex <= last; pageIndex++) {
    void pageGeometry(pageIndex).catch(() => undefined)
  }
}

function eventClientPoint(event: MouseEvent | PointerEvent): ClientPoint {
  return { x: event.clientX, y: event.clientY }
}

function canonicalPoint(point: ClientPoint, pageIndex: number, layer: HTMLElement): Position | null {
  const page = renderDocument.value?.pages[pageIndex]
  if (!page) return null
  const rect = layer.getBoundingClientRect()
  return pdfDisplayPointToCanonical(
    page,
    { x: point.x - rect.left, y: point.y - rect.top },
    rotation.value,
    zoom.value,
  )
}

function resolveSelectionGlyph(
  geometry: PdfPageGeometry,
  point: Position,
  previousFocus = -1,
): number {
  const direct = glyphAt(geometry, point)
  return direct >= 0
    ? direct
    : nearestGlyphWithinContainingRun(geometry, point, previousFocus)
}

function selectionRects(geometry: PdfPageGeometry, anchor: number, focus: number): PdfRenderRect[] {
  const from = Math.min(anchor, focus)
  const to = Math.max(anchor, focus)
  return rectsWithinSlice(geometry, from, to).map((rect) => ({
    x: rect.origin.x,
    y: rect.origin.y,
    width: rect.size.width,
    height: rect.size.height,
  }))
}

/**
 * Nearest character on the page, used when a cross-page drag enters a margin or
 * a blank area instead of a text run.
 */
function clampSelectionChar(geometry: PdfPageGeometry, point: Position): number {
  const total = pdfPageCharCount(geometry)
  if (total <= 0) return -1
  let bestRun: PdfPageGeometry["runs"][number] | undefined
  let bestDistance = Number.POSITIVE_INFINITY
  for (const run of geometry.runs) {
    if (run.glyphs.length === 0) continue
    const bottom = run.rect.y + run.rect.height
    const distance = point.y < run.rect.y ? run.rect.y - point.y : point.y > bottom ? point.y - bottom : 0
    if (distance < bestDistance) {
      bestDistance = distance
      bestRun = run
    }
  }
  if (!bestRun) return -1
  const before = point.x < bestRun.rect.x + bestRun.rect.width / 2
  return before ? bestRun.charStart : bestRun.charStart + bestRun.glyphs.length - 1
}

function resolveGesturePoint(
  gesture: NonNullable<typeof selectionGesture>,
  point: ClientPoint,
  pageIndex: number,
  layer: HTMLElement,
  previous?: PdfSelectionPoint,
  clampToPage = false,
): PdfSelectionPoint | null {
  const geometry = gesture.geometries.get(pageIndex)
  if (!geometry) return null
  const canonical = canonicalPoint(point, pageIndex, layer)
  if (!canonical) return null
  const previousFocus = previous?.pageIndex === pageIndex ? previous.charIndex : -1
  let charIndex = resolveSelectionGlyph(geometry, canonical, previousFocus)
  if (charIndex < 0 && clampToPage) charIndex = clampSelectionChar(geometry, canonical)
  return charIndex >= 0 ? { pageIndex, charIndex } : null
}

/** Page layer under a client point. Pointer capture keeps events on the anchor page. */
function selectionPageAtPoint(point: ClientPoint): { pageIndex: number; layer: HTMLElement } | null {
  const ownerDocument = scrollRef.value?.ownerDocument
  const element = ownerDocument?.elementFromPoint?.(point.x, point.y) as Element | null | undefined
  const slot = element?.closest?.<HTMLElement>("[data-page-index]")
  if (!slot) return null
  const pageIndex = Number(slot.dataset.pageIndex)
  if (!Number.isSafeInteger(pageIndex) || pageIndex < 0) return null
  const layer = slot.querySelector<HTMLElement>(".pdf-surface__interaction") ?? slot
  return { pageIndex, layer }
}

function cancelSelectionFrame(): void {
  if (selectionFrame === undefined) return
  if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(selectionFrame)
  selectionFrame = undefined
}

function stopSelectionGesture(): void {
  selectionGesture = null
  cancelSelectionFrame()
}

function resetCommittedSelection(clearVisual: boolean): void {
  if (clearVisual) selectionVisual.value = null
  if (selectionState.value.kind === "none") return
  selectionState.value = { kind: "none" }
  emit("selectionChange", selectionState.value)
}

function beginGestureSelection(gesture: NonNullable<typeof selectionGesture>): void {
  if (gesture.started) return
  gesture.started = true
  resetCommittedSelection(true)
}

function updateGesturePoint(
  gesture: NonNullable<typeof selectionGesture>,
  point: ClientPoint,
): void {
  gesture.latest = point
  if (gesture.thresholdExceeded) return
  if (Math.hypot(point.x - gesture.down.x, point.y - gesture.down.y) < PDF_SELECTION_DRAG_THRESHOLD_PX) return
  gesture.thresholdExceeded = true
  beginGestureSelection(gesture)
}

function selectionRangesForGesture(gesture: NonNullable<typeof selectionGesture>): PageTextSlice[] {
  if (!gesture.anchor) return []
  const focus = gesture.focus ?? gesture.anchor
  return selectionPageRanges(
    gesture.anchor,
    focus,
    (pageIndex) => pdfPageCharCount(gesture.geometries.get(pageIndex)),
  )
}

function buildSelectionVisual(gesture: NonNullable<typeof selectionGesture>): PageSelectionRects[] {
  const pages: PageSelectionRects[] = []
  for (const range of selectionRangesForGesture(gesture)) {
    const geometry = gesture.geometries.get(range.pageIndex)
    if (!geometry) continue
    const rects = selectionRects(geometry, range.charIndex, range.charIndex + range.charCount - 1)
    if (rects.length > 0) pages.push({ pageIndex: range.pageIndex, rects })
  }
  return pages
}

function requestGestureGeometry(
  gesture: NonNullable<typeof selectionGesture>,
  pageIndex: number,
): void {
  if (gesture.geometries.has(pageIndex) || gesture.pendingGeometries.has(pageIndex)) return
  gesture.pendingGeometries.add(pageIndex)
  void pageGeometry(pageIndex)
    .then((geometry) => {
      gesture.pendingGeometries.delete(pageIndex)
      if (selectionGesture !== gesture) return
      gesture.geometries.set(pageIndex, geometry)
      flushSelectionGesture(gesture)
      if (!gesture.pointerActive) finalizeSelectionGesture(gesture)
    })
    .catch(() => {
      gesture.pendingGeometries.delete(pageIndex)
      if (selectionGesture !== gesture) return
      if (pageIndex === gesture.pageIndex) {
        if (gesture.layer.hasPointerCapture?.(gesture.pointerId)) {
          gesture.layer.releasePointerCapture?.(gesture.pointerId)
        }
        clearSelection()
        return
      }
      if (!gesture.pointerActive) finalizeSelectionGesture(gesture)
    })
}

function flushSelectionGesture(gesture: NonNullable<typeof selectionGesture>): void {
  if (selectionGesture !== gesture || !gesture.thresholdExceeded) return
  beginGestureSelection(gesture)

  if (!gesture.anchor) {
    gesture.anchor = resolveGesturePoint(gesture, gesture.down, gesture.pageIndex, gesture.layer) ?? undefined
    if (!gesture.anchor) {
      requestGestureGeometry(gesture, gesture.pageIndex)
      return
    }
  }

  const target = selectionPageAtPoint(gesture.latest)
  const focusPageIndex = target?.pageIndex ?? gesture.pageIndex
  const focusLayer = target?.layer ?? gesture.layer
  const previousFocus = gesture.focus?.pageIndex === focusPageIndex ? gesture.focus : undefined
  const resolved = resolveGesturePoint(
    gesture,
    gesture.latest,
    focusPageIndex,
    focusLayer,
    previousFocus,
    focusPageIndex !== gesture.pageIndex,
  )
  if (resolved) gesture.focus = resolved
  else requestGestureGeometry(gesture, focusPageIndex)

  const anchor = gesture.anchor
  const focus = gesture.focus ?? anchor
  const measured = [...gesture.geometries.keys()].sort((a, b) => a - b).join(",")
  const visualKey = `${anchor.pageIndex}:${anchor.charIndex}|${focus.pageIndex}:${focus.charIndex}|${measured}`
  if (visualKey === gesture.visualKey) return
  gesture.visualKey = visualKey
  selectionVisual.value = { ranges: buildSelectionVisual(gesture) }
}

function scheduleSelectionFrame(gesture: NonNullable<typeof selectionGesture>): void {
  if (selectionFrame !== undefined) return
  if (typeof requestAnimationFrame !== "function") {
    flushSelectionGesture(gesture)
    return
  }
  selectionFrame = requestAnimationFrame(() => {
    selectionFrame = undefined
    if (selectionGesture === gesture) flushSelectionGesture(gesture)
  })
}

function setSelectionUnavailable(pageIndex: number): void {
  const unavailable: PdfSelectionState = { kind: "unavailable", pageIndex, reason: "empty" }
  selectionState.value = unavailable
  selectionVisual.value = null
  emit("selectionChange", unavailable)
}

async function commitSelection(
  anchor: PdfSelectionPoint,
  focus: PdfSelectionPoint,
  geometries: Map<number, PdfPageGeometry>,
  request: number,
): Promise<void> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt) return
  try {
    for (const pageIndex of selectionPageIndexes(anchor, focus)) {
      if (geometries.has(pageIndex)) continue
      geometries.set(pageIndex, await pageGeometry(pageIndex))
      if (request !== selectionRequest || renderDocument.value !== doc || runtimeForDocument !== rt) return
    }
    const ranges = selectionPageRanges(
      anchor,
      focus,
      (pageIndex) => pdfPageCharCount(geometries.get(pageIndex)),
    )
    const texts = ranges.length > 0 ? await rt.getTextSlices(doc, ranges) : []
    if (request !== selectionRequest || renderDocument.value !== doc || runtimeForDocument !== rt) return
    if (ranges.length === 0 || texts.length !== ranges.length) {
      setSelectionUnavailable(anchor.pageIndex)
      return
    }

    const rectsByPage = new Map<number, readonly PdfRenderRect[]>()
    for (const range of ranges) {
      const geometry = geometries.get(range.pageIndex)
      if (!geometry) continue
      rectsByPage.set(
        range.pageIndex,
        selectionRects(geometry, range.charIndex, range.charIndex + range.charCount - 1),
      )
    }
    const visuals = ranges.flatMap((range) => {
      const rects = rectsByPage.get(range.pageIndex) ?? []
      return rects.length > 0 ? [{ pageIndex: range.pageIndex, rects }] : []
    })
    selectionVisual.value = visuals.length > 0 ? { ranges: visuals } : null

    const invalidIndex = texts.findIndex((text) => hasUnmappedPdfText(text))
    if (invalidIndex >= 0) {
      const unavailable: PdfSelectionState = {
        kind: "unavailable",
        pageIndex: ranges[invalidIndex].pageIndex,
        reason: "unicode-map",
      }
      selectionState.value = unavailable
      selectionVisual.value = null
      emit("selectionChange", unavailable)
      return
    }

    const segments: PdfSelectionSegment[] = ranges.map((range, index) => ({
      range,
      text: texts[index] ?? "",
      rects: rectsByPage.get(range.pageIndex) ?? [],
    }))
    if (segments.every((segment) => segment.text.length === 0)) {
      setSelectionUnavailable(anchor.pageIndex)
      return
    }
    const first = segments[0]
    const next: PdfSelectionState = {
      kind: "text",
      range: first.range,
      text: joinSelectionTexts(segments.map((segment) => segment.text)),
      rects: first.rects,
      segments,
    }
    selectionState.value = next
    emit("selectionChange", next)
    // PDF references address one page; multi-page selections stay copy/selection
    // only until the reference model can express a page span.
    if (segments.length === 1 && selectionMode.value === "content") {
      try {
        const context = referenceContext()
        await ensurePageContentHash(first.range.pageIndex)
        if (renderDocument.value?.id !== context.document.id) return
        const draft = createPdfTextReferenceDraft(context, next)
        confirmDraft(draft, "pointer", {
          label: next.text.length > 80 ? `${next.text.slice(0, 77)}…` : next.text,
          path: [{ kind: "document", label: descriptor("pdf") }, { kind: "page", label: descriptor("page", { number: next.range.pageIndex + 1 }) }, { kind: "text-range", label: m.value.common.selectedText }],
          content: { text: next.text },
        })
      } catch (reason) {
        emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : formatOfficeMessage(m.value.common.reference.createFailed, { format: "PDF" }))
      }
    }
  } catch {
    if (request === selectionRequest) setSelectionUnavailable(anchor.pageIndex)
  }
}

function finalizeSelectionGesture(gesture: NonNullable<typeof selectionGesture>): void {
  if (selectionGesture !== gesture) return
  cancelSelectionFrame()
  flushSelectionGesture(gesture)
  // Wait for the anchor and focus page geometry so a cross-page commit cannot
  // fall back to a single-page range that omits the page under the pointer.
  if (!gesture.anchor || gesture.pendingGeometries.size > 0) return
  selectionGesture = null
  void commitSelection(gesture.anchor, gesture.focus ?? gesture.anchor, gesture.geometries, gesture.request)
}

function onSelectionPointerDown(event: PointerEvent, pageIndex: number): void {
  if (!event.isPrimary || event.button !== 0) return
  const layer = event.currentTarget as HTMLElement
  scrollRef.value?.focus()
  stopSelectionGesture()
  selectionClickCandidate = false
  const request = ++selectionRequest
  const point = eventClientPoint(event)
  const gesture: NonNullable<typeof selectionGesture> = {
    request,
    pageIndex,
    pointerId: event.pointerId,
    layer,
    down: point,
    latest: point,
    pointerActive: true,
    thresholdExceeded: false,
    started: false,
    geometries: new Map(),
    pendingGeometries: new Set(),
  }
  selectionGesture = gesture
  layer.setPointerCapture(event.pointerId)
  requestGestureGeometry(gesture, pageIndex)
}

function onSelectionPointerMove(event: PointerEvent, _pageIndex: number): void {
  const gesture = selectionGesture
  // Pointer capture keeps events on the anchor page; the page under the pointer
  // is resolved from the client position so a drag can cross page boundaries.
  if (!gesture || gesture.pointerId !== event.pointerId) return
  updateGesturePoint(gesture, eventClientPoint(event))
  if (gesture.thresholdExceeded) scheduleSelectionFrame(gesture)
}

function onSelectionPointerUp(event: PointerEvent, _pageIndex: number): void {
  const gesture = selectionGesture
  const layer = event.currentTarget as HTMLElement
  if (!gesture || gesture.pointerId !== event.pointerId) return
  updateGesturePoint(gesture, eventClientPoint(event))
  gesture.pointerActive = false
  if (layer.hasPointerCapture?.(event.pointerId)) layer.releasePointerCapture?.(event.pointerId)
  if (!gesture.thresholdExceeded) {
    stopSelectionGesture()
    clearSelection()
    selectionClickCandidate = true
    return
  }
  selectionClickCandidate = false
  finalizeSelectionGesture(gesture)
}

async function characterTextsForRange(
  pageIndex: number,
  from: number,
  to: number,
  request: number,
): Promise<readonly string[] | null> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt) return null
  let pageCache = characterTextCache.get(pageIndex)
  if (!pageCache) {
    pageCache = new Map<number, string>()
    characterTextCache.set(pageIndex, pageCache)
  }

  const missing: PageTextSlice[] = []
  for (let charIndex = from; charIndex <= to; charIndex++) {
    if (!pageCache.has(charIndex)) missing.push({ pageIndex, charIndex, charCount: 1 })
  }
  if (missing.length > 0) {
    const texts = await rt.getTextSlices(doc, missing)
    if (request !== selectionRequest || renderDocument.value !== doc || runtimeForDocument !== rt) return null
    for (let index = 0; index < missing.length; index++) {
      pageCache.set(missing[index].charIndex, texts[index] ?? "")
    }
  }

  return Array.from({ length: to - from + 1 }, (_, index) => pageCache!.get(from + index) ?? "")
}

function isPdfLinkUrlAllowed(url: string): boolean {
  const policy = props.urlPolicy
  if (!policy || policy.enabled === false) return true
  try {
    resolveOfficeUrl(url, policy)
    return true
  } catch {
    return false
  }
}

function annotationAtPoint(
  annotations: readonly PdfRenderAnnotation[],
  point: Position,
): PdfRenderAnnotation | undefined {
  for (let index = annotations.length - 1; index >= 0; index--) {
    const annotation = annotations[index]
    const rect = annotation.rect
    if (rect.width <= 0 || rect.height <= 0) continue
    if (
      point.x >= rect.x && point.x <= rect.x + rect.width &&
      point.y >= rect.y && point.y <= rect.y + rect.height
    ) {
      return annotation
    }
  }
  return undefined
}

/**
 * Forward a click that landed on a link, form widget or annotation rectangle to
 * the host. The raster page bakes those objects into the image, so the Surface
 * owns the hit test and never navigates by itself.
 */
async function activateAnnotationAt(event: MouseEvent, pageIndex: number): Promise<void> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt || !pageSupportsAnnotations()) return
  const layer = event.currentTarget as HTMLElement | null
  if (!layer) return
  const point = canonicalPoint(eventClientPoint(event), pageIndex, layer)
  if (!point) return
  const annotations = pageAnnotations[pageIndex] ?? await pageAnnotationsFor(pageIndex)
  if (renderDocument.value !== doc || runtimeForDocument !== rt) return
  const annotation = annotationAtPoint(annotations, point)
  if (!annotation) return
  event.preventDefault?.()
  if (annotation.kind === "link") {
    const blocked = annotation.blocked === true
      || (annotation.url !== undefined && !isPdfLinkUrlAllowed(annotation.url))
    emit("linkActivate", {
      pageIndex,
      annotationId: annotation.id,
      clientX: event.clientX,
      clientY: event.clientY,
      ...(!blocked && annotation.url ? { url: annotation.url } : {}),
      ...(annotation.targetPageIndex !== undefined ? { targetPageIndex: annotation.targetPageIndex } : {}),
      ...(blocked ? { blocked: true } : {}),
    })
    if (!blocked && annotation.targetPageIndex !== undefined) scrollToPage(annotation.targetPageIndex)
    return
  }
  emit("annotationActivate", {
    pageIndex,
    annotationId: annotation.id,
    kind: annotation.kind,
    clientX: event.clientX,
    clientY: event.clientY,
    ...(annotation.fieldName ? { fieldName: annotation.fieldName } : {}),
    ...(annotation.fieldValue ? { fieldValue: annotation.fieldValue } : {}),
    ...(annotation.contents ? { contents: annotation.contents } : {}),
  })
}

async function onSelectionClick(event: MouseEvent, pageIndex: number): Promise<void> {
  const isNativeDoubleClick = selectionClickCandidate && event.button === 0 && event.detail === 2
  selectionClickCandidate = false
  if (!isNativeDoubleClick) {
    if (event.button === 0 && (event.detail ?? 0) <= 1) await activateAnnotationAt(event, pageIndex)
    return
  }
  event.preventDefault()

  const layer = event.currentTarget as HTMLElement
  const request = ++selectionRequest
  resetCommittedSelection(true)
  try {
    const geometry = await pageGeometry(pageIndex)
    if (request !== selectionRequest) return
    const point = canonicalPoint(eventClientPoint(event), pageIndex, layer)
    const charIndex = point ? resolveSelectionGlyph(geometry, point) : -1
    if (charIndex < 0) return
    const candidate = expandToWordBoundary(geometry, charIndex)
    if (!candidate) return
    const texts = await characterTextsForRange(pageIndex, candidate.from, candidate.to, request)
    const Segmenter = (Intl as typeof Intl & { Segmenter?: PdfWordSegmenterConstructor }).Segmenter
    if (!texts || request !== selectionRequest || !Segmenter) return
    const range = wordRangeFromCharacterTexts(
      candidate.from,
      charIndex,
      texts,
      new Segmenter(undefined, { granularity: "word" }),
    )
    if (!range) return
    await commitSelection(
      { pageIndex, charIndex: range.from },
      { pageIndex, charIndex: range.to },
      new Map([[pageIndex, geometry]]),
      request,
    )
  } catch {
    // A failed word lookup must not revive a stale or partial selection.
  }
}

function onSelectionPointerCancel(event: PointerEvent): void {
  if (selectionGesture?.pointerId !== event.pointerId) return
  selectionClickCandidate = false
  clearSelection()
}

function clearSelection(): void {
  selectionRequest += 1
  selectionClickCandidate = false
  stopSelectionGesture()
  resetCommittedSelection(true)
}

function onCopy(event: ClipboardEvent): void {
  if (selectionState.value.kind !== "text" || !event.clipboardData) return
  event.preventDefault()
  event.clipboardData.setData("text/plain", selectionState.value.text)
}

// ── Scroll tracking ──────────────────────────────────────────────────
function onScroll(): void {
  if (!renderDocument.value) return
  if (scrollFrame !== undefined) return
  if (typeof requestAnimationFrame !== "function") {
    flushScrollFrame()
    return
  }
  scrollFrame = requestAnimationFrame(() => {
    scrollFrame = undefined
    flushScrollFrame()
  })
}

function cancelScrollFrame(): void {
  if (scrollFrame === undefined) return
  if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(scrollFrame)
  scrollFrame = undefined
}

function flushScrollFrame(): void {
  const doc = renderDocument.value
  if (!doc || !scrollRef.value) return
  // A container without a ResizeObserver keeps its mount-time height; re-read it
  // once there is no cached value instead of on every frame.
  if (viewportHeightPx.value <= 0) measureViewportHeight()
  const { offsets, heights } = pageLayout()
  const pageIndex = pageIndexAtPosition(
    offsets,
    heights,
    scrollTopPx() + viewportHeightPx.value / 2,
  )
  if (pageIndex !== visiblePageIndex) {
    visiblePageIndex = pageIndex
    emit("visible-page-change", pageIndex)
    prefetchPageGeometry(pageIndex)
    prefetchPageAnnotations(pageIndex)
    prefetchPageTextLayer(pageIndex)
  }
  if (!pageRangeEquals(currentRenderRange(), renderedRange)) {
    scheduleVisibleRender(PDF_SCROLL_RENDER_DEBOUNCE_MS)
  }
}

function emitContextMenuAt(clientX: number, clientY: number): void {
  const el = scrollRef.value
  if (!el || !renderDocument.value) return
  const rect = el.getBoundingClientRect()
  const clickY = clientY - rect.top + el.scrollTop
  const { offsets, heights } = pageLayout()
  const pageIndex = pageIndexAtPosition(offsets, heights, clickY)
  emit("contextMenu", {
    pageIndex,
    clientX,
    clientY,
    containerX: clientX - rect.left,
    containerY: clientY - rect.top,
  })
}

function onContextMenu(event: MouseEvent): void {
  if (touchGestures.shouldSuppressContextMenu(event.clientX, event.clientY)) {
    event.preventDefault()
    return
  }
  emitContextMenuAt(event.clientX, event.clientY)
}

// ── Gesture zoom ─────────────────────────────────────────────────────
type PageAnchor = {
  pageIndex: number
  xRatio: number
  yRatio: number
  clientX: number
  clientY: number
}
type WebKitGestureEvent = Event & { clientX?: number; clientY?: number; scale?: number }

let pendingAnchor: (PageAnchor & { requestedZoom: number; token: number }) | undefined
let pendingZoom: number | undefined
let gestureStartZoom = 1
let gestureToken = 0
let webkitGestureActive = false

function controlledZoom(): number | undefined {
  return typeof props.zoom === "number" && Number.isFinite(props.zoom)
    ? clampSurfaceZoom(props.zoom)
    : undefined
}

const touchGestures = createTouchGestureTracker({
  isEnabled: () => selectionMode.value === "content",
  isDoubleTapEnabled: () => controlledZoom() !== undefined,
  getZoom: controlledZoom,
  onZoom: (zoom, center) => requestGestureZoom(zoom, center.clientX, center.clientY),
  onDoubleTap: (pointer) => {
    const current = controlledZoom()
    if (current === undefined) return
    requestGestureZoom(nextDoubleTapZoom(current), pointer.clientX, pointer.clientY)
  },
  onLongPress: (pointer) => emitContextMenuAt(pointer.clientX, pointer.clientY),
})

function onTouchPointerDown(event: PointerEvent): void { touchGestures.pointerDown(event) }
function onTouchPointerMove(event: PointerEvent): void { touchGestures.pointerMove(event) }
function onTouchPointerUp(event: PointerEvent): void { touchGestures.pointerUp(event) }
function onTouchPointerCancel(event: PointerEvent): void { touchGestures.pointerCancel(event) }

function captureZoomAnchor(clientX: number, clientY: number): PageAnchor | undefined {
  if (typeof document === "undefined" || typeof document.createElement !== "function") return undefined
  const page = document.elementFromPoint?.(clientX, clientY)
    ?.closest<HTMLElement>("[data-page-index]")
  const pageIndex = Number(page?.dataset.pageIndex)
  if (!page || !Number.isInteger(pageIndex)) return undefined
  const rect = page.getBoundingClientRect()
  if (rect.width <= 0 || rect.height <= 0) return undefined
  return {
    pageIndex,
    xRatio: (clientX - rect.left) / rect.width,
    yRatio: (clientY - rect.top) / rect.height,
    clientX,
    clientY,
  }
}

async function restoreZoomAnchor(anchor: PageAnchor & { token: number }): Promise<void> {
  await nextTick()
  requestAnimationFrame(() => {
    if (anchor.token !== gestureToken) return
    const container = scrollRef.value
    const page = container?.querySelector<HTMLElement>(`[data-page-index="${anchor.pageIndex}"]`)
    if (!container || !page) return
    const rect = page.getBoundingClientRect()
    container.scrollLeft += rect.left + rect.width * anchor.xRatio - anchor.clientX
    container.scrollTop += rect.top + rect.height * anchor.yRatio - anchor.clientY
  })
}

function requestGestureZoom(nextZoom: number, clientX: number, clientY: number): void {
  const current = controlledZoom()
  if (current === undefined || nextZoom === (pendingZoom ?? current)) return
  const token = ++gestureToken
  const anchor = captureZoomAnchor(clientX, clientY)
  pendingZoom = nextZoom
  pendingAnchor = anchor ? { ...anchor, requestedZoom: nextZoom, token } : undefined
  emit("update:zoom", nextZoom)
}

function onZoomWheel(event: WheelEvent): void {
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
  const rect = scrollRef.value?.getBoundingClientRect()
  requestGestureZoom(
    clampSurfaceZoom(gestureStartZoom * (event.scale ?? 1)),
    event.clientX ?? (rect ? rect.left + rect.width / 2 : 0),
    event.clientY ?? (rect ? rect.top + rect.height / 2 : 0),
  )
}

function onGestureEnd(): void {
  webkitGestureActive = false
}

watch(() => props.zoom, (next) => {
  if (pendingAnchor && typeof next === "number" && clampSurfaceZoom(next) === pendingAnchor.requestedZoom) {
    void restoreZoomAnchor(pendingAnchor)
  } else if (pendingZoom !== undefined) {
    gestureToken += 1
  }
  pendingAnchor = undefined
  pendingZoom = undefined
})

watch(effectiveSource, () => {
  gestureToken += 1
  pendingAnchor = undefined
  pendingZoom = undefined
})

// ── fit-width zoom ───────────────────────────────────────────────────
watch([scrollRef, () => props.fitWidth, renderDocument, rotation], ([el, fit, doc]) => {
  if (!fit || props.zoom !== undefined || !el || !doc) return
  const fitToWidth = (): void => {
    // The widest page is the one that limits the fit, in the orientation the
    // reader currently sees — the intrinsic /Rotate alone is not enough.
    const maxPageWidth = Math.max(
      ...doc.pages.map((page) => pdfDisplaySize(page, rotation.value, 1).width),
      200,
    )
    const containerW = el.clientWidth - 32
    zoom.value = clampSurfaceZoom(containerW / maxPageWidth)
  }
  const observer = new ResizeObserver(fitToWidth)
  observer.observe(el)
  fitToWidth()
  return () => observer.disconnect()
})

onMounted(() => {
  setupPrintDetection()
  const element = scrollRef.value
  if (!element) return
  element.addEventListener("wheel", onZoomWheel, { passive: false })
  element.addEventListener("gesturestart", onGestureStart, { passive: false })
  element.addEventListener("gesturechange", onGestureChange as EventListener, { passive: false })
  element.addEventListener("gestureend", onGestureEnd)
  element.addEventListener("pointerdown", onTouchPointerDown, { passive: true })
  element.addEventListener("pointermove", onTouchPointerMove, { passive: true })
  element.addEventListener("pointerup", onTouchPointerUp)
  element.addEventListener("pointercancel", onTouchPointerCancel)
  measureViewportHeight()
  stopDprWatch = watchDevicePixelRatio((next) => { renderDpr.value = next })
  const ResizeObserverConstructor = globalThis.ResizeObserver
  if (typeof ResizeObserverConstructor === "function") {
    resizeObserver = new ResizeObserverConstructor(() => onSurfaceResize())
    resizeObserver.observe(element)
  }
  if (renderDocument.value) scheduleVisibleRender(0)
})

function onSurfaceResize(): void {
  if (resizeTimer !== undefined) clearTimeout(resizeTimer)
  resizeTimer = setTimeout(() => {
    resizeTimer = undefined
    measureViewportHeight()
    scheduleVisibleRender(0)
  }, PDF_RESIZE_RENDER_DEBOUNCE_MS)
}

// ── Expose ───────────────────────────────────────────────────────────
function scrollToPage(pageIndex: number): void {
  const el = scrollRef.value
  if (!el) return
  const slot = el.querySelector(`[data-page-index="${pageIndex}"]`) as HTMLElement | null
  slot?.scrollIntoView({ behavior: "smooth", block: "start" })
  // Programmatic scrolling may not emit a scroll event (already at the target,
  // reduced motion); rasterize whichever page ends up in the viewport anyway.
  scheduleVisibleRender()
}

function onSurfaceEscape(event?: KeyboardEvent): void {
  if (event?.defaultPrevented) return
  clearSelection()
  if (selectionMode.value !== "content") emit("selectionCancel", { mode: selectionMode.value, reason: "escape" })
}

defineExpose({
  scrollToPage,
  zoom,
  rotation,
  search: (query: string): Promise<PdfSearchState> => surfaceSearch.search(query),
  activateSearchMatch: surfaceSearch.activateSearchMatch,
  searchNext: surfaceSearch.searchNext,
  searchPrevious: surfaceSearch.searchPrevious,
  clearSearch: surfaceSearch.clearSearch,
  getSearchState: (): PdfSearchState => surfaceSearch.getSearchState(),
  clearSelection,
  getSelectionState: () => selectionState.value,
  getPageAnnotations: (pageIndex: number): Promise<readonly PdfRenderAnnotation[]> => pageAnnotationsFor(pageIndex),
  setFormFieldValue,
  updateAnnotationAppearance,
  getDocumentRevision: (): OfficeDocumentRevision => documentRevision.value,
  hitTest,
  async describeReference(reference: OfficeObjectReference, signal?: AbortSignal) {
    const descriptor = await describePdfReference(referenceContext(), reference, signal)
    if (!descriptor) throw new Error("PDF reference is not resolvable in the current revision.")
    return descriptor
  },
  async resolveReference(reference: OfficeObjectReference): Promise<ResolveReferenceResult> {
    const result = await resolvePdfReference(referenceContext(), reference)
    emit("referenceResolve", { referenceId: reference.referenceId, result })
    return result
  },
  async scrollToReference(reference: OfficeObjectReference): Promise<void> {
    const result = await resolvePdfReference(referenceContext(), reference)
    if (result.status !== "exact" && result.status !== "relocated") return
    const locator = (result.reference as PdfOfficeReference).locator
    const pageIndex = locator.type === "manual-region" ? locator.value.pageIndex : locator.value.pageIndex
    if (pageIndex !== undefined) scrollToPage(pageIndex)
  },
  async captureReferencePreview(reference: OfficeObjectReference): Promise<Blob> {
    emitReferenceError("capture", "CAPTURE_UNSUPPORTED", formatOfficeMessage(m.value.common.reference.captureUnsupported, { format: "PDF" }), reference.referenceId)
    throw new Error("PDF reference preview capture is unsupported")
  },
})

onBeforeUnmount(async () => {
  unmounted = true
  stopPrintDetection?.()
  stopPrintDetection = undefined
  latestRequestId += 1
  gestureToken += 1
  stopDprWatch?.()
  stopDprWatch = undefined
  cancelScrollFrame()
  cancelScheduledRender()
  if (resizeTimer !== undefined) {
    clearTimeout(resizeTimer)
    resizeTimer = undefined
  }
  resizeObserver?.disconnect()
  resizeObserver = undefined
  touchGestures.reset()
  const element = scrollRef.value
  element?.removeEventListener("wheel", onZoomWheel)
  element?.removeEventListener("gesturestart", onGestureStart)
  element?.removeEventListener("gesturechange", onGestureChange as EventListener)
  element?.removeEventListener("gestureend", onGestureEnd)
  element?.removeEventListener("pointerdown", onTouchPointerDown)
  element?.removeEventListener("pointermove", onTouchPointerMove)
  element?.removeEventListener("pointerup", onTouchPointerUp)
  element?.removeEventListener("pointercancel", onTouchPointerCancel)
  cancelRender()
  pdfLoadCoordinator.dispose()
  pdfTaskSequence.dispose()
  surfaceSearch.dispose()
  await releaseDocument()
  if (ownedRuntime) await ownedRuntime.dispose()
})
</script>

<style scoped>
.pdf-surface {
  flex: 1;
  overflow-y: auto;
  overflow-x: auto;
  min-height: 0;
  touch-action: pan-x pan-y;
  min-width: 0;
  --pdf-surface-bg: #525659;
  --pdf-surface-page-filter: none;
  background: var(--pdf-surface-bg);
}
.pdf-surface--dark {
  --pdf-surface-bg: #020617;
  --pdf-surface-page-filter: var(--pdf-dark-page-filter, invert(1) hue-rotate(180deg));
}
.pdf-surface__empty {
  display: flex; align-items: center; justify-content: center;
  min-height: 200px; color: #9ca3af; font-size: 14px;
}
.pdf-surface__error {
  padding: 16px; color: #ef4444; font-size: 13px;
}
.pdf-surface__pages {
  display: flex; flex-direction: column; align-items: center;
  /* Vertical gap and top padding are bound inline from PDF_PAGE_GAP_PX /
     PDF_PAGES_PADDING_TOP_PX so the cached page offsets cannot drift. */
  padding: 0 8px 16px;
}
.pdf-surface__page-slot {
  position: relative;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
}
.pdf-surface__page-img { display: block; filter: var(--pdf-surface-page-filter); }
.pdf-surface--dark .pdf-surface__page-slot { box-shadow: 0 4px 16px rgba(0, 0, 0, 0.7); }
/* Real DOM text over the bitmap: transparent, positioned per glyph and
   transformed once per page; pointer events stay on the interaction layer. */
.pdf-surface__text-layer {
  left: 0;
  position: absolute;
  top: 0;
  transform-origin: 0 0;
  pointer-events: none;
  user-select: text;
  z-index: 1;
}
.pdf-surface__text-glyph {
  color: transparent;
  -webkit-text-fill-color: transparent;
  line-height: 1;
  overflow: hidden;
  pointer-events: none;
  position: absolute;
  user-select: text;
  white-space: pre;
}
.pdf-surface__text-glyph--unmapped { user-select: none; }
.pdf-surface__overlay,
.pdf-surface__interaction {
  inset: 0;
  position: absolute;
}
.pdf-surface__overlay { pointer-events: none; z-index: 2; }
.pdf-surface__interaction { cursor: text; touch-action: pan-x pan-y; z-index: 3; }
.pdf-surface__interaction[data-selection-mode="object"] { cursor: pointer; }
.pdf-surface__interaction[data-selection-mode="region"] { cursor: crosshair; touch-action: none; }
.pdf-surface__region-frame {
  background: rgb(37 99 235 / 0.08);
  border: 2px solid #2563eb;
  box-sizing: border-box;
  pointer-events: none;
  position: absolute;
}
.pdf-surface__search-hit,
.pdf-surface__selection-hit { box-sizing: border-box; position: absolute; }
.pdf-surface__search-hit { background: rgba(250, 204, 21, .38); }
.pdf-surface__search-hit--active {
  background: rgba(245, 158, 11, .56);
  outline: 2px solid rgba(180, 83, 9, .9);
}
.pdf-surface__selection-hit { background: rgba(37, 99, 235, .34); }
/* Invisible link/widget hit rectangles; the Surface hit-tests them on click. */
.pdf-surface__annotation-hit { box-sizing: border-box; position: absolute; pointer-events: none; }
/* Native AcroForm controls sit above the interaction layer so typing wins over text selection. */
.pdf-surface__form-layer { inset: 0; position: absolute; z-index: 4; pointer-events: none; }
.pdf-surface__form-control {
  box-sizing: border-box;
  margin: 0;
  min-width: 0;
  pointer-events: auto;
  position: absolute;
}
.pdf-surface__form-control--text {
  background: rgb(255 255 255 / 0.92);
  border: 1px solid rgb(37 99 235 / 0.45);
  border-radius: 2px;
  color: #111827;
  font: inherit;
  padding: 0 2px;
}
.pdf-surface__form-control:disabled { opacity: 0.7; }
.pdf-surface__page-loading,
.pdf-surface__page-error,
.pdf-surface__page-placeholder {
  display: flex; align-items: center; justify-content: center;
  background: #fff; min-height: 200px; min-width: 200px;
  font-size: 12px; color: #6b7280;
  box-sizing: border-box; width: 100%; height: 100%;
}
.pdf-surface__page-error { color: #ef4444; }

@media print {
  .pdf-surface {
    overflow: visible !important;
    background: #fff !important;
  }
  .pdf-surface__interaction,
  .pdf-surface__overlay,
  .pdf-surface__empty,
  .pdf-surface__error,
  .pdf-surface__page-loading,
  .pdf-surface__page-error,
  .pdf-surface__page-placeholder { display: none !important; }
  .pdf-surface__pages { display: block !important; padding: 0 !important; }
  .pdf-surface__page-slot {
    break-after: page;
    box-shadow: none !important;
    height: auto !important;
    margin: 0 auto;
    max-width: 100%;
    width: 100% !important;
  }
  .pdf-surface__page-img {
    filter: none !important;
    height: auto !important;
    max-width: 100% !important;
    width: 100% !important;
  }
}
</style>
