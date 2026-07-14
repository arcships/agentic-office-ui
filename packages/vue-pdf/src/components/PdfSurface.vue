<template>
  <div
    ref="scrollRef"
    class="pdf-surface"
    data-testid="pdf-surface"
    tabindex="0"
    @scroll.passive="onScroll"
    @contextmenu.prevent="onContextMenu"
    @copy="onCopy"
    @keydown.esc="onSurfaceEscape"
  >
    <div v-if="!effectiveSource" class="pdf-surface__empty">No PDF loaded</div>
    <div v-else-if="loading" class="pdf-surface__empty">Opening PDF…</div>
    <div v-else-if="loadError" class="pdf-surface__error" :data-error-code="loadError.code">
      {{ loadError.code }}: {{ loadError.message }}
    </div>
    <div v-else-if="renderDocument" class="pdf-surface__pages">
      <div
        v-for="pageIndex in Array.from({ length: renderDocument.pageCount }, (_, i) => i)"
        :key="`page-${pageIndex}`"
        class="pdf-surface__page-slot"
        :data-page-index="pageIndex"
        :style="pageSlotStyle(pageIndex)"
      >
        <div v-if="pageSlots[pageIndex]?.state === 'loading'" class="pdf-surface__page-loading">
          Rendering page {{ pageIndex + 1 }}…
        </div>
        <div v-else-if="pageSlots[pageIndex]?.state === 'error'" class="pdf-surface__page-error">
          Failed to render page {{ pageIndex + 1 }}
        </div>
        <img
          v-else-if="pageSlots[pageIndex]?.url"
          :src="pageSlots[pageIndex].url"
          :alt="`Page ${pageIndex + 1}`"
          :style="pageImgStyle(pageIndex)"
          class="pdf-surface__page-img"
          draggable="false"
        />
        <div v-else class="pdf-surface__page-placeholder" />
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
            v-for="(rect, rectIndex) in selectionRectsForPage(pageIndex)"
            :key="`selection-${rectIndex}`"
            class="pdf-surface__selection-hit"
            :style="displayRectStyle(pageIndex, rect)"
          />
        </div>
        <div
          class="pdf-surface__interaction"
          :data-selection-mode="selectionMode"
          :aria-label="selectionMode === 'region' ? '选择 PDF 页面区域' : 'PDF page text selection layer'"
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
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, reactive, ref, watch } from "vue"
import type { PageTextSlice, PdfPageGeometry, Position } from "@embedpdf/models"
import { expandToWordBoundary, glyphAt, rectsWithinSlice } from "@embedpdf/plugin-selection"
import { clampSurfaceZoom, nextSurfaceZoom } from "@arcships/office-runtime/gesture-zoom"
import {
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
  createSurfaceSearchSession,
  type OfficeSource,
} from "@arcships/office-runtime"
import {
  DEFAULT_PDF_MAX_FILE_SIZE,
  loadVerifiedPdfSource,
  toPdfLoadError,
  type PdfDiagnostic,
  type PdfLoadError,
  type PdfSource,
  type PdfUrlPolicy,
} from "../pdf-url-policy"
import {
  createPdfRenderRuntime,
  type PdfRenderDocument,
  type PdfRenderRect,
  type PdfRenderRuntime,
  type PdfRotation,
  type PdfSearchHit,
  type PdfSearchState,
  type PdfSelectionState,
} from "../pdf/pdf-render-runtime"
import {
  pdfCanonicalRectToDisplay,
  pdfDisplayPointToCanonical,
  pdfDisplaySize,
} from "../pdf/pdf-coordinate-transform"
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
  type OfficeSelectionMode,
  type ResolveReferenceResult,
} from "@arcships/office-interaction"

// ── Types ────────────────────────────────────────────────────────────
interface PageSlot {
  url: string | null
  state: "pending" | "loading" | "ready" | "error"
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
  }>(),
  {
    maxFileSize: DEFAULT_PDF_MAX_FILE_SIZE,
    defaultZoom: 1,
    enableGestureZoom: true,
    selectionMode: "content",
    emitReferenceCandidates: false,
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
  regionDraftChange: [event: { phase: "start" | "change"; region: { space: "page"; pageIndex: number; rect: NormalizedRect } }]
  selectionCancel: [event: { mode: OfficeSelectionMode; reason: "escape" | "pointer-cancel" | "programmatic" }]
  referenceResolve: [event: { referenceId: string; result: ResolveReferenceResult }]
  referenceError: [error: OfficeReferenceError]
}>()

// ── Runtime & loading ────────────────────────────────────────────────
const scrollRef = ref<HTMLElement | null>(null)
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
const loading = ref(false)
const loadError = ref<PdfLoadError | null>(null)
const renderDocument = ref<PdfRenderDocument | null>(null)
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
let unmounted = false

const searchState = ref<PdfSearchState>({
  status: "idle",
  query: "",
  matches: [],
  activeIndex: -1,
})
const selectionState = ref<PdfSelectionState>({ kind: "none" })
const selectionVisual = ref<{ pageIndex: number; rects: readonly PdfRenderRect[] } | null>(null)
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
  geometry?: PdfPageGeometry
  anchor?: number
  focus?: number
} | null = null
let selectionFrame: number | undefined
let selectionClickCandidate = false

function referenceContext(): PdfReferenceContext {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) throw new Error("PDF reference context is unavailable before the document loads.")
  return { revision: documentRevision.value, document, runtime }
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
    emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : "Unable to confirm PDF reference.")
  }
}

function pageCandidate(pageIndex: number): OfficeReferenceCandidatePreview {
  const draft = createPdfPageReferenceDraft(referenceContext(), pageIndex)
  return {
    candidateId: `pdf:page:${pageIndex}`,
    draft,
    preview: { label: `Page ${pageIndex + 1}`, path: [{ kind: "document", label: "PDF" }, { kind: "page", label: `Page ${pageIndex + 1}` }] },
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
    emitReferenceError("hit-test", "HIT_TEST_FAILED", reason instanceof Error ? reason.message : "PDF hit test failed.")
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
  regionGesture.value = { pointerId: event.pointerId, pageIndex, layer, startX: event.clientX, startY: event.clientY }
  regionPageIndex.value = pageIndex
  regionRect.value = { x: 0, y: 0, width: 0, height: 0 }
  layer.setPointerCapture?.(event.pointerId)
  emit("regionDraftChange", { phase: "start", region: { space: "page", pageIndex, rect: regionRect.value } })
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
  confirmDraft(createPdfRegionReferenceDraft(referenceContext(), { space: "page", pageIndex: gesture.pageIndex, rect }), event.pointerType === "touch" ? "touch" : "pointer")
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
  else if (selectionMode.value === "object") confirmDraft(pageCandidate(pageIndex).draft, event.detail === 0 ? "keyboard" : "pointer", pageCandidate(pageIndex).preview, event.shiftKey)
}
function onRegionKeyboardCancel(): void {
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

// ── Page slots ───────────────────────────────────────────────────────
const pageSlots = reactive<Record<number, PageSlot>>({})

function initPageSlots(): void {
  for (const key of Object.keys(pageSlots)) delete pageSlots[Number(key)]
  if (!renderDocument.value) return
  for (let i = 0; i < renderDocument.value.pageCount; i++) {
    pageSlots[i] = { url: null, state: "pending" }
  }
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
    ownedRuntime = createPdfRenderRuntime({ wasmUrl })
    ownedRuntimeWasmUrl = wasmUrl
  }
  return ownedRuntime
}

function cancelRender(): void {
  renderGeneration++
  renderAbort?.abort()
  renderAbort = undefined
}

async function releaseDocument(): Promise<void> {
  cancelRender()
  surfaceSearch.clearSearch()
  clearSelection()
  geometryCache.clear()
  geometryRequests.clear()
  characterTextCache.clear()
  for (const slot of Object.values(pageSlots)) {
    if (slot.url) URL.revokeObjectURL(slot.url)
  }
  for (const key of Object.keys(pageSlots)) delete pageSlots[Number(key)]
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

// ── Render all pages sequentially ────────────────────────────────────
async function renderAllPages(requestId: number, parentSignal?: AbortSignal): Promise<void> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt) return
  const gen = ++renderGeneration
  renderAbort?.abort()
  const { controller, unlink } = createLinkedAbort(parentSignal)
  renderAbort = controller

  for (let i = 0; i < doc.pageCount; i++) {
    if (controller.signal.aborted || gen !== renderGeneration || requestId !== latestRequestId) return
    const slot = pageSlots[i]
    if (!slot || slot.state === "ready") continue
    slot.state = "loading"
    try {
      const blob = await rt.renderPage(doc, i, {
        scale: zoom.value,
        rotation: rotation.value,
        dpr: typeof devicePixelRatio === "number" ? Math.min(2, Math.max(1, devicePixelRatio)) : 1,
        signal: controller.signal,
      })
      if (controller.signal.aborted || gen !== renderGeneration || requestId !== latestRequestId) return
      const url = URL.createObjectURL(blob)
      if (slot.url) URL.revokeObjectURL(slot.url)
      slot.url = url
      slot.state = "ready"
    } catch (err) {
      if ((err as Error)?.name === "AbortError" || controller.signal.aborted) return
      slot.state = "error"
    }
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
  if (requestId !== latestRequestId || unmounted) return
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
    task.assertCurrent()
    selectedRuntime = await selectRuntime()
    task.assertCurrent()
    const content = verified.bytes.buffer.slice(
      verified.bytes.byteOffset,
      verified.bytes.byteOffset + verified.bytes.byteLength,
    ) as ArrayBuffer
    openedDocument = await selectedRuntime.openDocument(content, { signal: task.signal })
    task.assertCurrent()
    if (openedDocument.pageCount <= 0) throw new Error("PDF document contains no pages.")
    runtimeForDocument = selectedRuntime
    renderDocument.value = openedDocument
    revisionCounter.value += 1
    emit("documentRevisionChange", documentRevision.value)
    verifiedBlob.value = verified.blob
    initPageSlots()
    emit("document-load-success", openedDocument.pageCount)
    loading.value = false
    void renderAllPages(requestId, task.signal)
    prefetchPageGeometry(0)
  } catch (cause) {
    if (openedDocument && selectedRuntime) {
      await selectedRuntime.closeDocument(openedDocument)
    }
    if (!task.isCurrent() || requestId !== latestRequestId || (cause as Error)?.name === "AbortError") return
    const sourceKind = source.kind
    const normalized = toPdfLoadError(cause, sourceKind, sourceKind === "url" ? source.url : undefined)
    loadError.value = normalized
    loading.value = false
    emit("document-load-error", normalized)
  }
}

// ── Watchers ─────────────────────────────────────────────────────────
watch([zoom, rotation], () => {
  if (!renderDocument.value) return
  cancelRender()
  initPageSlots()
  void renderAllPages(latestRequestId)
})

watch(effectiveSource, () => {
  void load()
}, { immediate: true })

watch(() => props.documentId, () => emit("documentRevisionChange", documentRevision.value))

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

function selectionRectsForPage(pageIndex: number): readonly PdfRenderRect[] {
  return selectionVisual.value?.pageIndex === pageIndex ? selectionVisual.value.rects : []
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

function flushSelectionGesture(gesture: NonNullable<typeof selectionGesture>): void {
  const geometry = gesture.geometry
  if (selectionGesture !== gesture || !gesture.thresholdExceeded || !geometry) return
  beginGestureSelection(gesture)

  if (gesture.anchor === undefined) {
    const down = canonicalPoint(gesture.down, gesture.pageIndex, gesture.layer)
    gesture.anchor = down ? resolveSelectionGlyph(geometry, down) : -1
  }
  if (gesture.anchor < 0) return

  const latest = canonicalPoint(gesture.latest, gesture.pageIndex, gesture.layer)
  if (!latest) return
  const resolvedFocus = resolveSelectionGlyph(geometry, latest, gesture.focus ?? gesture.anchor)
  const focus = resolvedFocus >= 0 ? resolvedFocus : (gesture.focus ?? gesture.anchor)
  if (focus === gesture.focus && selectionVisual.value?.pageIndex === gesture.pageIndex) return

  gesture.focus = focus
  selectionVisual.value = {
    pageIndex: gesture.pageIndex,
    rects: selectionRects(geometry, gesture.anchor, focus),
  }
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

async function commitSelectionRange(
  pageIndex: number,
  anchor: number,
  focus: number,
  rects: readonly PdfRenderRect[],
  request: number,
): Promise<void> {
  const doc = renderDocument.value
  const rt = runtimeForDocument
  if (!doc || !rt) return
  const from = Math.min(anchor, focus)
  const range: PageTextSlice = {
    pageIndex,
    charIndex: from,
    charCount: Math.max(anchor, focus) - from + 1,
  }
  try {
    const [text] = await rt.getTextSlices(doc, [range])
    if (request !== selectionRequest || renderDocument.value !== doc || runtimeForDocument !== rt) return
    if (!text) {
      setSelectionUnavailable(pageIndex)
      return
    }
    if (/[\u0000\ufffd\ufffe]/u.test(text)) {
      const unavailable: PdfSelectionState = { kind: "unavailable", pageIndex, reason: "unicode-map" }
      selectionState.value = unavailable
      selectionVisual.value = null
      emit("selectionChange", unavailable)
      return
    }
    const next: PdfSelectionState = { kind: "text", range, text, rects }
    selectionState.value = next
    emit("selectionChange", next)
    if (selectionMode.value === "content") {
      try {
        const draft = createPdfTextReferenceDraft(referenceContext(), next)
        confirmDraft(draft, "pointer", {
          label: text.length > 80 ? `${text.slice(0, 77)}…` : text,
          path: [{ kind: "document", label: "PDF" }, { kind: "page", label: `Page ${pageIndex + 1}` }, { kind: "text-range", label: "Selected text" }],
          content: { text },
        })
      } catch (reason) {
        emitReferenceError("describe", "INVALID_REFERENCE", reason instanceof Error ? reason.message : "Unable to create PDF text reference.")
      }
    }
  } catch {
    if (request === selectionRequest) setSelectionUnavailable(pageIndex)
  }
}

function finalizeSelectionGesture(gesture: NonNullable<typeof selectionGesture>): void {
  if (selectionGesture !== gesture) return
  cancelSelectionFrame()
  flushSelectionGesture(gesture)
  selectionGesture = null
  if ((gesture.anchor ?? -1) < 0 || (gesture.focus ?? -1) < 0) {
    selectionVisual.value = null
    return
  }
  const rects = selectionVisual.value?.pageIndex === gesture.pageIndex
    ? selectionVisual.value.rects
    : []
  void commitSelectionRange(
    gesture.pageIndex,
    gesture.anchor!,
    gesture.focus!,
    rects,
    gesture.request,
  )
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
  }
  selectionGesture = gesture
  layer.setPointerCapture(event.pointerId)
  void pageGeometry(pageIndex).then((geometry) => {
    if (selectionGesture !== gesture || request !== selectionRequest) return
    gesture.geometry = geometry
    if (!gesture.thresholdExceeded) return
    flushSelectionGesture(gesture)
    if (!gesture.pointerActive) finalizeSelectionGesture(gesture)
  }).catch(() => {
    if (selectionGesture !== gesture || request !== selectionRequest) return
    if (layer.hasPointerCapture(event.pointerId)) layer.releasePointerCapture(event.pointerId)
    clearSelection()
  })
}

function onSelectionPointerMove(event: PointerEvent, pageIndex: number): void {
  const gesture = selectionGesture
  if (!gesture || gesture.pointerId !== event.pointerId || gesture.pageIndex !== pageIndex) return
  updateGesturePoint(gesture, eventClientPoint(event))
  if (gesture.thresholdExceeded && gesture.geometry) scheduleSelectionFrame(gesture)
}

function onSelectionPointerUp(event: PointerEvent, pageIndex: number): void {
  const gesture = selectionGesture
  const layer = event.currentTarget as HTMLElement
  if (!gesture || gesture.pointerId !== event.pointerId || gesture.pageIndex !== pageIndex) return
  updateGesturePoint(gesture, eventClientPoint(event))
  gesture.pointerActive = false
  if (layer.hasPointerCapture(event.pointerId)) layer.releasePointerCapture(event.pointerId)
  if (!gesture.thresholdExceeded) {
    stopSelectionGesture()
    clearSelection()
    selectionClickCandidate = true
    return
  }
  selectionClickCandidate = false
  if (gesture.geometry) finalizeSelectionGesture(gesture)
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

async function onSelectionClick(event: MouseEvent, pageIndex: number): Promise<void> {
  const isNativeDoubleClick = selectionClickCandidate && event.button === 0 && event.detail === 2
  selectionClickCandidate = false
  if (!isNativeDoubleClick) return
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
    const rects = selectionRects(geometry, range.from, range.to)
    selectionVisual.value = { pageIndex, rects }
    await commitSelectionRange(pageIndex, range.from, range.to, rects, request)
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
  const el = scrollRef.value
  if (!el || !renderDocument.value) return
  const center = el.scrollTop + el.clientHeight / 2
  let current = 0
  for (let i = 0; i < renderDocument.value.pageCount; i++) {
    const slot = el.querySelector(`[data-page-index="${i}"]`) as HTMLElement | null
    if (!slot) continue
    if (slot.offsetTop <= center) current = i
  }
  emit("visible-page-change", current)
  prefetchPageGeometry(current)
}

function onContextMenu(event: MouseEvent): void {
  const el = scrollRef.value
  if (!el || !renderDocument.value) return
  const rect = el.getBoundingClientRect()
  const clickY = event.clientY - rect.top + el.scrollTop
  let pageIndex = 0
  for (let i = 0; i < renderDocument.value.pageCount; i++) {
    const slot = el.querySelector(`[data-page-index="${i}"]`) as HTMLElement | null
    if (!slot) continue
    if (clickY >= slot.offsetTop && clickY < slot.offsetTop + slot.offsetHeight) {
      pageIndex = i; break
    }
  }
  emit("contextMenu", {
    pageIndex,
    clientX: event.clientX,
    clientY: event.clientY,
    containerX: event.clientX - rect.left,
    containerY: event.clientY - rect.top,
  })
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

function captureZoomAnchor(clientX: number, clientY: number): PageAnchor | undefined {
  if (typeof document === "undefined") return undefined
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
watch([scrollRef, () => props.fitWidth, renderDocument], ([el, fit, doc]) => {
  if (!fit || props.zoom !== undefined || !el || !doc) return
  const observer = new ResizeObserver(() => {
    const maxPageWidth = Math.max(
      ...doc.pages.map((p) => (p.rotation / 90) % 2 ? p.height : p.width),
      200,
    )
    const containerW = el.clientWidth - 32
    zoom.value = clampSurfaceZoom(containerW / maxPageWidth)
  })
  observer.observe(el)
  return () => observer.disconnect()
})

onMounted(() => {
  const element = scrollRef.value
  if (!element) return
  element.addEventListener("wheel", onZoomWheel, { passive: false })
  element.addEventListener("gesturestart", onGestureStart, { passive: false })
  element.addEventListener("gesturechange", onGestureChange as EventListener, { passive: false })
  element.addEventListener("gestureend", onGestureEnd)
})

// ── Expose ───────────────────────────────────────────────────────────
function scrollToPage(pageIndex: number): void {
  const el = scrollRef.value
  if (!el) return
  const slot = el.querySelector(`[data-page-index="${pageIndex}"]`) as HTMLElement | null
  slot?.scrollIntoView({ behavior: "smooth", block: "start" })
}

function onSurfaceEscape(): void {
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
    emitReferenceError("capture", "CAPTURE_UNSUPPORTED", "PDF preview capture is not provided by the Surface; use the host's capture pipeline.", reference.referenceId)
    throw new Error("PDF reference preview capture is unsupported")
  },
})

onBeforeUnmount(async () => {
  unmounted = true
  gestureToken += 1
  const element = scrollRef.value
  element?.removeEventListener("wheel", onZoomWheel)
  element?.removeEventListener("gesturestart", onGestureStart)
  element?.removeEventListener("gesturechange", onGestureChange as EventListener)
  element?.removeEventListener("gestureend", onGestureEnd)
  cancelRender()
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
  min-width: 0;
  background: var(--pdf-surface-bg, #525659);
}
.pdf-surface__empty {
  display: flex; align-items: center; justify-content: center;
  height: 200px; color: #9ca3af; font-size: 14px;
}
.pdf-surface__error {
  padding: 16px; color: #ef4444; font-size: 13px;
}
.pdf-surface__pages {
  display: flex; flex-direction: column; align-items: center;
  gap: 8px; padding: 16px 8px;
}
.pdf-surface__page-slot {
  position: relative;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
}
.pdf-surface__page-img { display: block; }
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
.pdf-surface__page-loading,
.pdf-surface__page-error,
.pdf-surface__page-placeholder {
  display: flex; align-items: center; justify-content: center;
  background: #fff; min-height: 200px; min-width: 200px;
  font-size: 12px; color: #6b7280;
  box-sizing: border-box; width: 100%; height: 100%;
}
.pdf-surface__page-error { color: #ef4444; }
</style>
