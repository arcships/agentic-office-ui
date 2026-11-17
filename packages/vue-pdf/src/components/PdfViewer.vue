<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue"
import type { PdfPageGeometry, Position } from "@embedpdf/models"
import { expandToWordBoundary, glyphAt, rectsWithinSlice } from "@embedpdf/plugin-selection"
import { createTouchGestureTracker, nextDoubleTapZoom } from "@arcships/office-runtime/touch-gestures"
import {
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
  sanitizeOfficeUrl,
  type OfficeSource,
  type OfficeUrlPolicy,
} from "@arcships/office-runtime"
import {
  DEFAULT_PDF_MAX_FILE_SIZE,
  loadVerifiedPdfSource,
  localizePdfErrorMessage,
  toPdfLoadError,
  type PdfDiagnostic,
  type PdfLoadError,
  type PdfSource,
} from "../pdf-url-policy"
import {
  createPdfRenderRuntime,
  type PdfRenderDocument,
  type PdfRenderRect,
  type PdfRenderRuntime,
  type PdfRotation,
  type PdfSearchHit,
  type PdfSelectionState,
  type PdfPageTextLayer,
  type PdfPageTextLayerGlyph,
} from "../pdf/pdf-render-runtime"
import {
  pdfCanonicalRectToDisplay,
  pdfCanonicalToDisplayTransform,
  pdfDisplayPointToCanonical,
} from "../pdf/pdf-coordinate-transform"
import {
  nearestGlyphWithinContainingRun,
  wordRangeFromCharacterTexts,
} from "../pdf/pdf-selection-interaction"
import { hasUnmappedPdfText } from "../pdf/pdf-selection-range"
import { currentRenderDpr, watchDevicePixelRatio } from "../pdf/pdf-device-pixel-ratio"
import type { PdfViewerProps } from "../types"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { providePdfMessages } from "../i18n"

const props = withDefaults(defineProps<PdfViewerProps>(), {
  src: undefined,
  source: undefined,
  urlPolicy: undefined,
  runtime: undefined,
  pdfiumWasmUrl: undefined,
  maxFileSize: DEFAULT_PDF_MAX_FILE_SIZE,
  password: undefined,
  onPassword: undefined,
  fileName: undefined,
  defaultZoom: 1,
  className: "",
  showToolbar: true,
  showDownload: true,
  showRotateControls: true,
  minHeight: undefined,
  isDark: false,
  locale: "en",
  messages: undefined,
})

const emit = defineEmits<{
  "document-load-success": [numPages: number]
  "document-load-error": [error: PdfLoadError]
  "active-page-change": [pageNumber: number]
  diagnostic: [event: PdfDiagnostic]
  selectionChange: [selection: PdfSelectionState]
}>()

const m = providePdfMessages(() => props.locale, () => props.messages)

const zoomOptions = [0.5, 0.75, 1, 1.25, 1.5, 2]
const zoom = ref(Math.min(2, Math.max(0.5, props.defaultZoom)))
const rotation = ref<PdfRotation>(0)
/** Pixel ratio the current page bitmap was rendered at; re-renders on display change. */
const renderDpr = ref(currentRenderDpr())
const activePage = ref(1)
const pageInput = ref("1")
const showThumbnails = ref(true)
const searchDraft = ref("")
const viewerRef = ref<HTMLElement | null>(null)
const frameRef = ref<HTMLElement | null>(null)
const pageImageRef = ref<HTMLElement | null>(null)
const searchInputRef = ref<HTMLInputElement | null>(null)
const searchResult = ref("")
const searchHits = ref<readonly PdfSearchHit[]>([])
const searchHitIndex = ref(-1)
const loadError = ref<PdfLoadError | null>(null)
const loadErrorMessage = computed(() => loadError.value
  ? localizePdfErrorMessage(loadError.value, m.value.pdf.errors)
  : "")
const pageRenderError = ref("")
const downloadError = ref("")
const isLoading = ref(false)
const numPages = ref(0)
const verifiedBlob = ref<Blob | null>(null)
const loadedFileName = ref("")
const renderDocument = ref<PdfRenderDocument | null>(null)
const pageImageUrl = ref("")
const pageRenderState = ref<"idle" | "rendering" | "rendered" | "error">("idle")
const renderedPageNumber = ref(0)
const renderedZoom = ref(1)
const renderedRotation = ref<PdfRotation>(0)
const thumbnailUrls = ref<Record<number, string>>({})
const printPageUrls = ref<Record<number, string>>({})
const printMode = ref(false)
const selectionState = ref<PdfSelectionState>({ kind: "none" })
const pageTextLayer = ref<PdfPageTextLayer | null>(null)
let textLayerRequest = 0
const selectionRects = ref<readonly PdfRenderRect[]>([])

/** Drag threshold before a single-page selection starts, in CSS px. */
const PDF_VIEWER_SELECTION_DRAG_THRESHOLD_PX = 4
/** Print pages are rasterized above 1x so the print snapshot stays sharp. */
const PRINT_RENDER_SCALE = 1.5
const selectionGeometries = new Map<number, PdfPageGeometry>()
const selectionGeometryRequests = new Map<number, Promise<PdfPageGeometry>>()
const characterTextCache = new Map<number, Map<number, string>>()
let selectionRequest = 0
interface ViewerSelectionGesture {
  request: number
  pageIndex: number
  pointerId: number
  layer: HTMLElement
  down: { x: number; y: number }
  latest: { x: number; y: number }
  pointerActive: boolean
  thresholdExceeded: boolean
  geometry?: PdfPageGeometry
  anchor?: number
  focus?: number
}
let selectionGesture: ViewerSelectionGesture | null = null
let selectionClickCandidate = false

let latestRequestId = 0
let latestTaskId: string | undefined
let renderGeneration = 0
let searchGeneration = 0
let thumbnailGeneration = 0
let renderAbort: AbortController | undefined
let searchAbort: AbortController | undefined
let thumbnailAbort: AbortController | undefined
let printAbort: AbortController | undefined
let printGeneration = 0
let printMediaQuery: MediaQueryList | undefined
let printPageStyle: HTMLStyleElement | undefined
let stopPrintDetection: (() => void) | undefined
let runtimeForDocument: PdfRenderRuntime | undefined
let ownedRuntime: PdfRenderRuntime | undefined
let ownedRuntimeWasmUrl: string | undefined
let unmounted = false
let stopDprWatch: (() => void) | undefined

const pdfRuntimeId = `pdf-controller-${
  typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}`
const pdfTaskSequence = createOfficeTaskSequence(pdfRuntimeId)
const pdfLoadCoordinator = createLatestTaskCoordinator(pdfTaskSequence)

const effectiveSource = computed<PdfSource | undefined>(() => {
  if (props.source) return props.source
  if (props.src) return { kind: "url", url: props.src }
  return undefined
})
const viewerState = computed(() => {
  if (!effectiveSource.value) return "idle"
  if (isLoading.value || pageRenderState.value === "rendering") return "loading"
  if (loadError.value || pageRenderState.value === "error") return "error"
  return renderDocument.value && pageImageUrl.value ? "ready" : "idle"
})
const controlsDisabled = computed(() => !renderDocument.value || !!loadError.value)
const pageLabels = computed(() => Array.from({ length: numPages.value }, (_, index) => index + 1))
const currentSearchHit = computed(() => searchHits.value[searchHitIndex.value])
const pageImageStyle = computed(() => {
  const document = renderDocument.value
  const page = document?.pages[renderedPageNumber.value - 1]
  if (!page) return undefined
  const quarterTurns = ((page.rotation + renderedRotation.value) / 90) % 4
  const width = (quarterTurns % 2 ? page.height : page.width) * renderedZoom.value
  return { width: `${Math.max(1, Math.round(width))}px` }
})
const printPageEntries = computed(() =>
  Object.entries(printPageUrls.value)
    .map(([page, url]) => ({ page: Number(page), url }))
    .filter((entry) => Number.isFinite(entry.page) && Boolean(entry.url))
    .sort((left, right) => left.page - right.page)
)
const viewerMinHeightStyle = computed(() => {
  const value = props.minHeight
  if (value === undefined) return undefined
  return { minHeight: typeof value === "number" ? `${value}px` : value }
})

function cn(...classes: (string | undefined | null | false)[]): string {
  return classes.filter(Boolean).join(" ")
}

function isAbortError(error: unknown): boolean {
  return error instanceof Error && (error.name === "AbortError" || /aborted|cancelled/i.test(error.message))
}

// ── Single-page text selection ───────────────────────────────────────

const renderedPage = computed(() => renderDocument.value?.pages[renderedPageNumber.value - 1])

function selectionRectStyle(rect: PdfRenderRect): Record<string, string> {
  const page = renderedPage.value
  const displayed = page
    ? pdfCanonicalRectToDisplay(page, rect, renderedRotation.value, renderedZoom.value)
    : rect
  return {
    left: `${displayed.x}px`,
    top: `${displayed.y}px`,
    width: `${displayed.width}px`,
    height: `${displayed.height}px`,
  }
}

// ── Accessible text layer ────────────────────────────────────────────

async function loadPageTextLayer(pageNumber: number): Promise<void> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  const request = ++textLayerRequest
  if (!document || !runtime || typeof runtime.getPageTextLayer !== "function") {
    pageTextLayer.value = null
    return
  }
  try {
    const layer = await runtime.getPageTextLayer(document, pageNumber - 1)
    if (request !== textLayerRequest || renderDocument.value !== document || runtimeForDocument !== runtime) return
    pageTextLayer.value = layer
  } catch {
    if (request === textLayerRequest) pageTextLayer.value = null
  }
}

const textLayerTransform = computed(() => {
  const page = renderedPage.value
  if (!page) return ""
  return pdfCanonicalToDisplayTransform(page, renderedRotation.value, renderedZoom.value)
})

const textLayerStyle = computed<Record<string, string> | undefined>(() => {
  const layer = pageTextLayer.value
  if (!layer) return undefined
  return {
    width: `${layer.width}px`,
    height: `${layer.height}px`,
    transform: textLayerTransform.value,
  }
})

function textGlyphStyle(glyph: PdfPageTextLayerGlyph): Record<string, string> {
  return {
    left: `${glyph.rect.x}px`,
    top: `${glyph.rect.y}px`,
    width: `${glyph.rect.width}px`,
    height: `${glyph.rect.height}px`,
    fontSize: `${glyph.rect.height}px`,
  }
}

function emitSelection(next: PdfSelectionState): void {
  selectionState.value = next
  emit("selectionChange", next)
}

function clearSelection(): void {
  selectionRequest += 1
  selectionClickCandidate = false
  selectionGesture = null
  selectionRects.value = []
  if (selectionState.value.kind === "none") return
  emitSelection({ kind: "none" })
}

function selectionPageIndex(): number {
  return Math.max(0, renderedPageNumber.value - 1)
}

function pageGeometry(pageIndex: number): Promise<PdfPageGeometry> {
  const cached = selectionGeometries.get(pageIndex)
  if (cached) return Promise.resolve(cached)
  const pending = selectionGeometryRequests.get(pageIndex)
  if (pending) return pending
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) return Promise.reject(new Error("PDF document is not ready."))
  let request: Promise<PdfPageGeometry>
  request = runtime.getPageTextGeometry(document, pageIndex)
    .then((geometry) => {
      if (renderDocument.value === document && runtimeForDocument === runtime) {
        selectionGeometries.set(pageIndex, geometry)
      }
      return geometry
    })
    .finally(() => {
      if (selectionGeometryRequests.get(pageIndex) === request) selectionGeometryRequests.delete(pageIndex)
    })
  selectionGeometryRequests.set(pageIndex, request)
  return request
}

function selectionPoint(event: { clientX: number; clientY: number }, layer: HTMLElement): Position | null {
  const page = renderedPage.value
  if (!page) return null
  const bounds = layer.getBoundingClientRect()
  return pdfDisplayPointToCanonical(
    page,
    { x: event.clientX - bounds.left, y: event.clientY - bounds.top },
    renderedRotation.value,
    renderedZoom.value,
  )
}

function resolveSelectionChar(geometry: PdfPageGeometry, point: Position, previous = -1): number {
  const direct = glyphAt(geometry, point)
  return direct >= 0 ? direct : nearestGlyphWithinContainingRun(geometry, point, previous)
}

function updateSelectionRects(geometry: PdfPageGeometry, anchor: number, focus: number): void {
  selectionRects.value = rectsWithinSlice(geometry, Math.min(anchor, focus), Math.max(anchor, focus))
    .map((rect) => ({
      x: rect.origin.x,
      y: rect.origin.y,
      width: rect.size.width,
      height: rect.size.height,
    }))
}

function beginViewerSelection(gesture: ViewerSelectionGesture): void {
  if (gesture.thresholdExceeded) return
  gesture.thresholdExceeded = true
  selectionRects.value = []
  if (selectionState.value.kind !== "none") emitSelection({ kind: "none" })
}

function updateViewerGesture(gesture: ViewerSelectionGesture, point: { x: number; y: number }): void {
  gesture.latest = point
  if (gesture.thresholdExceeded) return
  if (Math.hypot(point.x - gesture.down.x, point.y - gesture.down.y) < PDF_VIEWER_SELECTION_DRAG_THRESHOLD_PX) return
  beginViewerSelection(gesture)
}

function flushViewerSelection(gesture: ViewerSelectionGesture): void {
  if (selectionGesture !== gesture || !gesture.thresholdExceeded || !gesture.geometry) return
  if (gesture.anchor === undefined) {
    const downPoint = selectionPoint({ clientX: gesture.down.x, clientY: gesture.down.y }, gesture.layer)
    gesture.anchor = downPoint ? resolveSelectionChar(gesture.geometry, downPoint) : -1
  }
  if (gesture.anchor < 0) return
  const latestPoint = selectionPoint({ clientX: gesture.latest.x, clientY: gesture.latest.y }, gesture.layer)
  if (!latestPoint) return
  const resolved = resolveSelectionChar(gesture.geometry, latestPoint, gesture.focus ?? gesture.anchor)
  const focus = resolved >= 0 ? resolved : (gesture.focus ?? gesture.anchor)
  if (focus === gesture.focus && selectionRects.value.length > 0) return
  gesture.focus = focus
  updateSelectionRects(gesture.geometry, gesture.anchor, focus)
}

async function commitViewerSelection(gesture: ViewerSelectionGesture): Promise<void> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) return
  const from = Math.min(gesture.anchor ?? -1, gesture.focus ?? -1)
  const to = Math.max(gesture.anchor ?? -1, gesture.focus ?? -1)
  const request = gesture.request
  if (from < 0 || to < from) {
    clearSelection()
    return
  }
  const range = { pageIndex: gesture.pageIndex, charIndex: from, charCount: to - from + 1 }
  try {
    const [text] = await runtime.getTextSlices(document, [range])
    if (request !== selectionRequest || renderDocument.value !== document || runtimeForDocument !== runtime) return
    if (!text || hasUnmappedPdfText(text)) {
      emitSelection({ kind: "unavailable", pageIndex: gesture.pageIndex, reason: text ? "unicode-map" : "empty" })
      selectionRects.value = []
      return
    }
    emitSelection({ kind: "text", range, text, rects: [...selectionRects.value] })
  } catch {
    if (request === selectionRequest) emitSelection({ kind: "unavailable", pageIndex: gesture.pageIndex, reason: "empty" })
  }
}

function onSelectionPointerDown(event: PointerEvent): void {
  if (!event.isPrimary || event.button !== 0) return
  const layer = event.currentTarget as HTMLElement
  viewerRef.value?.focus()
  selectionGesture = null
  selectionClickCandidate = false
  const pageIndex = selectionPageIndex()
  const request = ++selectionRequest
  const gesture: ViewerSelectionGesture = {
    request,
    pageIndex,
    pointerId: event.pointerId,
    layer,
    down: { x: event.clientX, y: event.clientY },
    latest: { x: event.clientX, y: event.clientY },
    pointerActive: true,
    thresholdExceeded: false,
  }
  selectionGesture = gesture
  layer.setPointerCapture?.(event.pointerId)
  const cached = selectionGeometries.get(pageIndex)
  if (cached) {
    gesture.geometry = cached
    return
  }
  void pageGeometry(pageIndex)
    .then((geometry) => {
      if (selectionGesture !== gesture || request !== selectionRequest) return
      gesture.geometry = geometry
      if (!gesture.thresholdExceeded) return
      flushViewerSelection(gesture)
      if (!gesture.pointerActive) void commitViewerSelection(gesture)
    })
    .catch(() => {
      if (selectionGesture !== gesture || request !== selectionRequest) return
      if (layer.hasPointerCapture?.(event.pointerId)) layer.releasePointerCapture?.(event.pointerId)
      clearSelection()
    })
}

function onSelectionPointerMove(event: PointerEvent): void {
  const gesture = selectionGesture
  if (!gesture || gesture.pointerId !== event.pointerId) return
  updateViewerGesture(gesture, { x: event.clientX, y: event.clientY })
  if (gesture.thresholdExceeded && gesture.geometry) flushViewerSelection(gesture)
}

function onSelectionPointerUp(event: PointerEvent): void {
  const gesture = selectionGesture
  if (!gesture || gesture.pointerId !== event.pointerId) return
  updateViewerGesture(gesture, { x: event.clientX, y: event.clientY })
  gesture.pointerActive = false
  const layer = event.currentTarget as HTMLElement
  if (layer.hasPointerCapture?.(event.pointerId)) layer.releasePointerCapture?.(event.pointerId)
  if (!gesture.thresholdExceeded) {
    selectionGesture = null
    clearSelection()
    selectionClickCandidate = true
    return
  }
  selectionClickCandidate = false
  if (gesture.geometry) void commitViewerSelection(gesture)
}

function onSelectionPointerCancel(event: PointerEvent): void {
  if (selectionGesture?.pointerId !== event.pointerId) return
  clearSelection()
}

async function characterTextsForRange(pageIndex: number, from: number, to: number, request: number): Promise<readonly string[] | null> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) return null
  let cache = characterTextCache.get(pageIndex)
  if (!cache) {
    cache = new Map<number, string>()
    characterTextCache.set(pageIndex, cache)
  }
  const missing = []
  for (let charIndex = from; charIndex <= to; charIndex++) {
    if (!cache.has(charIndex)) missing.push({ pageIndex, charIndex, charCount: 1 })
  }
  if (missing.length > 0) {
    const texts = await runtime.getTextSlices(document, missing)
    if (request !== selectionRequest || renderDocument.value !== document || runtimeForDocument !== runtime) return null
    for (let index = 0; index < missing.length; index++) cache.set(missing[index].charIndex, texts[index] ?? "")
  }
  return Array.from({ length: to - from + 1 }, (_, index) => cache!.get(from + index) ?? "")
}

async function onSelectionClick(event: MouseEvent): Promise<void> {
  const isNativeDoubleClick = selectionClickCandidate && event.button === 0 && event.detail === 2
  selectionClickCandidate = false
  if (!isNativeDoubleClick) return
  event.preventDefault()
  const pageIndex = selectionPageIndex()
  const layer = event.currentTarget as HTMLElement
  const request = ++selectionRequest
  selectionRects.value = []
  if (selectionState.value.kind !== "none") emitSelection({ kind: "none" })
  try {
    const geometry = await pageGeometry(pageIndex)
    if (request !== selectionRequest) return
    const point = selectionPoint(event, layer)
    const charIndex = point ? resolveSelectionChar(geometry, point) : -1
    if (charIndex < 0) return
    const candidate = expandToWordBoundary(geometry, charIndex)
    if (!candidate) return
    const texts = await characterTextsForRange(pageIndex, candidate.from, candidate.to, request)
    const Segmenter = (Intl as typeof Intl & {
      Segmenter?: new (
        locales?: string | readonly string[],
        options?: { granularity: "word" },
      ) => { segment(input: string): Iterable<{ index: number; segment: string; isWordLike?: boolean }> }
    }).Segmenter
    if (!texts || request !== selectionRequest || !Segmenter) return
    const range = wordRangeFromCharacterTexts(candidate.from, charIndex, texts, new Segmenter(undefined, { granularity: "word" }))
    if (!range) return
    updateSelectionRects(geometry, range.from, range.to)
    selectionGesture = {
      request,
      pageIndex,
      pointerId: -1,
      layer,
      down: { x: event.clientX, y: event.clientY },
      latest: { x: event.clientX, y: event.clientY },
      pointerActive: false,
      thresholdExceeded: true,
      geometry,
      anchor: range.from,
      focus: range.to,
    }
    await commitViewerSelection(selectionGesture)
  } catch {
    // A failed word lookup must not revive a stale or partial selection.
  }
}

function onViewerCopy(event: ClipboardEvent): void {
  if (selectionState.value.kind !== "text" || !event.clipboardData) return
  event.preventDefault?.()
  event.clipboardData.setData("text/plain", selectionState.value.text)
}

function getDownloadFileName(): string {
  if (props.fileName?.trim()) return props.fileName.toLowerCase().endsWith(".pdf") ? props.fileName : `${props.fileName}.pdf`
  if (loadedFileName.value) return loadedFileName.value
  if (effectiveSource.value?.kind === "blob") {
    const file = effectiveSource.value.blob as File
    if (typeof file.name === "string" && file.name.trim()) {
      return file.name.toLowerCase().endsWith(".pdf") ? file.name : `${file.name}.pdf`
    }
  }
  const raw =
    (effectiveSource.value?.kind === "url" ? effectiveSource.value.url.split(/[?#]/)[0]?.split("/").pop() : undefined) ||
    "document.pdf"
  return raw.toLowerCase().endsWith(".pdf") ? raw : `${raw}.pdf`
}

function emitDiagnostic(event: Omit<PdfDiagnostic, "requestId" | "taskId">): void {
  emit("diagnostic", {
    ...event,
    requestId: latestRequestId,
    taskId: latestTaskId,
    runtimeId: runtimeForDocument?.runtimeId,
  })
}

function revokeObjectUrl(url: string): void {
  if (!url) return
  try {
    URL.revokeObjectURL(url)
    emitDiagnostic({ type: "object-url-revoked" })
  } catch {
    // The URL is already unusable; local state is still cleared below.
  }
}

function replacePageImage(nextUrl = ""): void {
  const previous = pageImageUrl.value
  pageImageUrl.value = nextUrl
  if (previous && previous !== nextUrl) revokeObjectUrl(previous)
}

function clearThumbnails(): void {
  for (const url of Object.values(thumbnailUrls.value)) revokeObjectUrl(url)
  thumbnailUrls.value = {}
}

function clearPrintPages(): void {
  printGeneration++
  printAbort?.abort()
  printAbort = undefined
  for (const url of Object.values(printPageUrls.value)) revokeObjectUrl(url)
  printPageUrls.value = {}
}

// Print must include every page, so print mode replaces the lazy single-page
// raster with a complete page cache before the print snapshot is taken.
async function renderPrintPages(requestId: number): Promise<void> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime || !printMode.value) return
  const generation = ++printGeneration
  printAbort?.abort()
  const controller = new AbortController()
  printAbort = controller
  for (let pageIndex = 0; pageIndex < document.pageCount; pageIndex++) {
    if (
      controller.signal.aborted ||
      generation !== printGeneration ||
      requestId !== latestRequestId ||
      document.id !== renderDocument.value?.id
    ) return
    if (printPageUrls.value[pageIndex + 1]) continue
    let url = ""
    try {
      const blob = await runtime.renderPage(document, pageIndex, {
        scale: PRINT_RENDER_SCALE,
        rotation: rotation.value,
        dpr: 1,
        signal: controller.signal,
      })
      if (
        controller.signal.aborted ||
        generation !== printGeneration ||
        requestId !== latestRequestId ||
        document.id !== renderDocument.value?.id
      ) return
      url = URL.createObjectURL(blob)
      printPageUrls.value = { ...printPageUrls.value, [pageIndex + 1]: url }
      url = ""
    } catch (error) {
      if (controller.signal.aborted || generation !== printGeneration || isAbortError(error)) return
      // A failed page must not abort the rest of the print set.
    } finally {
      if (url) revokeObjectUrl(url)
    }
  }
  if (printAbort === controller) printAbort = undefined
}

function setPrintMode(enabled: boolean): void {
  if (printMode.value === enabled) {
    if (enabled) void renderPrintPages(latestRequestId)
    return
  }
  printMode.value = enabled
  if (enabled) void renderPrintPages(latestRequestId)
  else clearPrintPages()
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
    if (printPageStyle.dataset) printPageStyle.dataset.officePrintPage = "pdf"
    document.head?.append(printPageStyle)
  }
  printPageStyle.textContent = `@page { size: ${Math.round(width)}px ${Math.round(height)}px; margin: 0; }`
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
    clearPrintPages()
  }
}

async function printAllPages(): Promise<void> {
  if (!renderDocument.value) return
  printMode.value = true
  await renderPrintPages(latestRequestId)
  if (typeof window !== "undefined" && typeof window.print === "function") window.print()
}

function cancelRenderWork(): void {
  renderGeneration++
  renderAbort?.abort()
  renderAbort = undefined
  searchGeneration++
  searchAbort?.abort()
  searchAbort = undefined
  thumbnailGeneration++
  thumbnailAbort?.abort()
  thumbnailAbort = undefined
}

function resetSearch(): void {
  searchResult.value = ""
  searchHits.value = []
  searchHitIndex.value = -1
}

function clearSearchAndRestoreFocus(): void {
  searchDraft.value = ""
  searchGeneration++
  searchAbort?.abort()
  searchAbort = undefined
  resetSearch()
  viewerRef.value?.focus()
}

function onViewerKeydown(event: KeyboardEvent): void {
  if (event.key === "Escape" && selectionState.value.kind !== "none") {
    clearSelection()
    return
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLocaleLowerCase() === "p") {
    event.preventDefault()
    void printAllPages()
    return
  }
  if (
    !props.showToolbar
    || !(event.ctrlKey || event.metaKey)
    || event.key.toLocaleLowerCase() !== "f"
  ) return
  event.preventDefault()
  searchInputRef.value?.focus()
}

async function releaseCurrentDocument(options: { closeDocument?: boolean } = {}): Promise<void> {
  cancelRenderWork()
  clearSelection()
  textLayerRequest += 1
  pageTextLayer.value = null
  selectionGeometries.clear()
  selectionGeometryRequests.clear()
  characterTextCache.clear()
  pendingAnchor = undefined
  replacePageImage()
  clearThumbnails()
  clearPrintPages()
  pageRenderState.value = "idle"
  pageRenderError.value = ""
  renderedPageNumber.value = 0
  const document = renderDocument.value
  const runtime = runtimeForDocument
  renderDocument.value = null
  runtimeForDocument = undefined
  numPages.value = 0
  activePage.value = 1
  pageInput.value = "1"
  verifiedBlob.value = null
  loadedFileName.value = ""
  resetSearch()
  if (document && runtime && options.closeDocument !== false) await runtime.closeDocument(document)
}

async function selectRuntime(): Promise<PdfRenderRuntime> {
  if (props.runtime) {
    if (ownedRuntime) {
      const runtime = ownedRuntime
      ownedRuntime = undefined
      ownedRuntimeWasmUrl = undefined
      await runtime.dispose()
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

function createLinkedAbort(parent?: AbortSignal): { controller: AbortController; unlink: () => void } {
  const controller = new AbortController()
  const abort = () => controller.abort()
  if (parent?.aborted) controller.abort()
  else parent?.addEventListener("abort", abort, { once: true })
  return { controller, unlink: () => parent?.removeEventListener("abort", abort) }
}

async function confirmImageDecodes(url: string): Promise<void> {
  if (typeof Image === "undefined") return
  const image = new Image()
  image.src = url
  if (typeof image.decode === "function") await image.decode()
  else {
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve()
      image.onerror = () => reject(new Error("Rendered PDF page is not a decodable image."))
    })
  }
  if (image.naturalWidth <= 0 || image.naturalHeight <= 0) {
    throw new Error("Rendered PDF page has no visible pixels.")
  }
}

/**
 * Viewport-center anchor of the rendered page. The bitmap only resizes once the
 * replacement arrives, so the anchor is held until that render lands.
 */
type PageAnchor = { xRatio: number; yRatio: number; clientX: number; clientY: number }

let pendingAnchor: (PageAnchor & { zoom: number; rotation: PdfRotation }) | undefined

function clampRatio(value: number): number {
  return Math.min(1, Math.max(0, value))
}

function measureAnchor(): PageAnchor | undefined {
  const frame = frameRef.value
  const image = pageImageRef.value
  if (!frame || !image) return undefined
  if (typeof frame.getBoundingClientRect !== "function" || typeof image.getBoundingClientRect !== "function") return undefined
  const rect = image.getBoundingClientRect()
  if (!(rect.width > 0) || !(rect.height > 0)) return undefined
  const frameRect = frame.getBoundingClientRect()
  const clientX = frameRect.left + frameRect.width / 2
  const clientY = frameRect.top + frameRect.height / 2
  return {
    xRatio: clampRatio((clientX - rect.left) / rect.width),
    yRatio: clampRatio((clientY - rect.top) / rect.height),
    clientX,
    clientY,
  }
}

async function restoreAnchor(zoomValue: number, rotationValue: PdfRotation): Promise<void> {
  const anchor = pendingAnchor
  if (!anchor || anchor.zoom !== zoomValue || anchor.rotation !== rotationValue) return
  pendingAnchor = undefined
  await nextTick()
  await nextPaint()
  const frame = frameRef.value
  const image = pageImageRef.value
  if (!frame || !image) return
  if (typeof frame.getBoundingClientRect !== "function" || typeof image.getBoundingClientRect !== "function") return
  const rect = image.getBoundingClientRect()
  const scrollLeft = typeof frame.scrollLeft === "number" ? frame.scrollLeft : 0
  const scrollTop = typeof frame.scrollTop === "number" ? frame.scrollTop : 0
  frame.scrollLeft = scrollLeft + rect.left + rect.width * anchor.xRatio - anchor.clientX
  frame.scrollTop = scrollTop + rect.top + rect.height * anchor.yRatio - anchor.clientY
}

function nextPaint(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof requestAnimationFrame === "function") requestAnimationFrame(() => resolve())
    else setTimeout(resolve, 0)
  })
}

function applyZoom(next: number | undefined): void {
  if (next === undefined || next === zoom.value) return
  const anchor = measureAnchor()
  pendingAnchor = anchor ? { ...anchor, zoom: next, rotation: rotation.value } : undefined
  zoom.value = next
}

function zoomIn(): void {
  applyZoom(zoomOptions.find((value) => value > zoom.value))
}

function zoomOut(): void {
  applyZoom([...zoomOptions].reverse().find((value) => value < zoom.value))
}

const touchGestures = createTouchGestureTracker({
  getZoom: () => zoom.value,
  onZoom: (next) => { zoom.value = next },
  onDoubleTap: () => { zoom.value = nextDoubleTapZoom(zoom.value) },
})

function onTouchPointerDown(event: PointerEvent): void { touchGestures.pointerDown(event) }
function onTouchPointerMove(event: PointerEvent): void { touchGestures.pointerMove(event) }
function onTouchPointerUp(event: PointerEvent): void { touchGestures.pointerUp(event) }
function onTouchPointerCancel(event: PointerEvent): void { touchGestures.pointerCancel(event) }

function rotateClockwise(): void {
  const anchor = measureAnchor()
  const next = ((rotation.value + 90) % 360) as PdfRotation
  pendingAnchor = anchor ? { ...anchor, zoom: zoom.value, rotation: next } : undefined
  rotation.value = next
}

function setActivePage(page: number): void {
  const next = Math.max(1, Math.min(numPages.value || 1, Math.trunc(page) || 1))
  if (next !== activePage.value) pendingAnchor = undefined
  activePage.value = next
  pageInput.value = String(next)
  emit("active-page-change", next)
}

function commitPageInput(): void {
  setActivePage(Number(pageInput.value))
}

async function handleDownload(): Promise<void> {
  if (!verifiedBlob.value) return
  downloadError.value = ""
  try {
    const url = URL.createObjectURL(verifiedBlob.value)
    const anchor = document.createElement("a")
    anchor.href = url
    anchor.download = getDownloadFileName()
    anchor.rel = "noopener"
    document.body.append(anchor)
    anchor.click()
    anchor.remove()
    setTimeout(() => revokeObjectUrl(url), 0)
    emitDiagnostic({
      type: "download",
      sourceKind: effectiveSource.value?.kind,
      bytes: verifiedBlob.value.size,
    })
  } catch (error) {
    downloadError.value = error instanceof Error ? error.message : String(error)
  }
}

async function renderCurrentPage(
  requestId: number,
  parentSignal?: AbortSignal,
  failDocumentOnError = false,
): Promise<void> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) return
  const pageNumber = activePage.value
  const renderZoom = zoom.value
  const renderRotation = rotation.value
  const renderDprValue = renderDpr.value
  const generation = ++renderGeneration
  renderAbort?.abort()
  const { controller, unlink } = createLinkedAbort(parentSignal)
  renderAbort = controller
  pageRenderError.value = ""
  // The previous bitmap stays mounted while this render runs, so the stage
  // never collapses to an empty white box between pages or zoom steps.
  pageRenderState.value = "rendering"
  emitDiagnostic({
    type: "render-start",
    pageNumber,
    pageCount: document.pageCount,
    zoom: renderZoom,
    rotation: renderRotation,
  })
  let candidateUrl = ""
  try {
    const blob = await runtime.renderPage(document, pageNumber - 1, {
      scale: renderZoom,
      rotation: renderRotation,
      dpr: renderDprValue,
      signal: controller.signal,
    })
    if (
      controller.signal.aborted ||
      generation !== renderGeneration ||
      requestId !== latestRequestId ||
      document.id !== renderDocument.value?.id
    ) return
    candidateUrl = URL.createObjectURL(blob)
    await confirmImageDecodes(candidateUrl)
    if (
      controller.signal.aborted ||
      generation !== renderGeneration ||
      requestId !== latestRequestId ||
      document.id !== renderDocument.value?.id
    ) return
    replacePageImage(candidateUrl)
    candidateUrl = ""
    renderedPageNumber.value = pageNumber
    renderedZoom.value = renderZoom
    renderedRotation.value = renderRotation
    pageRenderState.value = "rendered"
    void loadPageTextLayer(pageNumber)
    void restoreAnchor(renderZoom, renderRotation)
    emitDiagnostic({
      type: "render-success",
      pageNumber,
      pageCount: document.pageCount,
      zoom: renderZoom,
      rotation: renderRotation,
      bytes: blob.size,
    })
  } catch (error) {
    if (controller.signal.aborted || generation !== renderGeneration || requestId !== latestRequestId || isAbortError(error)) return
    pageRenderState.value = "error"
    pageRenderError.value = m.value.pdf.viewer.renderFailed
    // A stale bitmap no longer describes the current page or zoom.
    pendingAnchor = undefined
    replacePageImage()
    const source = effectiveSource.value
    const normalized = toPdfLoadError(error, source?.kind ?? "blob", source?.kind === "url" ? source.url : undefined)
    emitDiagnostic({ type: "render-error", pageNumber, pageCount: document.pageCount, error: normalized })
    if (failDocumentOnError) throw error
  } finally {
    if (candidateUrl) revokeObjectUrl(candidateUrl)
    unlink()
    if (renderAbort === controller) renderAbort = undefined
  }
}

async function renderThumbnails(requestId: number): Promise<void> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  if (!document || !runtime) return
  const generation = ++thumbnailGeneration
  thumbnailAbort?.abort()
  const controller = new AbortController()
  thumbnailAbort = controller
  let nextIndex = 0
  const worker = async () => {
    while (nextIndex < document.pageCount) {
      const pageIndex = nextIndex++
      let url = ""
      try {
        const blob = await runtime.renderThumbnail(document, pageIndex, { signal: controller.signal })
        if (
          controller.signal.aborted ||
          generation !== thumbnailGeneration ||
          requestId !== latestRequestId ||
          document.id !== renderDocument.value?.id
        ) return
        url = URL.createObjectURL(blob)
        thumbnailUrls.value = { ...thumbnailUrls.value, [pageIndex + 1]: url }
        url = ""
      } catch (error) {
        if (!controller.signal.aborted && !isAbortError(error)) {
          // A failed thumbnail must not hide an otherwise usable page.
        }
      } finally {
        if (url) revokeObjectUrl(url)
      }
    }
  }
  await Promise.all([worker(), worker()])
  if (thumbnailAbort === controller) thumbnailAbort = undefined
}

async function submitSearch(): Promise<void> {
  const query = searchDraft.value.trim()
  const document = renderDocument.value
  const runtime = runtimeForDocument
  searchAbort?.abort()
  const generation = ++searchGeneration
  searchHits.value = []
  searchHitIndex.value = -1
  if (!query) {
    searchResult.value = ""
    return
  }
  if (!document || !runtime) return
  const controller = new AbortController()
  searchAbort = controller
  searchResult.value = m.value.pdf.viewer.searchSearching
  emitDiagnostic({ type: "search-start" })
  try {
    const hits = await runtime.search(document, query, { signal: controller.signal })
    if (
      controller.signal.aborted ||
      generation !== searchGeneration ||
      document.id !== renderDocument.value?.id
    ) return
    searchHits.value = hits
    if (hits.length) {
      searchHitIndex.value = 0
      setActivePage(hits[0].pageIndex + 1)
      searchResult.value = formatOfficeMessage(m.value.pdf.viewer.searchResult, {
        current: 1,
        total: hits.length,
        page: hits[0].pageIndex + 1,
      })
    } else {
      searchResult.value = m.value.pdf.viewer.searchNoResults
    }
    emitDiagnostic({ type: "search-success", matches: hits.length })
  } catch (error) {
    if (controller.signal.aborted || generation !== searchGeneration || isAbortError(error)) return
    searchResult.value = m.value.pdf.viewer.searchError
    const source = effectiveSource.value
    emitDiagnostic({
      type: "search-error",
      error: toPdfLoadError(error, source?.kind ?? "blob", source?.kind === "url" ? source.url : undefined),
    })
  } finally {
    if (searchAbort === controller) searchAbort = undefined
  }
}

function moveSearch(delta: number): void {
  if (!searchHits.value.length) return
  const next = (searchHitIndex.value + delta + searchHits.value.length) % searchHits.value.length
  searchHitIndex.value = next
  const hit = searchHits.value[next]
  setActivePage(hit.pageIndex + 1)
  searchResult.value = formatOfficeMessage(m.value.pdf.viewer.searchResult, {
    current: next + 1,
    total: searchHits.value.length,
    page: hit.pageIndex + 1,
  })
}

function toOfficeSource(source: PdfSource): OfficeSource {
  return source.kind === "url"
    ? { kind: "url", url: source.url }
    : {
        kind: "file",
        name: source.fileName,
        file: {
          name: source.fileName,
          size: source.blob.size,
          type: source.blob.type,
          arrayBuffer: () => source.blob.arrayBuffer(),
        },
      }
}

async function load(): Promise<void> {
  const source = effectiveSource.value
  const requestId = ++latestRequestId
  pdfLoadCoordinator.cancel()
  latestTaskId = undefined
  isLoading.value = !!source
  loadError.value = null
  downloadError.value = ""
  rotation.value = 0
  zoom.value = Math.min(2, Math.max(0.5, props.defaultZoom))
  await releaseCurrentDocument()
  if (requestId !== latestRequestId || unmounted) return
  if (!source) {
    isLoading.value = false
    return
  }

  const task = pdfLoadCoordinator.start(toOfficeSource(source), {
    urlPolicy: props.urlPolicy as OfficeUrlPolicy | undefined,
  })
  latestTaskId = task.context.taskId
  emitDiagnostic({
    type: "load-start",
    sourceKind: source.kind,
    url: source.kind === "url" ? sanitizeOfficeUrl(source.url) : undefined,
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
    openedDocument = await selectedRuntime.openDocument(content, {
      signal: task.signal,
      password: props.password,
      onPassword: props.onPassword,
    })
    task.assertCurrent()
    if (openedDocument.pageCount <= 0) throw new Error(m.value.pdf.viewer.invalidPdf)
    runtimeForDocument = selectedRuntime
    renderDocument.value = openedDocument
    numPages.value = openedDocument.pageCount
    activePage.value = 1
    pageInput.value = "1"
    verifiedBlob.value = verified.blob
    loadedFileName.value = verified.fileName
    await renderCurrentPage(requestId, task.signal, true)
    task.assertCurrent()
    emit("document-load-success", openedDocument.pageCount)
    emit("active-page-change", 1)
    emitDiagnostic({
      type: "load-success",
      sourceKind: verified.sourceKind,
      url: verified.resolvedUrl,
      bytes: verified.bytes.byteLength,
      pageNumber: 1,
      pageCount: openedDocument.pageCount,
    })
    void renderThumbnails(requestId)
    if (printMode.value) void renderPrintPages(requestId)
  } catch (error) {
    if (openedDocument && selectedRuntime) {
      await selectedRuntime.closeDocument(openedDocument)
    }
    if (!task.isCurrent() || requestId !== latestRequestId || isAbortError(error)) return
    const normalized = toPdfLoadError(error, source.kind, source.kind === "url" ? source.url : undefined)
    const localizedError = {
      ...normalized,
      message: localizePdfErrorMessage(normalized, m.value.pdf.errors),
    }
    loadError.value = localizedError
    pageRenderState.value = "error"
    replacePageImage()
    clearThumbnails()
    verifiedBlob.value = null
    numPages.value = 0
    renderDocument.value = null
    runtimeForDocument = undefined
    emit("document-load-error", localizedError)
    emitDiagnostic({ type: "load-error", sourceKind: source.kind, error: localizedError })
  } finally {
    if (task.isCurrent() && requestId === latestRequestId) isLoading.value = false
    task.finish()
  }
}

watch(
  () => [props.src, props.source, props.urlPolicy, props.runtime, props.pdfiumWasmUrl, props.maxFileSize],
  () => void load(),
  { immediate: true },
)
watch(activePage, (page) => { pageInput.value = String(page) })
watch([activePage, zoom, rotation, renderDpr], () => {
  clearSelection()
  if (renderDocument.value && !isLoading.value) void renderCurrentPage(latestRequestId)
})
watch([zoom, rotation, renderDpr], () => {
  if (!printMode.value) return
  clearPrintPages()
  void renderPrintPages(latestRequestId)
})
watch([renderDocument, rotation], syncPrintPageSize, { immediate: true })

onMounted(() => {
  setupPrintDetection()
  stopDprWatch = watchDevicePixelRatio((next) => { renderDpr.value = next })
  const element = frameRef.value
  if (!element) return
  element.addEventListener("pointerdown", onTouchPointerDown, { passive: true })
  element.addEventListener("pointermove", onTouchPointerMove, { passive: true })
  element.addEventListener("pointerup", onTouchPointerUp)
  element.addEventListener("pointercancel", onTouchPointerCancel)
})

onBeforeUnmount(() => {
  unmounted = true
  stopPrintDetection?.()
  stopPrintDetection = undefined
  stopDprWatch?.()
  stopDprWatch = undefined
  touchGestures.reset()
  const frame = frameRef.value
  frame?.removeEventListener("pointerdown", onTouchPointerDown)
  frame?.removeEventListener("pointermove", onTouchPointerMove)
  frame?.removeEventListener("pointerup", onTouchPointerUp)
  frame?.removeEventListener("pointercancel", onTouchPointerCancel)
  latestRequestId++
  latestTaskId = undefined
  pdfLoadCoordinator.dispose()
  pdfTaskSequence.dispose()
  const runtime = ownedRuntime
  ownedRuntime = undefined
  const closeDocument = !runtime || runtimeForDocument !== runtime
  void releaseCurrentDocument({ closeDocument }).finally(() => runtime?.dispose())
})

defineExpose({
  setActivePage,
  clearSelection,
  getSelectionState: () => selectionState.value,
  print: printAllPages,
})
</script>

<template>
  <div
    ref="viewerRef"
    :class="cn('pdf-viewer', props.isDark && 'pdf-viewer--dark', printMode && 'pdf-viewer--print', props.className)"
    :style="viewerMinHeightStyle"
    data-testid="pdf-viewer"
    tabindex="0"
    @keydown="onViewerKeydown"
    @copy="onViewerCopy"
  >
    <div v-if="props.showToolbar" data-testid="pdf-toolbar" class="pdf-toolbar">
      <button
        data-testid="pdf-sidebar-toggle"
        type="button"
        :disabled="controlsDisabled"
        :aria-label="m.pdf.viewer.toggleThumbnails"
        :aria-pressed="showThumbnails"
        @click="showThumbnails = !showThumbnails"
      >☰</button>
      <span class="file-name">{{ getDownloadFileName() }}</span>
      <button data-testid="pdf-first-page" type="button" :disabled="controlsDisabled || activePage <= 1" :aria-label="m.common.page.first" @click="setActivePage(1)">«</button>
      <button data-testid="pdf-previous-page" type="button" :disabled="controlsDisabled || activePage <= 1" :aria-label="m.common.page.previous" @click="setActivePage(activePage - 1)">‹</button>
      <input
        v-model="pageInput"
        data-testid="pdf-page-input"
        class="page-input"
        type="number"
        min="1"
        :max="numPages || 1"
        :disabled="controlsDisabled"
        :aria-label="m.pdf.viewer.pageInput"
        @change="commitPageInput"
        @keydown.enter.prevent="commitPageInput"
      />
      <span class="page-indicator">/ <span data-testid="pdf-page-count">{{ numPages || '—' }}</span></span>
      <button data-testid="pdf-next-page" type="button" :disabled="controlsDisabled || activePage >= numPages" :aria-label="m.common.page.next" @click="setActivePage(activePage + 1)">›</button>
      <button data-testid="pdf-last-page" type="button" :disabled="controlsDisabled || activePage >= numPages" :aria-label="m.common.page.last" @click="setActivePage(numPages)">»</button>
      <button data-testid="pdf-zoom-out" type="button" :disabled="controlsDisabled || zoom <= zoomOptions[0]" :aria-label="m.common.zoom.out" @click="zoomOut">−</button>
      <select v-model.number="zoom" data-testid="pdf-zoom-select" :disabled="controlsDisabled" :aria-label="m.common.zoom.level">
        <option v-for="value in zoomOptions" :key="value" :value="value">{{ Math.round(value * 100) }}%</option>
      </select>
      <button data-testid="pdf-zoom-in" type="button" :disabled="controlsDisabled || zoom >= zoomOptions[zoomOptions.length - 1]" :aria-label="m.common.zoom.in" @click="zoomIn">+</button>
      <button v-if="props.showRotateControls" data-testid="pdf-rotate" type="button" :disabled="controlsDisabled" :aria-label="m.pdf.viewer.rotateClockwise" @click="rotateClockwise">⟳</button>
      <form class="pdf-search" role="search" @submit.prevent="submitSearch">
        <input
          ref="searchInputRef"
          v-model="searchDraft"
          data-testid="pdf-search-input"
          :disabled="controlsDisabled"
          :aria-label="m.pdf.viewer.searchLabel"
          :placeholder="m.pdf.viewer.searchPlaceholder"
          @keydown.esc.prevent="clearSearchAndRestoreFocus"
        />
        <button data-testid="pdf-search-submit" type="submit" :disabled="controlsDisabled || !searchDraft.trim()">{{ m.pdf.viewer.search }}</button>
        <button data-testid="pdf-search-prev" type="button" :disabled="!searchHits.length" :aria-label="m.common.search.previous" @click="moveSearch(-1)">↑</button>
        <button data-testid="pdf-search-next" type="button" :disabled="!searchHits.length" :aria-label="m.common.search.next" @click="moveSearch(1)">↓</button>
      </form>
      <span v-if="searchResult" data-testid="pdf-search-result" class="search-result" role="status" aria-atomic="true">{{ searchResult }}</span>
      <button v-if="props.showDownload" data-testid="pdf-download" type="button" :disabled="!verifiedBlob || !!loadError" :aria-label="m.pdf.viewer.download" @click="handleDownload">⬇</button>
    </div>

    <div data-testid="pdf-status" class="pdf-status" :data-state="viewerState" role="status" aria-live="polite" aria-atomic="true">{{ viewerState }}</div>
    <div v-if="!effectiveSource" class="pdf-empty">{{ m.pdf.viewer.noPdf }}</div>
    <div v-else-if="isLoading && !renderDocument" class="pdf-empty">{{ m.pdf.viewer.opening }}</div>
    <div v-else-if="loadError" data-testid="load-error" class="pdf-error" :data-error-code="loadError.code">
      {{ loadError.code }}: {{ loadErrorMessage }}
    </div>
    <div
      v-else-if="renderDocument"
      data-testid="pdf-document"
      class="pdf-body"
      :data-page-count="numPages"
      :data-active-page="activePage"
      :data-zoom="zoom"
      :data-rotation="rotation"
    >
      <aside v-if="showThumbnails" data-testid="pdf-thumbnails" class="pdf-thumbnails" :aria-label="m.pdf.viewer.thumbnails">
        <button
          v-for="page in pageLabels"
          :key="page"
          data-testid="pdf-thumbnail"
          type="button"
          :data-page-number="page"
          :aria-current="page === activePage ? 'page' : undefined"
          :class="['thumbnail', { active: page === activePage }]"
          @click="setActivePage(page)"
        >
          <img
            v-if="thumbnailUrls[page]"
            data-testid="pdf-thumbnail-image"
            :src="thumbnailUrls[page]"
            :alt="formatOfficeMessage(m.pdf.viewer.thumbnailAlt, { page })"
          />
          <span v-else class="thumbnail-placeholder">{{ page }}</span>
          <span>{{ formatOfficeMessage(m.pdf.viewer.pageLabel, { page }) }}</span>
        </button>
      </aside>
      <main ref="frameRef" class="pdf-frame-wrap">
        <div v-if="currentSearchHit" class="pdf-search-context" role="status">
          <span>{{ currentSearchHit.before }}</span><mark>{{ currentSearchHit.match }}</mark><span>{{ currentSearchHit.after }}</span>
        </div>
        <div
          data-testid="pdf-page"
          class="pdf-page-stage"
          :data-page-number="activePage"
          :data-render-state="pageRenderState"
        >
          <div v-if="pageRenderState === 'rendering'" class="page-loading">{{ formatOfficeMessage(m.pdf.viewer.renderingPage, { page: activePage }) }}</div>
          <div v-if="pageRenderError" class="pdf-error page-error">{{ pageRenderError }}</div>
          <div v-if="pageImageUrl" data-testid="pdf-page-render" class="pdf-page-render">
            <img
              ref="pageImageRef"
              data-testid="pdf-page-image"
              class="pdf-page-image"
              :src="pageImageUrl"
              :style="pageImageStyle"
              :alt="formatOfficeMessage(m.pdf.viewer.pageAlt, { page: renderedPageNumber })"
            />
            <div
              v-if="pageTextLayer"
              class="pdf-viewer__text-layer"
              data-testid="pdf-text-layer"
              :data-text-transform="textLayerTransform"
              :style="textLayerStyle"
            >
              <span
                v-for="glyph in pageTextLayer.glyphs"
                :key="`glyph-${glyph.charIndex}`"
                class="pdf-viewer__text-glyph"
                :class="{ 'pdf-viewer__text-glyph--unmapped': glyph.unmapped }"
                :data-char-index="glyph.charIndex"
                :aria-hidden="glyph.unmapped ? 'true' : undefined"
                :style="textGlyphStyle(glyph)"
              >{{ glyph.text }}</span>
            </div>
            <div
              class="pdf-viewer__selection-layer"
              data-testid="pdf-selection-layer"
              :aria-label="m.pdf.viewer.selectionLayer"
              @pointerdown="onSelectionPointerDown"
              @pointermove="onSelectionPointerMove"
              @pointerup="onSelectionPointerUp"
              @pointercancel="onSelectionPointerCancel"
              @click="onSelectionClick"
            >
              <div
                v-for="(rect, rectIndex) in selectionRects"
                :key="rectIndex"
                class="pdf-viewer__selection-rect"
                :style="selectionRectStyle(rect)"
              />
            </div>
          </div>
        </div>
      </main>
    </div>
    <div
      v-if="printPageEntries.length"
      class="pdf-print-pages"
      data-testid="pdf-print-pages"
      aria-hidden="true"
    >
      <figure
        v-for="entry in printPageEntries"
        :key="entry.page"
        class="pdf-print-page"
        data-testid="pdf-print-page"
        :data-page-number="entry.page"
      >
        <img :src="entry.url" :alt="formatOfficeMessage(m.pdf.viewer.pageAlt, { page: entry.page })" />
      </figure>
    </div>
    <div v-if="downloadError" class="download-error">{{ downloadError }}</div>
  </div>
</template>

<style scoped>
.pdf-viewer {
  --pdf-bg: #fff;
  --pdf-fg: #111;
  --pdf-border: #e5e5e5;
  --pdf-control-bg: #fff;
  --pdf-control-border: #d4d4d4;
  --pdf-canvas-bg: #e5e7eb;
  --pdf-panel-bg: #fafafa;
  --pdf-muted: #737373;
  --pdf-link: #1d4ed8;
  --pdf-page-shadow: 0 3px 14px rgb(0 0 0 / 20%);
  --pdf-page-filter: none;
  container: pdf-viewer / inline-size;
  display: flex; flex-direction: column; height: 100%; min-height: 0; background: var(--pdf-bg); color: var(--pdf-fg);
}
.pdf-viewer--dark {
  --pdf-bg: #0f172a;
  --pdf-fg: #e5e7eb;
  --pdf-border: #334155;
  --pdf-control-bg: #1e293b;
  --pdf-control-border: #475569;
  --pdf-canvas-bg: #020617;
  --pdf-panel-bg: #111827;
  --pdf-muted: #94a3b8;
  --pdf-link: #93c5fd;
  --pdf-page-shadow: 0 3px 14px rgb(0 0 0 / 60%);
  --pdf-page-filter: var(--pdf-dark-page-filter, invert(1) hue-rotate(180deg));
}
.pdf-toolbar { display: flex; align-items: center; gap: 7px; min-height: 48px; padding: 8px; border-bottom: 1px solid var(--pdf-border); flex-wrap: wrap; }
.pdf-toolbar button, .pdf-toolbar select, .pdf-toolbar input { height: 32px; border: 1px solid var(--pdf-control-border); border-radius: 6px; background: var(--pdf-control-bg); color: inherit; padding: 0 8px; }
.pdf-toolbar button { cursor: pointer; }
.pdf-toolbar button:focus-visible, .pdf-toolbar select:focus-visible, .pdf-toolbar input:focus-visible, .thumbnail:focus-visible { outline: 2px solid #2563eb; outline-offset: 2px; }
.pdf-toolbar button:disabled, .pdf-toolbar input:disabled, .pdf-toolbar select:disabled { opacity: .45; cursor: not-allowed; }
.file-name { max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 13px; font-weight: 600; margin-right: auto; }
.page-input { width: 58px; text-align: center; }
.page-indicator { font-size: 12px; white-space: nowrap; }
.pdf-search { display: flex; align-items: center; gap: 4px; }
.pdf-search input { width: 150px; }
.search-result { font-size: 12px; color: var(--pdf-link); white-space: nowrap; }
.pdf-status { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); clip-path: inset(50%); white-space: nowrap; }
.pdf-empty, .pdf-error { display: grid; place-items: center; flex: 1; min-height: 0; color: var(--pdf-muted); }
.pdf-error, .download-error { color: #b91c1c; }
.download-error { padding: 8px 12px; font-size: 12px; border-top: 1px solid var(--pdf-border); }
.pdf-body { display: flex; min-height: 0; flex: 1; min-width: 0; background: var(--pdf-canvas-bg); }
.pdf-thumbnails { width: 150px; flex: 0 0 150px; overflow: auto; border-right: 1px solid var(--pdf-border); padding: 10px; background: var(--pdf-panel-bg); }
.thumbnail { display: flex; width: 100%; min-height: 118px; margin-bottom: 10px; padding: 6px; align-items: center; justify-content: center; flex-direction: column; gap: 5px; overflow: hidden; border: 2px solid transparent; border-radius: 6px; background: transparent; color: var(--pdf-muted); font-size: 11px; cursor: pointer; }
.thumbnail img { display: block; max-width: 106px; max-height: 86px; object-fit: contain; border: 1px solid var(--pdf-control-border); background: var(--pdf-control-bg); box-shadow: 0 1px 3px rgb(0 0 0 / 12%); filter: var(--pdf-page-filter); }
.thumbnail-placeholder { display: grid; place-items: center; width: 74px; height: 84px; border: 1px solid var(--pdf-control-border); background: var(--pdf-control-bg); color: var(--pdf-muted); }
.thumbnail.active { border-color: #2563eb; background: color-mix(in srgb, #2563eb 12%, var(--pdf-panel-bg)); color: var(--pdf-link); font-weight: 700; }
.pdf-frame-wrap { flex: 1; min-width: 0; min-height: 0; overflow: auto; position: relative; padding: 24px; touch-action: pan-x pan-y; }
.pdf-page-stage { display: flex; min-width: min-content; min-height: 100%; align-items: flex-start; justify-content: center; position: relative; }
.pdf-page-render { position: relative; width: fit-content; margin: 0 auto; }
.pdf-viewer__text-layer { left: 0; position: absolute; top: 0; transform-origin: 0 0; pointer-events: none; user-select: text; z-index: 1; }
.pdf-viewer__text-glyph {
  color: transparent;
  -webkit-text-fill-color: transparent;
  line-height: 1;
  overflow: hidden;
  pointer-events: none;
  position: absolute;
  user-select: text;
  white-space: pre;
}
.pdf-viewer__text-glyph--unmapped { user-select: none; }
.pdf-viewer__selection-layer { inset: 0; position: absolute; cursor: text; touch-action: pan-x pan-y; z-index: 2; }
.pdf-viewer__selection-rect { background: rgba(37, 99, 235, .34); box-sizing: border-box; position: absolute; }
.pdf-page-image { display: block; max-width: none; height: auto; border: 0; background: var(--pdf-control-bg); box-shadow: var(--pdf-page-shadow); filter: var(--pdf-page-filter); }
.page-loading { position: absolute; z-index: 2; top: 12px; left: 50%; transform: translateX(-50%); border: 1px solid var(--pdf-control-border); border-radius: 999px; background: var(--pdf-control-bg); color: var(--pdf-link); padding: 5px 12px; font-size: 12px; white-space: nowrap; }
.page-error { min-height: 0; min-width: min(300px, 100%); border-radius: 8px; background: var(--pdf-control-bg); padding: 24px; box-shadow: 0 2px 8px rgb(0 0 0 / 12%); }
.pdf-search-context { position: sticky; top: 0; z-index: 3; width: fit-content; max-width: min(720px, 90%); margin: 0 auto 10px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; border: 1px solid var(--pdf-control-border); border-radius: 8px; background: var(--pdf-control-bg); color: var(--pdf-fg); font-size: 12px; padding: 6px 10px; }
.pdf-search-context mark { background: #fde047; color: #111827; padding: 1px 2px; }
.pdf-print-pages { display: none; }
@container pdf-viewer (max-width: 760px) {
  .pdf-body { flex-direction: column; }
  .pdf-thumbnails { width: auto; flex-basis: auto; display: flex; gap: 8px; border-right: 0; border-bottom: 1px solid var(--pdf-border); }
  .thumbnail { width: 108px; min-height: 104px; flex: 0 0 auto; margin-bottom: 0; }
  .pdf-frame-wrap { padding: 12px; }
  .pdf-search { order: 3; width: 100%; }
  .pdf-search input { flex: 1; width: auto; }
}

@media print {
  .pdf-viewer { height: auto !important; min-height: 0 !important; overflow: visible !important; background: #fff !important; color: #000 !important; }
  .pdf-toolbar, .pdf-thumbnails, .pdf-status, .pdf-search-context, .download-error, .page-loading { display: none !important; }
  .pdf-viewer--print .pdf-body { display: none !important; }
  .pdf-body { display: block !important; min-height: 0 !important; background: #fff !important; }
  .pdf-frame-wrap { overflow: visible !important; min-height: 0 !important; padding: 0 !important; }
  .pdf-page-stage { display: block !important; min-height: 0 !important; }
  .pdf-page-render { width: auto !important; margin: 0 auto; }
  .pdf-print-pages { display: block !important; }
  .pdf-print-page {
    break-after: page;
    margin: 0;
    padding: 0;
  }
  .pdf-print-page:last-child { break-after: auto; }
  .pdf-print-page img {
    display: block;
    width: 100% !important;
    max-width: 100% !important;
    height: auto !important;
    filter: none !important;
  }
  .pdf-page-image {
    width: 100% !important;
    max-width: 100% !important;
    height: auto !important;
    filter: none !important;
    box-shadow: none !important;
  }
}

@media (pointer: coarse) {
  .pdf-toolbar button,
  .pdf-toolbar select,
  .pdf-toolbar input {
    min-height: 44px;
  }
  .pdf-toolbar button {
    min-width: 44px;
  }
}
</style>
