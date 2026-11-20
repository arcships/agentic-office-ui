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
import {
  normalizePdfPageLayout,
  pdfPageIndexAtOffset,
  pdfSpreadPagePair,
  resolvePdfPageLayoutMetrics,
  type PdfPageLayout,
} from "../pdf/pdf-page-layout"
import { currentRenderDpr, watchDevicePixelRatio } from "../pdf/pdf-device-pixel-ratio"
import type { PdfViewerProps } from "../types"
import { formatOfficeMessage } from "@arcships/office-interaction"
import {
  DocumentViewerSidebar,
  DocumentViewerSidebarSkeleton,
  OfficeButton,
  OfficeEmptyState,
  OfficeErrorState,
  OfficeIcon,
  OfficeIconButton,
  OfficeInput,
  OfficeLoadingState,
  OfficePopover,
  OfficeStatusBar,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
  isEditableViewerTarget,
  useViewerKeys,
  useFileDrop,
  FILE_DROP_OUTLINE,
} from "@arcships/vue-ui"
import type { ViewerKeyCommand } from "@arcships/vue-ui"
import { providePdfMessages } from "../i18n"
import { PDF_FILE_EXTENSIONS, PDF_FILE_MIME_TYPES } from "../pdf-file-types"
import { usePdfThumbnails } from "../composables/use-pdf-thumbnails"
import PdfThumbnailStrip from "./PdfThumbnailStrip.vue"

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
  pageLayout: null,
  defaultPageLayout: "single" as PdfPageLayout,
  className: "",
  showToolbar: true,
  showDownload: true,
  showRotateControls: true,
  showPrint: true,
  showStatusBar: false,
  minHeight: undefined,
  isDark: false,
  locale: "en",
  messages: undefined,
})

const emit = defineEmits<{
  "document-load-success": [numPages: number]
  "document-load-error": [error: PdfLoadError]
  "active-page-change": [pageNumber: number]
  /** The toolbar page-layout control changed; hosts can drive the prop. */
  "update:pageLayout": [layout: PdfPageLayout]
  diagnostic: [event: PdfDiagnostic]
  selectionChange: [selection: PdfSelectionState]
  /** A supported PDF was dropped; the viewer loaded it through `load()`. */
  fileDrop: [file: File]
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
/** A file dropped on the viewer; wins over the `source`/`src` props until they change. */
const droppedSource = ref<PdfSource | null>(null)
const dropError = ref("")
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
const thumbnailPipeline = usePdfThumbnails({
  onRevoke: () => emitDiagnostic({ type: "object-url-revoked" }),
})
const thumbnailUrls = thumbnailPipeline.urls
const printPageUrls = ref<Record<number, string>>({})
const printMode = ref(false)
const selectionState = ref<PdfSelectionState>({ kind: "none" })
/** Viewport point of the last committed text selection; anchors the selection toolbar. */
const selectionAnchor = ref<{ x: number; y: number } | null>(null)
const selectionCopied = ref(false)
/**
 * Reka overlays need a real DOM. Node component tests mount the viewer with a
 * host renderer, so the floating toolbar only mounts in a browser.
 */
const overlaysAvailable = typeof window !== "undefined" && typeof document !== "undefined"
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
let renderAbort: AbortController | undefined
let searchAbort: AbortController | undefined
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
  if (droppedSource.value) return droppedSource.value
  if (props.source) return props.source
  if (props.src) return { kind: "url", url: props.src }
  return undefined
})

// ── Drag-and-drop file import ──────────────────────────────────────
// A dropped PDF becomes the effective source, so it loads through the same
// `load()` pipeline as the `source`/`src` props and the file picker.
const fileDrop = useFileDrop({
  accept: { extensions: PDF_FILE_EXTENSIONS, mimeTypes: PDF_FILE_MIME_TYPES },
  onFile: (file) => {
    dropError.value = ""
    emit("fileDrop", file)
    droppedSource.value = { kind: "blob", blob: file, fileName: file.name }
  },
  onReject: () => {
    dropError.value = m.value.pdf.viewer.unsupportedDrop
  },
})
const { isDragOver, onDragEnter, onDragOver, onDragLeave, onDrop } = fileDrop

const viewerState = computed(() => {
  if (!effectiveSource.value) return "idle"
  if (isLoading.value || pageRenderState.value === "rendering") return "loading"
  if (loadError.value || pageRenderState.value === "error") return "error"
  return renderDocument.value && pageImageUrl.value ? "ready" : "idle"
})
const controlsDisabled = computed(() => !renderDocument.value || !!loadError.value)
/** Encrypted documents get their own unlock state instead of a generic error. */
const encryptedError = computed(() => loadError.value?.code === "ENCRYPTED")
const errorTitle = computed(() =>
  encryptedError.value ? m.value.pdf.errors.encrypted : m.value.common.chrome.errorTitle
)
const errorDetail = computed(() =>
  encryptedError.value
    ? m.value.pdf.viewer.passwordRequired
    : loadError.value
      ? `${loadError.value.code}: ${loadErrorMessage.value}`
      : ""
)
const statusBarStatus = computed(() => {
  if (viewerState.value === "loading") return m.value.pdf.viewer.opening
  if (viewerState.value === "error") return loadErrorMessage.value
  return ""
})
/** Thumbnails are still rasterizing; show the shared sidebar skeleton above them. */
const thumbnailsLoading = computed(() =>
  !!renderDocument.value && Object.keys(thumbnailUrls.value).length === 0
)
const pdfPageLabels = ref<readonly string[]>([])
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

// ── Page layout (single / continuous / spread) ─────────────────────
// The active page owns the one interactive stage (bitmap, text layer, search
// hits, selection). The other pages are sized placeholders inside the same
// scroll content, and in a spread the facing page is rasterized next to it.
const PDF_VIEWER_ROW_GAP_PX = 16
const PDF_VIEWER_COLUMN_GAP_PX = 16
/** Placeholder window painted before the viewer has measured its scroll frame. */
const PDF_PLACEHOLDER_FALLBACK_VIEWPORT_PX = 1200
const internalPageLayout = ref<PdfPageLayout>(normalizePdfPageLayout(props.defaultPageLayout))
const pageLayoutMode = computed<PdfPageLayout>(() =>
  props.pageLayout == null
    ? internalPageLayout.value
    : normalizePdfPageLayout(props.pageLayout)
)
const viewerScrollTop = ref(0)
const viewerViewportHeight = ref(0)

/** On-screen size of a page at the current zoom and user rotation. */
function viewerPageSize(pageIndex: number): { width: number; height: number } {
  const page = renderDocument.value?.pages[pageIndex]
  if (!page) return { width: 1, height: 1 }
  const quarterTurns = ((page.rotation + rotation.value) / 90) % 4
  return {
    width: Math.max(1, Math.round((quarterTurns % 2 ? page.height : page.width) * zoom.value)),
    height: Math.max(1, Math.round((quarterTurns % 2 ? page.width : page.height) * zoom.value)),
  }
}

const viewerLayoutMetrics = computed(() =>
  resolvePdfPageLayoutMetrics({
    pageCount: renderDocument.value?.pageCount ?? 0,
    layout: pageLayoutMode.value,
    pageSize: viewerPageSize,
    rowGap: PDF_VIEWER_ROW_GAP_PX,
    columnGap: PDF_VIEWER_COLUMN_GAP_PX,
    paddingTop: 0,
    print: printMode.value,
  })
)

/** True when pages are laid out in a scrollable stack instead of one stage. */
const isStackedLayout = computed(() =>
  pageLayoutMode.value !== "single" && Boolean(renderDocument.value) && !printMode.value
)

/** Scroll content: every page is a row (or a facing pair) inside this box. */
const pageStackStyle = computed<Record<string, string> | undefined>(() => {
  if (!isStackedLayout.value) return undefined
  const metrics = viewerLayoutMetrics.value
  return {
    height: `${Math.max(0, metrics.totalHeightPx)}px`,
    minWidth: `${Math.max(0, metrics.maxRowWidthPx)}px`,
    position: "relative",
    width: "100%",
  }
})

function pageLayoutLeft(pageIndex: number): number {
  const metrics = viewerLayoutMetrics.value
  const rowIndex = metrics.pageGridRows[pageIndex] ?? 0
  const row = metrics.rows[rowIndex - 1]
  const offset = metrics.pageLefts[pageIndex] ?? 0
  if (!row) return offset
  return (metrics.maxRowWidthPx - row.widthPx) / 2 + offset
}

/** Placement of a page box inside the scroll content. */
function pageBoxStyle(pageNumber: number): Record<string, string> | undefined {
  if (!isStackedLayout.value) return undefined
  const pageIndex = pageNumber - 1
  const metrics = viewerLayoutMetrics.value
  return {
    position: "absolute",
    top: `${metrics.pageOffsets[pageIndex] ?? 0}px`,
    left: `${pageLayoutLeft(pageIndex)}px`,
    width: `${metrics.pageWidths[pageIndex] ?? 1}px`,
    height: `${metrics.pageHeights[pageIndex] ?? 1}px`,
  }
}

/** Bitmap width for a page rendered at `scale` (mirrors `pageImageStyle`). */
function pageBitmapStyle(pageNumber: number, scale: number): Record<string, string> | undefined {
  const page = renderDocument.value?.pages[pageNumber - 1]
  if (!page) return undefined
  const quarterTurns = ((page.rotation + rotation.value) / 90) % 4
  const width = (quarterTurns % 2 ? page.height : page.width) * scale
  return { width: `${Math.max(1, Math.round(width))}px` }
}

/** A stale bitmap must never paint inside the new page's box. */
const activeBitmapVisible = computed(() =>
  renderedPageNumber.value === activePage.value
)
/** Hides (but keeps mounted) the previous bitmap while another page rasterizes. */
const pageRenderStyle = computed<Record<string, string> | undefined>(() =>
  activeBitmapVisible.value ? undefined : { display: "none" }
)

/** Page rendered next to the active page while a spread is shown. */
const facingPageNumber = computed(() => {
  if (pageLayoutMode.value !== "spread" || printMode.value) return 0
  const count = renderDocument.value?.pageCount ?? 0
  if (count <= 0) return 0
  const pair = pdfSpreadPagePair(count, activePage.value - 1)
  const facingIndex = pair.find((index) => index + 1 !== activePage.value)
  return facingIndex === undefined ? 0 : facingIndex + 1
})
const facingPageUrl = ref("")
const facingPageState = ref<"idle" | "rendering" | "rendered" | "error">("idle")
let facingRenderAbort: AbortController | undefined
let facingRenderGeneration = 0

function replaceFacingImage(nextUrl = ""): void {
  const previous = facingPageUrl.value
  facingPageUrl.value = nextUrl
  if (previous && previous !== nextUrl) revokeObjectUrl(previous)
}

async function renderFacingPage(): Promise<void> {
  const document = renderDocument.value
  const runtime = runtimeForDocument
  const pageNumber = facingPageNumber.value
  const generation = ++facingRenderGeneration
  facingRenderAbort?.abort()
  facingRenderAbort = undefined
  if (!document || !runtime || !pageNumber) {
    replaceFacingImage()
    facingPageState.value = "idle"
    return
  }
  const controller = new AbortController()
  facingRenderAbort = controller
  facingPageState.value = "rendering"
  let candidate = ""
  try {
    const blob = await runtime.renderPage(document, pageNumber - 1, {
      scale: zoom.value,
      rotation: rotation.value,
      dpr: renderDpr.value,
      signal: controller.signal,
    })
    if (controller.signal.aborted || generation !== facingRenderGeneration) return
    candidate = URL.createObjectURL(blob)
    await confirmImageDecodes(candidate)
    if (controller.signal.aborted || generation !== facingRenderGeneration) return
    replaceFacingImage(candidate)
    candidate = ""
    facingPageState.value = "rendered"
  } catch (error) {
    if (controller.signal.aborted || generation !== facingRenderGeneration || isAbortError(error)) return
    facingPageState.value = "error"
    replaceFacingImage()
  } finally {
    if (candidate) revokeObjectUrl(candidate)
    if (facingRenderAbort === controller) facingRenderAbort = undefined
  }
}

// ── Scroll-driven active page (continuous / spread) ────────────────
let frameScrollFrame: number | undefined

function onFrameScroll(): void {
  if (!isStackedLayout.value) return
  if (frameScrollFrame !== undefined) return
  const flush = () => {
    frameScrollFrame = undefined
    const frame = frameRef.value
    if (!frame) return
    const top = typeof frame.scrollTop === "number" ? frame.scrollTop : 0
    const height = typeof frame.clientHeight === "number" ? frame.clientHeight : 0
    viewerScrollTop.value = top
    viewerViewportHeight.value = height
    const pageIndex = pdfPageIndexAtOffset(viewerLayoutMetrics.value.rows, top, height)
    const page = pageIndex + 1
    if (page >= 1 && page <= numPages.value && page !== activePage.value) {
      setActivePage(page, { reveal: false })
    }
  }
  if (typeof requestAnimationFrame === "function") {
    frameScrollFrame = requestAnimationFrame(flush)
    return
  }
  flush()
}

/**
 * Switch layout mode and stay on the page the reader is looking at.
 * The scroll offset has to follow: a spread row starts at its left page, so
 * keeping the raw offset would slide the reader onto the previous spread.
 */
function setPageLayout(mode: PdfPageLayout): void {
  const next = normalizePdfPageLayout(mode)
  if (next === pageLayoutMode.value) return
  if (props.pageLayout == null) internalPageLayout.value = next
  emit("update:pageLayout", next)
  void nextTick(() => scrollFrameToActivePage())
}

function onPageLayoutChange(event: Event): void {
  const value = (event.target as HTMLSelectElement | null)?.value
  if (value === "continuous" || value === "single" || value === "spread") setPageLayout(value)
}

function scrollFrameToActivePage(): void {
  const frame = frameRef.value
  if (!frame || !isStackedLayout.value) return
  const metrics = viewerLayoutMetrics.value
  const rowIndex = metrics.pageGridRows[activePage.value - 1] ?? 1
  const row = metrics.rows[rowIndex - 1]
  if (!row) return
  if (typeof frame.scrollTo === "function") frame.scrollTo({ top: row.topPx, behavior: "auto" })
  else frame.scrollTop = row.topPx
}

// Switching modes keeps the active page: the scroll offset follows the row
// that page now lives in (and returns to the top in single-page mode).
watch(pageLayoutMode, () => { void nextTick(scrollFrameToActivePage) })

/** Page boxes painted around the visible window (they are decoration only). */
const placeholderPages = computed(() => {
  if (!isStackedLayout.value || pageLayoutMode.value === "spread") return []
  const metrics = viewerLayoutMetrics.value
  // Before the frame has been measured (first paint, SSR, headless host) a
  // bounded fallback window is painted instead of every page.
  const viewport = viewerViewportHeight.value > 0
    ? viewerViewportHeight.value
    : PDF_PLACEHOLDER_FALLBACK_VIEWPORT_PX
  const from = viewerScrollTop.value - viewport
  const to = viewerScrollTop.value + viewport * 2
  const rows = metrics.rows.filter((row) => row.topPx + row.heightPx >= from && row.topPx <= to)
  const pages: number[] = []
  for (const row of rows) {
    for (const pageIndex of row.pageIndexes) {
      const page = pageIndex + 1
      if (page !== activePage.value) pages.push(page)
    }
  }
  return pages
})
const viewerMinHeightStyle = computed(() => {
  const value = props.minHeight
  const outline = isDragOver.value ? FILE_DROP_OUTLINE : undefined
  if (value === undefined) return outline ? { outline } : undefined
  const style: Record<string, string> = { minHeight: typeof value === "number" ? `${value}px` : value }
  if (outline) style.outline = outline
  return style
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
    // Image-only pages have no glyphs. An empty text layer would still mount a
    // selectable overlay, so treat "no glyphs" as "no selectable text".
    pageTextLayer.value = layer.glyphs.length > 0 ? layer : null
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
  selectionAnchor.value = null
  selectionCopied.value = false
  if (selectionState.value.kind === "none") return
  emitSelection({ kind: "none" })
}

const selectionAnchorStyle = computed<Record<string, string> | undefined>(() =>
  selectionAnchor.value
    ? {
        left: `${selectionAnchor.value.x}px`,
        top: `${selectionAnchor.value.y}px`,
      }
    : undefined
)

/** Close the floating selection toolbar when the reader dismisses it. */
function onSelectionPopoverOpen(open: boolean): void {
  if (!open) clearSelection()
}

async function copySelection(): Promise<void> {
  if (selectionState.value.kind !== "text") return
  const text = selectionState.value.text
  const clipboard = typeof navigator !== "undefined" ? navigator.clipboard : undefined
  if (!clipboard?.writeText) {
    selectionCopied.value = false
    return
  }
  try {
    await clipboard.writeText(text)
    selectionCopied.value = true
  } catch {
    selectionCopied.value = false
  }
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
    selectionCopied.value = false
    selectionAnchor.value = { x: gesture.latest.x, y: gesture.latest.y }
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
  thumbnailPipeline.clear()
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
  facingRenderGeneration++
  facingRenderAbort?.abort()
  facingRenderAbort = undefined
  searchGeneration++
  searchAbort?.abort()
  searchAbort = undefined
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
  // The print shortcut is PDF-specific, but it still follows the shared
  // scope rules: never fire from a focused input/textarea/select or when a
  // child already handled the key.
  if (event.defaultPrevented || isEditableViewerTarget(event.target)) return
  if ((event.ctrlKey || event.metaKey) && event.key.toLocaleLowerCase() === "p") {
    event.preventDefault()
    void printAllPages()
    return
  }
  viewerKeys.onKeydown(event)
}

/**
 * PDF keyboard map (shared with DOCX and PPTX through `useViewerKeys`):
 * PageDown/PageUp move one page, Home/End jump to the first/last page,
 * Ctrl/Cmd +/−/0 zoom, Ctrl/Cmd+F focuses search and Escape clears the
 * search before falling back to the current selection.
 */
const viewerKeys = useViewerKeys({
  pageNavigation: true,
  zoom: true,
  search: () => props.showToolbar,
  escape: true,
  onCommand: onViewerKeyCommand,
})

function onViewerKeyCommand(command: ViewerKeyCommand): boolean {
  switch (command) {
    case "page-next":
      return goToPage(activePage.value + 1)
    case "page-previous":
      return goToPage(activePage.value - 1)
    case "page-first":
      return goToPage(1)
    case "page-last":
      return goToPage(numPages.value)
    case "zoom-in":
      return stepZoom(1)
    case "zoom-out":
      return stepZoom(-1)
    case "zoom-reset":
      if (zoom.value === 1) return false
      applyZoom(1)
      return true
    case "search":
      if (!props.showToolbar) return false
      searchInputRef.value?.focus()
      return true
    case "escape":
      if (searchDraft.value || searchHits.value.length) {
        clearSearchAndRestoreFocus()
        return true
      }
      if (selectionState.value.kind !== "none") {
        clearSelection()
        return true
      }
      return false
    default:
      return false
  }
}

function goToPage(page: number): boolean {
  if (!renderDocument.value || numPages.value < 1) return false
  const target = Math.min(numPages.value, Math.max(1, page))
  if (target === activePage.value) return false
  setActivePage(target)
  return true
}

function stepZoom(direction: 1 | -1): boolean {
  const next = direction > 0
    ? zoomOptions.find((value) => value > zoom.value)
    : [...zoomOptions].reverse().find((value) => value < zoom.value)
  if (next === undefined) return false
  applyZoom(next)
  return true
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
  replaceFacingImage()
  facingPageState.value = "idle"
  clearThumbnails()
  pdfPageLabels.value = []
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

function setActivePage(page: number, options: { reveal?: boolean } = {}): void {
  const next = Math.max(1, Math.min(numPages.value || 1, Math.trunc(page) || 1))
  const changed = next !== activePage.value
  if (changed) pendingAnchor = undefined
  activePage.value = next
  pageInput.value = String(next)
  emit("active-page-change", next)
  // Continuous and spread layouts scroll; a toolbar, thumbnail, keyboard or
  // search jump has to bring the new page's row into view. Scroll tracking
  // passes `reveal: false` so the reader's own scroll is never fought.
  if (changed && options.reveal !== false) scrollFrameToActivePage()
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

async function loadPdfNavigation(
  runtime: PdfRenderRuntime,
  document: PdfRenderDocument,
  requestId: number,
): Promise<void> {
  pdfPageLabels.value = []
  if (typeof runtime.getPageLabels !== "function") return
  try {
    const labels = await runtime.getPageLabels(document)
    if (requestId !== latestRequestId || document.id !== renderDocument.value?.id) return
    pdfPageLabels.value = labels
  } catch {
    // Page labels are a navigation aid; a document without them still renders.
  }
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
  dropError.value = ""
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
    void thumbnailPipeline.load(selectedRuntime, openedDocument)
    void loadPdfNavigation(selectedRuntime, openedDocument, requestId)
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

// A host-owned source supersedes a file dropped earlier; the load watcher below
// then picks up the props path again.
watch(
  () => [props.src, props.source],
  () => { droppedSource.value = null },
)
watch(
  () => [effectiveSource.value, props.urlPolicy, props.runtime, props.pdfiumWasmUrl, props.maxFileSize],
  () => void load(),
  { immediate: true },
)
watch(activePage, (page) => { pageInput.value = String(page) })
watch([activePage, zoom, rotation, renderDpr], () => {
  clearSelection()
  if (renderDocument.value && !isLoading.value) void renderCurrentPage(latestRequestId)
})
// A spread always shows its facing page rasterized; the page mode, zoom,
// rotation or display scale decides which page that is.
watch(
  [facingPageNumber, zoom, rotation, renderDpr, renderDocument],
  () => { void renderFacingPage() },
)
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
  replaceFacingImage()
  facingRenderAbort?.abort()
  facingRenderAbort = undefined
  if (frameScrollFrame !== undefined && typeof cancelAnimationFrame === "function") {
    cancelAnimationFrame(frameScrollFrame)
  }
  frameScrollFrame = undefined
  latestTaskId = undefined
  pdfLoadCoordinator.dispose()
  pdfTaskSequence.dispose()
  thumbnailPipeline.dispose()
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
    :data-office-theme="props.isDark ? 'dark' : 'light'"
    :data-file-drop-active="isDragOver ? 'true' : undefined"
    data-testid="pdf-viewer"
    tabindex="0"
    @keydown="onViewerKeydown"
    @copy="onViewerCopy"
    @dragenter="onDragEnter"
    @dragover="onDragOver"
    @dragleave="onDragLeave"
    @drop="onDrop"
  >
    <OfficeToolbar
      v-if="props.showToolbar"
      class="pdf-toolbar"
      data-testid="pdf-toolbar"
      :label="m.common.chrome.toolbar"
    >
      <template #leading>
        <OfficeToolbarGroup gap="sm" class="pdf-toolbar__leading">
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            icon="columns-2"
            data-testid="pdf-sidebar-toggle"
            :label="m.pdf.viewer.toggleThumbnails"
            :active="showThumbnails"
            :disabled="controlsDisabled"
            @click="showThumbnails = !showThumbnails"
          />
          <span class="pdf-file-name" :title="getDownloadFileName()">{{ getDownloadFileName() }}</span>
          <OfficeToolbarGroup gap="sm" class="pdf-page-cluster" :label="m.common.page.navigation">
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pdf-first-page"
              :label="m.common.page.first"
              :disabled="controlsDisabled || activePage <= 1"
              @click="setActivePage(1)"
            >«</OfficeIconButton>
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pdf-previous-page"
              :label="m.common.page.previous"
              :disabled="controlsDisabled || activePage <= 1"
              @click="setActivePage(activePage - 1)"
            >‹</OfficeIconButton>
            <OfficeInput
              v-model="pageInput"
              class="pdf-page-input"
              size="sm"
              type="number"
              min="1"
              :max="numPages || 1"
              :disabled="controlsDisabled"
              :aria-label="m.pdf.viewer.pageInput"
              data-testid="pdf-page-input"
              @change="commitPageInput"
              @keydown.enter.prevent="commitPageInput"
            />
            <span class="pdf-page-total">{{ m.common.page.of }} <span data-testid="pdf-page-count">{{ numPages || '—' }}</span></span>
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pdf-next-page"
              :label="m.common.page.next"
              :disabled="controlsDisabled || activePage >= numPages"
              @click="setActivePage(activePage + 1)"
            >›</OfficeIconButton>
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pdf-last-page"
              :label="m.common.page.last"
              :disabled="controlsDisabled || activePage >= numPages"
              @click="setActivePage(numPages)"
            >»</OfficeIconButton>
          </OfficeToolbarGroup>
          <select
            data-testid="pdf-page-layout"
            class="office-ui pdf-page-layout"
            :value="pageLayoutMode"
            :disabled="controlsDisabled"
            :aria-label="m.pdf.viewer.pageLayout"
            @change="onPageLayoutChange"
          >
            <option value="continuous">{{ m.pdf.viewer.pageLayoutContinuous }}</option>
            <option value="single">{{ m.pdf.viewer.pageLayoutSingle }}</option>
            <option value="spread">{{ m.pdf.viewer.pageLayoutSpread }}</option>
          </select>
        </OfficeToolbarGroup>
      </template>

      <template #trailing>
        <OfficeToolbarGroup gap="sm" class="pdf-toolbar__actions">
          <OfficeToolbarGroup gap="sm" class="pdf-zoom-cluster" :label="m.common.zoom.level">
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              icon="minus"
              data-testid="pdf-zoom-out"
              :label="m.common.zoom.out"
              :disabled="controlsDisabled || zoom <= zoomOptions[0]"
              @click="zoomOut"
            />
            <select
              v-model.number="zoom"
              data-testid="pdf-zoom-select"
              class="office-ui pdf-zoom-select"
              :disabled="controlsDisabled"
              :aria-label="m.common.zoom.level"
            >
              <option v-for="value in zoomOptions" :key="value" :value="value">{{ Math.round(value * 100) }}%</option>
            </select>
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              icon="plus"
              data-testid="pdf-zoom-in"
              :label="m.common.zoom.in"
              :disabled="controlsDisabled || zoom >= zoomOptions[zoomOptions.length - 1]"
              @click="zoomIn"
            />
          </OfficeToolbarGroup>
          <OfficeToolbarSeparator v-if="props.showRotateControls" />
          <OfficeIconButton
            v-if="props.showRotateControls"
            variant="ghost"
            size="icon-sm"
            data-testid="pdf-rotate"
            :label="m.pdf.viewer.rotateClockwise"
            :disabled="controlsDisabled"
            @click="rotateClockwise"
          >⟳</OfficeIconButton>
          <OfficeToolbarSeparator />
          <form class="pdf-search" role="search" @submit.prevent="submitSearch">
            <span class="office-ui office-input-control pdf-search__field">
              <input
                ref="searchInputRef"
                v-model="searchDraft"
                class="office-ui office-input pdf-search__input"
                type="search"
                autocomplete="off"
                spellcheck="false"
                :disabled="controlsDisabled"
                :aria-label="m.pdf.viewer.searchLabel"
                :placeholder="m.pdf.viewer.searchPlaceholder"
                data-testid="pdf-search-input"
                @keydown.esc.prevent="clearSearchAndRestoreFocus"
              />
            </span>
            <OfficeButton
              variant="outline"
              size="sm"
              type="submit"
              data-testid="pdf-search-submit"
              :disabled="controlsDisabled || !searchDraft.trim()"
            >{{ m.pdf.viewer.search }}</OfficeButton>
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pdf-search-prev"
              :label="m.common.search.previous"
              :disabled="!searchHits.length"
              @click="moveSearch(-1)"
            >↑</OfficeIconButton>
            <OfficeIconButton
              variant="ghost"
              size="icon-sm"
              data-testid="pdf-search-next"
              :label="m.common.search.next"
              :disabled="!searchHits.length"
              @click="moveSearch(1)"
            >↓</OfficeIconButton>
          </form>
          <span v-if="searchResult" data-testid="pdf-search-result" class="search-result" role="status" aria-atomic="true">{{ searchResult }}</span>
          <OfficeToolbarSeparator />
          <OfficeIconButton
            v-if="props.showPrint"
            variant="ghost"
            size="icon-sm"
            icon="printer"
            data-testid="pdf-print"
            :label="m.pdf.viewer.print"
            :disabled="controlsDisabled"
            @click="printAllPages"
          />
          <OfficeIconButton
            v-if="props.showDownload"
            variant="ghost"
            size="icon-sm"
            icon="download"
            data-testid="pdf-download"
            :label="m.pdf.viewer.download"
            :disabled="!verifiedBlob || !!loadError"
            @click="handleDownload"
          />
        </OfficeToolbarGroup>
      </template>
    </OfficeToolbar>

    <div data-testid="pdf-status" class="pdf-status" :data-state="viewerState" role="status" aria-live="polite" aria-atomic="true">{{ viewerState }}</div>
    <div v-if="dropError" data-testid="pdf-drop-error" class="pdf-drop-error" role="alert">{{ dropError }}</div>
    <OfficeEmptyState
      v-if="!effectiveSource"
      class="pdf-state pdf-state--empty"
      data-testid="pdf-empty"
      :title="m.pdf.viewer.noPdf"
      icon="file-text"
    />
    <OfficeLoadingState
      v-else-if="isLoading && !renderDocument"
      class="pdf-state pdf-state--loading"
      data-testid="pdf-loading"
      :title="m.pdf.viewer.opening"
    />
    <OfficeErrorState
      v-else-if="loadError"
      class="pdf-state pdf-state--error"
      data-testid="load-error"
      :data-error-code="loadError.code"
      :data-encrypted="encryptedError ? 'true' : undefined"
      :title="errorTitle"
      :description="errorDetail"
      :icon="encryptedError ? 'lock' : 'triangle-alert'"
      :show-retry="false"
    />
    <div
      v-else-if="renderDocument"
      data-testid="pdf-document"
      class="pdf-body"
      :data-page-count="numPages"
      :data-active-page="activePage"
      :data-zoom="zoom"
      :data-rotation="rotation"
    >
      <DocumentViewerSidebar
        :open="showThumbnails"
        :close-on-escape="false"
        :label="m.pdf.viewer.thumbnails"
        class="pdf-sidebar"
        data-testid="pdf-thumbnails-sidebar"
        @update:open="showThumbnails = $event"
      >
        <DocumentViewerSidebarSkeleton v-if="thumbnailsLoading" :items="3" />
        <PdfThumbnailStrip
          :pageCount="numPages"
          :labels="pdfPageLabels"
          :urls="thumbnailUrls"
          :activePage="activePage"
          :label="m.pdf.viewer.thumbnails"
          :thumbnailAlt="m.pdf.viewer.thumbnailAlt"
          :pageLabel="m.pdf.viewer.pageLabel"
          @select="setActivePage"
        />
      </DocumentViewerSidebar>
      <main
        ref="frameRef"
        class="pdf-frame-wrap"
        :data-page-layout="pageLayoutMode"
        @scroll.passive="onFrameScroll"
      >
        <div v-if="currentSearchHit" class="pdf-search-context" role="status">
          <span>{{ currentSearchHit.before }}</span><mark>{{ currentSearchHit.match }}</mark><span>{{ currentSearchHit.after }}</span>
        </div>
        <div
          class="pdf-page-stack"
          :class="{ 'pdf-page-stack--stacked': isStackedLayout }"
          :style="pageStackStyle"
        >
          <div
            v-for="page in placeholderPages"
            :key="`placeholder-${page}`"
            class="pdf-page-box"
            data-testid="pdf-page-placeholder"
            :data-page-number="page"
            :style="pageBoxStyle(page)"
          >{{ formatOfficeMessage(m.pdf.viewer.pageLabel, { page }) }}</div>
          <div
            data-testid="pdf-page"
            class="pdf-page-stage"
            :class="{ 'pdf-page-stage--stacked': isStackedLayout }"
            :style="isStackedLayout ? pageBoxStyle(activePage) : undefined"
            :data-page-number="activePage"
            :data-page-rotation="renderedPage?.rotation ?? 0"
            :data-render-state="pageRenderState"
          >
            <div v-if="pageRenderState === 'rendering'" class="page-loading">{{ formatOfficeMessage(m.pdf.viewer.renderingPage, { page: activePage }) }}</div>
            <div v-if="pageRenderError" class="pdf-error page-error">{{ pageRenderError }}</div>
            <div v-if="pageImageUrl" data-testid="pdf-page-render" class="pdf-page-render" :style="pageRenderStyle">
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
          <div
            v-if="facingPageNumber"
            class="pdf-page-box pdf-page-facing"
            data-testid="pdf-facing-page"
            :data-page-number="facingPageNumber"
            :data-render-state="facingPageState"
            :style="pageBoxStyle(facingPageNumber)"
          >
            <img
              v-if="facingPageUrl"
              data-testid="pdf-facing-page-image"
              class="pdf-page-image"
              :src="facingPageUrl"
              :style="pageBitmapStyle(facingPageNumber, zoom)"
              :alt="formatOfficeMessage(m.pdf.viewer.pageAlt, { page: facingPageNumber })"
            />
            <span v-else class="pdf-page-box__label">{{ formatOfficeMessage(m.pdf.viewer.pageLabel, { page: facingPageNumber }) }}</span>
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
    <OfficeStatusBar
      v-if="props.showStatusBar"
      class="pdf-status-bar"
      data-testid="pdf-status-bar"
      :page="activePage"
      :total-pages="numPages"
      :status="statusBarStatus"
      :file-name="getDownloadFileName()"
      :busy="viewerState === 'loading'"
    />
    <OfficePopover
      v-if="overlaysAvailable && selectionState.kind === 'text' && selectionAnchor"
      :open="true"
      side="top"
      align="center"
      :side-offset="8"
      content-class="pdf-selection-popover"
      data-testid="pdf-selection-popover"
      @update:open="onSelectionPopoverOpen"
    >
      <template #trigger>
        <span class="pdf-selection-anchor" :style="selectionAnchorStyle" aria-hidden="true" />
      </template>
      <div
        class="pdf-selection-popover__body"
        data-testid="pdf-selection-popover"
        role="toolbar"
        :aria-label="m.pdf.viewer.selectionActions"
      >
        <p class="pdf-selection-popover__text">{{ selectionState.text }}</p>
        <div class="pdf-selection-popover__actions">
          <span
            v-if="selectionCopied"
            class="pdf-selection-popover__status"
            role="status"
          >{{ m.pdf.viewer.copiedSelection }}</span>
          <OfficeButton
            variant="outline"
            size="sm"
            data-testid="pdf-selection-copy"
            @click="copySelection"
          >
            <OfficeIcon name="copy" />
            {{ m.pdf.viewer.copySelection }}
          </OfficeButton>
        </div>
      </div>
    </OfficePopover>
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
/* Keyboard focus lands on the viewer root; keep a visible ring around it. */
.pdf-viewer:focus-visible { outline: 2px solid var(--pdf-link); outline-offset: -2px; }
/* The toolbar is the shared `OfficeToolbar`; this class stays because it is the
   public print/touch hook (`.pdf-toolbar`). Geometry, colours and control
   metrics come from the `--office-*` tokens. */
.pdf-toolbar { min-height: 48px; font-size: var(--office-text-base); line-height: var(--office-leading-base); }
.pdf-toolbar__leading { min-width: 0; }
.pdf-toolbar__actions { min-width: 0; justify-content: flex-end; }
.pdf-file-name {
  max-width: 220px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--office-foreground);
  font-size: var(--office-text-sm);
  font-weight: 600;
}
.pdf-page-cluster { flex-wrap: nowrap; }
.pdf-page-input {
  inline-size: 3.625rem;
  min-inline-size: 0;
  text-align: center;
}
.pdf-page-input :deep(input) { text-align: center; }
.pdf-page-total {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
}
.pdf-page-layout,
.pdf-zoom-select {
  block-size: var(--office-control-height-xs);
  box-sizing: border-box;
  min-block-size: 0;
  min-height: 0;
  min-inline-size: 0;
  padding-block: 0;
  padding-inline: var(--office-ui-space-2);
  border: var(--office-border-width) solid var(--office-input);
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  color: var(--office-foreground);
  font: inherit;
  font-size: var(--office-text-sm);
  box-shadow: var(--office-shadow-xs);
  cursor: pointer;
}
.pdf-page-layout { max-inline-size: 9rem; }
.pdf-zoom-select { inline-size: 6.5rem; font-variant-numeric: tabular-nums; }
.pdf-page-layout:focus-visible,
.pdf-zoom-select:focus-visible {
  border-color: var(--office-ring);
  outline: var(--office-ring-width) solid color-mix(in oklab, var(--office-ring) 40%, transparent);
  outline-offset: var(--office-ring-offset-width);
}
.pdf-page-layout:disabled,
.pdf-zoom-select:disabled { cursor: not-allowed; opacity: 0.64; }
.pdf-zoom-cluster { flex-wrap: nowrap; }
.pdf-search { display: flex; align-items: center; gap: var(--office-ui-space-1); min-inline-size: 0; }
.pdf-search__field {
  block-size: var(--office-control-height-xs);
  box-sizing: border-box;
  inline-size: 9.5rem;
  min-block-size: 0;
  min-height: 0;
  min-inline-size: 0;
}
.pdf-search__input {
  block-size: 100%;
  min-block-size: 0;
  min-height: 0;
  inline-size: 100%;
  padding-block: 0;
  line-height: 1.4;
}
.search-result { font-size: var(--office-text-xs); color: var(--office-muted-foreground); white-space: nowrap; }
.pdf-status { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); clip-path: inset(50%); white-space: nowrap; }
.pdf-status-bar { flex: 0 0 auto; }
.pdf-drop-error { padding: 8px 12px; font-size: 12px; color: #b91c1c; border-bottom: 1px solid var(--pdf-border); background: rgba(220, 38, 38, 0.08); }
.pdf-state { flex: 1 1 auto; min-height: 0; }
.download-error { padding: 8px 12px; font-size: 12px; color: #b91c1c; border-top: 1px solid var(--pdf-border); }
.pdf-body { display: flex; position: relative; min-height: 0; flex: 1; min-width: 0; background: var(--pdf-canvas-bg); }
.pdf-sidebar { --pdf-border: var(--office-border); }
.pdf-frame-wrap { flex: 1; min-width: 0; min-height: 0; overflow: auto; position: relative; padding: 24px; touch-action: pan-x pan-y; }
/* Single-page mode paints one stage exactly like before; the stack wrapper only
   exists so continuous and spread modes share the same template. */
.pdf-page-stack { display: contents; }
.pdf-page-stack--stacked { display: block; margin: 0 auto; }
.pdf-page-stage { display: flex; min-width: min-content; min-height: 100%; align-items: flex-start; justify-content: center; position: relative; }
.pdf-page-stage--stacked { min-height: 0; min-width: 0; }
/* Sized page boxes: the active stage and the spread's facing page. */
.pdf-page-box {
  align-items: center;
  border: 1px solid var(--pdf-border);
  box-sizing: border-box;
  display: flex;
  justify-content: center;
  overflow: hidden;
  box-shadow: var(--pdf-page-shadow);
  background: var(--pdf-canvas-bg);
  color: var(--pdf-muted);
  font-size: 12px;
}
.pdf-page-facing { position: absolute; }
.pdf-page-box__label { padding: 4px 8px; }
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
/* Floating text-selection toolbar. The zero-size fixed anchor is the Reka
   popover trigger, so the surface opens next to the pointer that ended the
   selection without owning any page geometry. */
.pdf-selection-anchor {
  position: fixed;
  inline-size: 0;
  block-size: 0;
  pointer-events: none;
}
.pdf-selection-popover { max-inline-size: min(22rem, calc(100vw - 2rem)); }
.pdf-selection-popover__body {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-2);
  min-inline-size: 0;
}
.pdf-selection-popover__text {
  max-block-size: 4.5rem;
  margin: 0;
  overflow: hidden;
  color: var(--office-foreground);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.pdf-selection-popover__actions {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: var(--office-ui-space-2);
}
.pdf-selection-popover__status {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}
.pdf-print-pages { display: none; }
@container pdf-viewer (max-width: 1100px) {
  /* The parity/reference toolbar has no file label or first/last pair; dropping
     them at laptop and tablet widths keeps every remaining command on one
     48px row. Page input, prev/next and Home/End keep full navigation. */
  .pdf-file-name { display: none; }
  .pdf-toolbar :deep([data-testid="pdf-first-page"]),
  .pdf-toolbar :deep([data-testid="pdf-last-page"]) { display: none; }
  .pdf-search__field { inline-size: 7.5rem; }
  .pdf-zoom-select { inline-size: 5.5rem; }
  .pdf-page-layout { max-inline-size: 7.5rem; }
}
@container pdf-viewer (max-width: 760px) {
  .pdf-body { flex-direction: column; }
  .pdf-frame-wrap { padding: 12px; }
  .pdf-file-name { max-inline-size: 10rem; }
  .pdf-page-layout { max-inline-size: 7rem; }
}
@container pdf-viewer (max-width: 520px) {
  .pdf-file-name { display: none; }
  .pdf-page-layout { max-inline-size: 6rem; }
  /* Keep paging usable without three rows of navigation chrome. The first /
     last buttons stay available at desktop widths and through the keyboard
     map (Home/End) on narrow screens. */
  .pdf-toolbar :deep([data-testid="pdf-first-page"]),
  .pdf-toolbar :deep([data-testid="pdf-last-page"]) { display: none; }
  .pdf-toolbar__actions { flex: 1 1 100%; justify-content: flex-end; }
  .pdf-search { order: 3; flex: 1 1 100%; }
  .pdf-search__field { flex: 1 1 auto; inline-size: auto; }
  .search-result { flex: 1 1 100%; order: 4; }
}
@container pdf-viewer (max-width: 390px) {
  .pdf-page-total { font-size: var(--office-text-xs); }
  .pdf-page-layout { max-inline-size: 4.75rem; }
  .pdf-zoom-select { inline-size: 5.25rem; }
}

@media print {
  .pdf-viewer { height: auto !important; min-height: 0 !important; overflow: visible !important; background: #fff !important; color: #000 !important; }
  .pdf-toolbar, .pdf-thumbnails, .pdf-sidebar, .pdf-status, .pdf-status-bar, .pdf-search-context, .download-error, .page-loading { display: none !important; }
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
  .pdf-toolbar :deep(button),
  .pdf-toolbar :deep(select),
  .pdf-toolbar :deep(input) {
    min-height: 44px;
  }
  .pdf-toolbar :deep(button) {
    min-width: 44px;
  }
}
</style>
