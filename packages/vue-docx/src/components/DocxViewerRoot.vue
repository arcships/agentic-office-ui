<template>
  <div
    ref="scrollContainerRef"
    class="docx-viewer-root"
    :data-docx-page-count="pageCount"
    :data-docx-document-id="documentId"
    :data-docx-page-count-target="paginationPageCountTarget.targetPageCount ?? 'none'"
    :data-docx-pagination-state="paginationState"
    :data-docx-pagination-iterations="paginationIterations"
    :data-docx-pagination-target-reached="paginationTargetReached ? 'true' : 'false'"
    :data-docx-view-mode="viewMode"
    :data-docx-fit-mode="fitMode ?? 'none'"
    :style="rootStyle"
    @scroll.passive="onScroll"
    @dragenter="onDragEnter"
    @dragover="onDragOver"
    @dragleave="onDragLeave"
    @drop="onDrop"
    @contextmenu.prevent="onContextMenu"
    @mouseup="onMouseUp"
    @copy="onCopy"
    @cut="onCut"
  >
    <div :style="spacerStyle">
      <div
        v-if="bookmarkNotice && !isPrinting"
        class="docx-bookmark-notice"
        role="status"
        data-testid="docx-bookmark-notice"
        :style="bookmarkNoticeStyle"
      >
        {{ bookmarkNotice }}
      </div>
      <div
        v-for="entry of visiblePageEntries"
        :key="`page-${entry.pageIndex}`"
        :data-docx-page-wrapper="true"
        :data-docx-page-index="entry.pageIndex"
        :data-docx-page-number="entry.pageNumber"
        :data-docx-page-row-index="entry.rowIndex"
        :data-docx-page-segmentation-key="entry.segmentationKey"
        :data-docx-page-revision="pageRenderRevision(entry.pageIndex)"
        :style="pageWrapperStyle(entry)"
      >
        <div :style="pageContentStyle(entry)">
          <DocxPageSurface
            :page-index="entry.pageIndex"
            :page-layout="entry.pageLayout"
            :page-node-segments="entry.pageNodeSegments"
            :page-columns="entry.columns"
            :column-layout="entry.columnLayout"
            :model="renderModel"
            :controller="controller"
            :editable="editable"
            :zoom-factor="zoomFactor"
            :page-width-px="entry.pageWidthPx"
            :page-content-width-px="entry.pageContentWidthPx"
            :page-number="entry.pageNumber"
            :total-pages="pageCount"
            :section-pages="entry.sectionPages"
            :page-number-format="entry.pageNumberFormat"
            :page-label="entry.pageLabel"
            :parity-blank-page="entry.parityBlankPage"
            :header-section="entry.headerSection"
            :footer-section="entry.footerSection"
            :section-index="entry.sectionIndex"
            :header-reference-type="entry.headerReferenceType"
            :footer-reference-type="entry.footerReferenceType"
            :tracked-changes-enabled="trackedChangesEnabled"
            :revision-display="renderRevisionDisplay"
            :comments-enabled="commentsEnabled"
            :tracked-changes="entry.trackedChanges"
            :comments="entry.comments"
            :footnotes="entry.footnotes"
            :endnotes="entry.endnotes"
            :footnote-labels="entry.noteLabels.footnote"
            :endnote-labels="entry.noteLabels.endnote"
            :note-marker-labels="entry.noteLabels"
            :page-background-color="documentBackgroundColor"
            :page-borders="entry.pageBorders ?? null"
            :theme="theme"
            :annotation-drawer-mode="annotationDrawerMode"
            :on-internal-link-click="revealBookmark"
            :bookmark-target="bookmarkTarget"
            :render-comment-card="renderCommentCard"
            :render-tracked-change-card="renderTrackedChangeCard"
            @measure="onPageMeasure(entry.pageIndex, $event)"
            @paragraph-measure="onParagraphMeasure"
            @table-row-measure="onTableRowMeasure"
            @content-measure="onPageContentMeasure(entry.pageIndex, $event)"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, onMounted, onUnmounted, watch, watchEffect } from "vue"
import { createTouchGestureTracker } from "@arcships/office-runtime/touch-gestures"
import { findDocxSearchTargetPage, type DocxSearchPageTarget } from "../composables/docx-search-page"
import type {
  DocModel,
  DocxComment,
  DocxCommentCardRenderProps,
  DocxDocumentTheme,
  DocxEditorController,
  DocxRevisionDisplay,
  DocxTrackedChange,
  DocxTrackedChangeCardRenderProps,
  DocumentNoteDefinition,
  DocumentPageNodeSegment,
  DocumentPageNotes,
  DocumentPageBorders,
  FooterSection,
  HeaderSection,
  LayoutOptions,
  LayoutSnapshotPage,
} from "@arcships/docx-core"
import {
  buildLayoutSnapshot,
  buildDocumentPageNotes,
  ensureDocumentFontInvalidation,
  collectCommentsFromModel,
  collectTrackedChangesFromModel,
  parseSectionLayout,
  parseSectionPageBorders,
  createSectionPageNumberTracker,
  primeCommentAnchors,
  getFontMetricsRevision,
  subscribeFontMetrics,
  resolveDocumentBackgroundColor,
  reviewDisplayModel,
  resolveDocumentSectionsFromMetadata,
  resolveDocumentLayout,
  resolveSectionVariantReferenceTypeForPage,
  groupGutterAnnotationsByPage,
  trackedChangeDuplicatesInlineMarkup,
  selectSectionVariantForPage,
  documentPageNodeSegmentsIdentityKey,
  PaginationMeasurementLoop,
  resolvePaginationPageCountTarget,
  resolvePaginationScrollAnchor,
  resolvePaginationScrollTop,
  reuseUnchangedPaginationItems,
  createDocxBookmarkIndex,
  type DocxBookmarkTarget,
  type PaginationMeasurementFeedback,
  type PaginationScrollAnchor,
  DEFAULT_DOC_PAGE_WIDTH,
  DEFAULT_DOC_PAGE_HEIGHT,
} from "@arcships/docx-core"
import DocxPageSurface from "./DocxPageSurface.vue"
import { useDocxMessages } from "../i18n"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useFileDrop } from "@arcships/vue-ui"
import type {
  DocxBookmarkRevealResult,
  DocxBookmarkTargetHighlight,
} from "../composables/docx-bookmarks"
import {
  DOCX_FILE_EXTENSIONS,
  DOCX_FILE_MIME_TYPES,
} from "../composables/docx-file-types"
import {
  ensureDocxViewerPageSurfaceRegistry,
  notifyDocxViewerPageSurfaceListeners,
} from "../composables/page-surface-registry"
import {
  annotationGutterReservedWidthPx,
  annotationGutterUsesDrawer,
} from "./gutter-metrics"
import {
  docxPageAtPoint,
  docxRowIndexAtOffset,
  docxRowsInViewport,
  docxSnapTopForRow,
  normalizeDocxFitMode,
  normalizeDocxViewMode,
  resolveDocxViewLayout,
  type DocxFitMode,
  type DocxPageRowLayout,
  type DocxViewMode,
} from "../composables/docx-view-mode"

// ── Constants ──────────────────────────────────────────────────────
const DOC_PAGE_BREAK_GAP = 28
const PAGE_OVERSCAN = 2
const SCROLL_MEASUREMENT_DEBOUNCE_MS = 80
/** Horizontal gap between the two pages of a spread. */
const SPREAD_COLUMN_GAP = 24
/** Ignore sub-pixel drift so single-page snapping cannot fight smooth scrolling. */
const SINGLE_PAGE_SNAP_TOLERANCE_PX = 8
const PAGE_SIDE_PADDING = 48
// Shared empty list keeps hidden annotations allocation-free.
const EMPTY_ANNOTATIONS: readonly never[] = []
// The scroll container uses `padding: 24px 0`; page offsets are expressed in
// scroll coordinates so scrollToPage / hit tests line up with the rendered page.
const PAGE_TOP_PADDING = 24

// ── Types ──────────────────────────────────────────────────────────
export interface DocxViewerRootProps {
  model: DocModel
  controller?: DocxEditorController
  editable?: boolean
  layoutOptions?: LayoutOptions
  theme?: DocxDocumentTheme
  zoomScale?: number
  /**
   * Identity persisted on `data-docx-document-id`; block/word geometry refuses
   * fragments whose source reference belongs to another document.
   */
  documentId?: string
  /** When true, zoom auto-fits page width to container. Overrides zoomScale. */
  fitWidth?: boolean
  /**
   * Fit the paper to the container instead of using `zoomScale`: `width` fits
   * the page edge horizontally, `page` fits the whole page (width or height,
   * whichever is tighter). `null`/`none` keeps the zoom factor.
   */
  fitMode?: DocxFitMode | null
  /**
   * Page layout: continuous scrolling, one page at a time, or two-page spreads.
   * Every mode keeps the page offsets, so page navigation and scroll sync work
   * the same way and switching modes can stay on the current page.
   */
  viewMode?: DocxViewMode
  /** `null` follows the controller flag; `false` always hides. */
  showTrackedChanges?: boolean | null
  /**
   * Review display mode. `null` follows the controller and falls back to
   * `balloons`. Display only: the source model and export keep every revision.
   */
  revisionDisplay?: DocxRevisionDisplay | null
  showComments?: boolean | null
  /** Host renderer that replaces every comment card in the gutter. */
  renderCommentCard?: (props: DocxCommentCardRenderProps) => unknown
  /** Host renderer that replaces every tracked-change card in the gutter. */
  renderTrackedChangeCard?: (
    props: DocxTrackedChangeCardRenderProps
  ) => unknown
}

interface VisiblePageEntry {
  pageIndex: number
  topPx: number
  heightPx: number
  pageLayout: ReturnType<typeof resolveDocumentLayout>
  pageNodeSegments: DocumentPageNodeSegment[]
  columns: DocumentPageNodeSegment[][]
  columnLayout: { count: number; gapPx: number; widthsPx: number[] }
  pageWidthPx: number
  pageContentWidthPx: number
  pageNumber: number
  /** Row that renders this page; a spread row owns two pages. */
  rowIndex: number
  pageNumberFormat?: string
  /** Label Word shows for the page (`iv`, `12`); `w:fmt="none"` → empty. */
  pageLabel: string
  /** True when the page only exists so a section starts on an even/odd page. */
  parityBlankPage: boolean
  sectionPages: number
  sectionIndex: number
  /** Header/footer `w:type` variant this page looks for. */
  headerReferenceType: string
  footerReferenceType: string
  headerSection?: HeaderSection
  footerSection?: FooterSection
  footnotes: DocumentNoteDefinition[]
  endnotes: DocumentNoteDefinition[]
  noteLabels: { footnote: Map<number, string>; endnote: Map<number, string> }
  pageBorders?: DocumentPageBorders
  trackedChanges: readonly DocxTrackedChange[]
  comments: readonly DocxComment[]
  /** Identity of this page's segmentation; unchanged pages keep their DOM. */
  segmentationKey: string
}

interface PageMeasureEvent {
  heightPx: number
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<DocxViewerRootProps>(),
  {
    editable: false,
    zoomScale: 100,
    showTrackedChanges: null,
    revisionDisplay: null,
    showComments: null,
    fitMode: null,
    viewMode: "continuous",
  }
)

const emit = defineEmits<{
  pageCountChange: [count: number]
  visiblePageRange: [range: { startPageIndex: number; endPageIndex: number }]
  contextMenu: [ctx: { pageIndex: number; clientX: number; clientY: number; containerX: number; containerY: number; target?: EventTarget | null }]
  selectionChange: [sel: { kind: string; text?: string; pageIndex?: number }]
  /** A bookmark reveal needs `pageIndex` mounted/visible in the host. */
  requestPageReveal: [pageIndex: number]
  /** An internal link named a bookmark this document does not contain. */
  bookmarkMissing: [name: string]
}>()

// ── State ──────────────────────────────────────────────────────────
const m = useDocxMessages()
const scrollContainerRef = ref<HTMLElement | undefined>()
const scrollTop = ref(0)
const viewportHeight = ref(800)
const containerWidth = ref(800)

// Track measured page heights (key: pageIndex). Each entry is tagged with the
// page layout it was measured for so layout changes invalidate stale values.
const measuredPageHeights = ref<Record<number, { heightPx: number; layoutKey: string }>>({})

// Measure→repaginate feedback (#99). Paragraph heights come from rendered
// paragraph hosts; table rows come from table hosts; page content heights come
// from the page body. A bounded convergence gate decides whether another pass
// is allowed, and resets whenever the document/layout changes.
const PAGINATION_MEASUREMENT_MAX_ITERATIONS = 3
const PAGINATION_MEASUREMENT_TOLERANCE_PX = 1
const paginationMeasurementFeedback = ref<PaginationMeasurementFeedback | undefined>()
const paginationPageCountTarget = computed(() =>
  resolvePaginationPageCountTarget(props.model.metadata.documentPageCount)
)
const paginationState = ref<"idle" | "repaginating" | "converged" | "oscillating" | "exhausted">("idle")
const paginationIterations = ref(0)
const paginationTargetReached = ref(false)
// A web font that arrives after the first measure changes advance widths; the
// revision makes every pagination/measurement computed depend on that event.
const fontMetricsRevision = ref(getFontMetricsRevision())
const paginationFontMetricsRevision = computed(() => fontMetricsRevision.value)
let stopFontMetricsSubscription: (() => void) | undefined
const measuredParagraphHeightsPxByNodeIndex = ref<Record<number, number>>({})
const measuredTableRowHeightsPxByNodeIndex = ref<Record<number, number[]>>({})
const measuredPageContentHeights = ref<
  Record<number, { heightPx: number; identityKey: string }>
>({})

function createPaginationMeasurementLoop(): PaginationMeasurementLoop<PaginationMeasurementFeedback> {
  return new PaginationMeasurementLoop<PaginationMeasurementFeedback>({
    maxIterations: PAGINATION_MEASUREMENT_MAX_ITERATIONS,
    tolerancePx: PAGINATION_MEASUREMENT_TOLERANCE_PX,
    targetPageCount: paginationPageCountTarget.value.targetPageCount,
    // Virtualization only measures pages currently in the DOM. Pagination
    // stability is therefore decided by page identities plus feedback, while
    // rendered page heights remain part of the oscillation signature.
    requireStablePageHeights: false,
  })
}

let paginationMeasurementLoop = createPaginationMeasurementLoop()
let measurementFlushScheduled = false
let measurementFlushHandle: ReturnType<typeof setTimeout> | undefined
// Captured before a feedback pass changes page offsets; restored afterwards so
// the reader's viewport keeps the same page position.
let pendingPaginationScrollAnchor: PaginationScrollAnchor | undefined

// ── Drag-over state for file drops ──────────────────────────────────
// The viewer root is the drop target; the shared composable owns the accept
// list, the nesting depth and the outline reset. Imports through the editor
// controller so the dropped file follows the same path as the picker.
const fileDrop = useFileDrop({
  accept: {
    extensions: DOCX_FILE_EXTENSIONS,
    mimeTypes: DOCX_FILE_MIME_TYPES,
  },
  enabled: () => props.editable && !!props.controller,
  onFile: (file) => props.controller?.importDocxFile(file),
  onReject: () => props.controller?.setStatus(m.value.docx.viewer.unsupportedFile),
})
const { isDragOver, onDragEnter, onDragOver, onDragLeave, onDrop } = fileDrop

// ── Print mode ─────────────────────────────────────────────────────
// Printing must include every page, not only the virtualized window that is
// currently mounted. `matchMedia("print")` covers browser print preview and
// automated media emulation; beforeprint/afterprint cover browsers that do not
// update the media query list.
const printMediaActive = ref(false)
const beforePrintActive = ref(false)
const isPrinting = computed(() => printMediaActive.value || beforePrintActive.value)
let printMediaQuery: MediaQueryList | undefined
let printPageStyle: HTMLStyleElement | undefined

function syncPrintPageSize(widthPx?: number, heightPx?: number): void {
  if (typeof document === "undefined" || typeof document.createElement !== "function") return
  if (!Number.isFinite(widthPx) || !Number.isFinite(heightPx) || (widthPx ?? 0) <= 0 || (heightPx ?? 0) <= 0) {
    printPageStyle?.remove?.()
    printPageStyle = undefined
    return
  }
  if (!printPageStyle) {
    printPageStyle = document.createElement("style")
    if (printPageStyle.dataset) printPageStyle.dataset.officePrintPage = "docx"
    document.head?.append(printPageStyle)
  }
  printPageStyle.textContent = `@page { size: ${Math.round(widthPx!)}px ${Math.round(heightPx!)}px; margin: 0; }`
}

function onPrintMediaChange(event: MediaQueryListEvent): void {
  printMediaActive.value = event.matches
}

function onBeforePrint(): void {
  beforePrintActive.value = true
}

function onAfterPrint(): void {
  beforePrintActive.value = false
  printMediaActive.value = printMediaQuery?.matches ?? false
}

function setupPrintDetection(): () => void {
  if (typeof window === "undefined") return () => {}
  printMediaQuery = typeof window.matchMedia === "function" ? window.matchMedia("print") : undefined
  if (printMediaQuery) {
    printMediaActive.value = printMediaQuery.matches
    if (typeof printMediaQuery.addEventListener === "function") {
      printMediaQuery.addEventListener("change", onPrintMediaChange)
    } else {
      printMediaQuery.addListener?.(onPrintMediaChange)
    }
  }
  if (typeof window.addEventListener === "function") {
    window.addEventListener("beforeprint", onBeforePrint)
    window.addEventListener("afterprint", onAfterPrint)
  }
  return () => {
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
  }
}

// ── Pagination ─────────────────────────────────────────────────────
// Review display is a view choice: `renderModel` drops deleted runs in
// balloons/none, while edits, comments and export keep reading `props.model`.
const trackedChangesVisible = computed(
  () => props.showTrackedChanges ?? props.controller?.showTrackedChanges ?? false
)
const effectiveRevisionDisplay = computed<DocxRevisionDisplay>(
  () => props.revisionDisplay ?? props.controller?.revisionDisplay ?? "balloons"
)
// Hiding revisions means the final text, regardless of the preferred markup
// mode; the selector keeps its value for when revisions are shown again.
const renderRevisionDisplay = computed<DocxRevisionDisplay>(() =>
  trackedChangesVisible.value ? effectiveRevisionDisplay.value : "none"
)
const renderModel = computed(() =>
  reviewDisplayModel(props.model, renderRevisionDisplay.value)
)
const baseDocumentLayout = computed(() => resolveDocumentLayout(renderModel.value))
watchEffect(() => {
  const layout = baseDocumentLayout.value
  syncPrintPageSize(layout?.pageWidthPx, layout?.pageHeightPx)
})
const documentBackgroundColor = computed(() =>
  resolveDocumentBackgroundColor(props.model.metadata.documentBackgroundColor)
)

const resolvedLayoutOptions = computed<LayoutOptions>(() => ({
  pageWidth: baseDocumentLayout.value.pageWidthPx,
  pageHeight: baseDocumentLayout.value.pageHeightPx,
  margin: baseDocumentLayout.value.marginsPx.top,
  ...props.layoutOptions,
}))

interface PaginationSegmentCache {
  segmentationKeys: string[]
  pageNodeSegments: DocumentPageNodeSegment[][]
  pageColumns: DocumentPageNodeSegment[][][]
}

let paginationSegmentCache: PaginationSegmentCache | undefined

function pageSegmentationKey(
  page: LayoutSnapshotPage,
  segments: DocumentPageNodeSegment[],
  columns: DocumentPageNodeSegment[][]
): string {
  return [
    page.index,
    page.columnCount,
    page.columnGapPx,
    (page.columnWidthsPx ?? []).join(","),
    documentPageNodeSegmentsIdentityKey(segments),
    columns
      .map((columnSegments) =>
        documentPageNodeSegmentsIdentityKey(columnSegments)
      )
      .join("~"),
  ].join("|")
}

const paginationPlan = computed(() => {
  // Track the font-metric revision: a late web font changes measured
  // advances and must re-run segmentation, not just repaint.
  void paginationFontMetricsRevision.value
  const snapshot = buildLayoutSnapshot(renderModel.value, {
    ...resolvedLayoutOptions.value,
    measurementFeedback: paginationMeasurementFeedback.value,
  })

  const segmentFromBlock = (block: (typeof snapshot.pages)[number]["blocks"][number]): DocumentPageNodeSegment | undefined => {
    const nodeIndex = block.source?.nodeIndex
    if (!Number.isFinite(nodeIndex) || (nodeIndex as number) < 0) return undefined

    const segment: DocumentPageNodeSegment = {
      nodeIndex: Math.round(nodeIndex as number),
    }
    if (block.kind === "table" && block.tableRowRange) {
      segment.tableRowRange = {
        startRowIndex: block.tableRowRange.startRowIndex,
        endRowIndex: block.tableRowRange.endRowIndex,
      }
    }
    if (block.kind === "paragraph" && block.paragraphLineRange) {
      segment.paragraphLineRange = {
        startLineIndex: block.paragraphLineRange.startLineIndex,
        endLineIndex: block.paragraphLineRange.endLineIndex,
        totalLineCount: block.paragraphLineRange.totalLineCount,
        lineHeightPx: block.paragraphLineRange.lineHeightPx,
      }
    }
    return segment
  }

  const rawPageNodeSegments: DocumentPageNodeSegment[][] = snapshot.pages.map((page) =>
    page.blocks.flatMap((block): DocumentPageNodeSegment[] => {
      const segment = segmentFromBlock(block)
      return segment ? [segment] : []
    })
  )

  const rawPageColumns: DocumentPageNodeSegment[][][] = snapshot.pages.map((page) => {
    const columnCount = Math.max(1, page.columnCount ?? 1)
    const columns = Array.from(
      { length: columnCount },
      () => [] as DocumentPageNodeSegment[]
    )
    page.blocks.forEach((block) => {
      const segment = segmentFromBlock(block)
      if (!segment) return
      const columnIndex = Math.max(
        0,
        Math.min(columnCount - 1, block.source?.columnIndex ?? 0)
      )
      columns[columnIndex].push(segment)
    })
    return columns
  })

  // Only pages whose segmentation actually changed get new arrays. Vue can then
  // skip re-rendering the other page surfaces when measured heights are fed
  // back into the segmenter.
  const segmentationKeys = snapshot.pages.map((page, pageIndex) =>
    pageSegmentationKey(
      page,
      rawPageNodeSegments[pageIndex] ?? [],
      rawPageColumns[pageIndex] ?? []
    )
  )
  const nextSegments = reuseUnchangedPaginationItems({
    previous: paginationSegmentCache?.pageNodeSegments,
    next: rawPageNodeSegments,
    isUnchanged: (_previous, _next, pageIndex) =>
      paginationSegmentCache?.segmentationKeys[pageIndex] ===
      segmentationKeys[pageIndex],
  })
  const nextColumns = reuseUnchangedPaginationItems({
    previous: paginationSegmentCache?.pageColumns,
    next: rawPageColumns,
    isUnchanged: (_previous, _next, pageIndex) =>
      paginationSegmentCache?.segmentationKeys[pageIndex] ===
      segmentationKeys[pageIndex],
  })

  paginationSegmentCache = {
    segmentationKeys,
    pageNodeSegments: nextSegments.items,
    pageColumns: nextColumns.items,
  }

  return {
    pages: snapshot.pages,
    pageNodeSegments: nextSegments.items,
    pageColumns: nextColumns.items,
    segmentationKeys,
    changedPageIndexes: nextSegments.changedIndexes,
    snapshot,
  }
})

const pages = computed(() => paginationPlan.value.pages)
const pageNodeSegments = computed(() => paginationPlan.value.pageNodeSegments)
const pageColumns = computed(() => paginationPlan.value.pageColumns)
const pageCount = computed(() => pages.value.length)
const FIT_ZOOM_MIN = 0.25
const FIT_ZOOM_MAX = 2
const FIT_ZOOM_SIDE_PAD = 48
/** Scroll-container padding above and below the paper, plus a small margin. */
const FIT_ZOOM_VERTICAL_PAD = 64

// Narrow viewports get the drawer overlay instead of a column beside the paper,
// so the paper can keep the full container width.
// The display model drops deleted runs from the text flow in balloons/none
// mode. Pagination, rendering and note bodies read it; edits, comments and
// export keep reading the source model.
const trackedChangeGutterVisible = computed(
  () =>
    trackedChangesVisible.value &&
    effectiveRevisionDisplay.value === "balloons"
)
const annotationsVisible = computed(
  () =>
    trackedChangeGutterVisible.value ||
    (props.showComments ?? props.controller?.showComments ?? false)
)

const annotationDrawerMode = computed(
  () =>
    annotationsVisible.value &&
    annotationGutterUsesDrawer(containerWidth.value)
)

// Horizontal space the annotation gutter reserves beside the paper, in screen
// px. The pane counter-scales itself (see gutter-metrics), so this footprint is
// constant at every zoom level. 0 when annotations are hidden or when the pane
// is an overlay drawer (it then takes no space from the paper).
const gutterScreenWidthPx = computed(() =>
  annotationsVisible.value && !annotationDrawerMode.value
    ? annotationGutterReservedWidthPx()
    : 0
)

const fitZoom = computed(() => {
  if (!baseDocumentLayout.value) return 1
  const pageWidth = baseDocumentLayout.value.pageWidthPx
  if (!pageWidth || pageWidth <= 0) return 1
  // The annotation gutter is part of the horizontal footprint; reserve its
  // screen-space width so fit-width does not introduce a horizontal scrollbar
  // when it is visible.
  const available = Math.max(
    200,
    containerWidth.value - FIT_ZOOM_SIDE_PAD - gutterScreenWidthPx.value
  )
  return Math.min(FIT_ZOOM_MAX, Math.max(FIT_ZOOM_MIN, available / pageWidth))
})

// "Fit page" keeps the whole sheet in view: the tighter of the width and the
// height fit, so a landscape section still fits a portrait viewport.
const fitPageZoom = computed(() => {
  const pageHeight = baseDocumentLayout.value?.pageHeightPx
  if (!pageHeight || pageHeight <= 0) return fitZoom.value
  const availableHeight = Math.max(200, viewportHeight.value - FIT_ZOOM_VERTICAL_PAD)
  return Math.min(
    FIT_ZOOM_MAX,
    Math.max(FIT_ZOOM_MIN, Math.min(fitZoom.value, availableHeight / pageHeight))
  )
})

const zoomFactor = computed(() => {
  if (fitMode.value === "page") return fitPageZoom.value
  if (fitMode.value === "width") return fitZoom.value
  if (props.fitWidth) return fitZoom.value
  return Math.min(2, Math.max(0.5, props.zoomScale / 100))
})

watch(pageCount, (count) => emit("pageCountChange", count), { immediate: true })

// ── Document layout ────────────────────────────────────────────────
// Per-section page numbering: `w:pgNumType` restart/format, `SECTIONPAGES` and
// the parity blank page a `w:type="oddPage"/"evenPage"` break inserts. The
// tracker owns the page → section attribution too, so the page numbers, the
// section headers/footers and the thumbnail labels cannot drift apart.
const sectionPageNumbers = computed(() => {
  const documentSectionPropertiesXml =
    renderModel.value.metadata.sectionPropertiesXml
  return createSectionPageNumberTracker({
    sections: resolveDocumentSectionsFromMetadata(
      renderModel.value.metadata
    ).map((section) => ({
        startNodeIndex: section.startNodeIndex,
        // A section without its own `w:sectPr` inherits the document's.
        sectionPropertiesXml:
          section.sectionPropertiesXml ?? documentSectionPropertiesXml,
      })
    ),
    pages: pageNodeSegments.value.map((segments) => ({
      startNodeIndex: segments[0]?.nodeIndex,
    })),
    nodeCount: props.model.nodes.length,
  })
})

const pagePresentations = computed(() => {
  const sections = resolveDocumentSectionsFromMetadata(renderModel.value.metadata)
  const firstPageBySection = new Map<number, number>()
  const evenAndOddHeaders =
    renderModel.value.metadata.compatibility?.evenAndOddHeaders ?? false
  const tracker = sectionPageNumbers.value

  return pageNodeSegments.value.map((segments, pageIndex) => {
    const sectionIndex = tracker.sectionIndexForPage(pageIndex) ?? 0
    if (!firstPageBySection.has(sectionIndex)) {
      firstPageBySection.set(sectionIndex, pageIndex)
    }

    const section = sections[sectionIndex]
    const sectionPageIndex = pageIndex - (firstPageBySection.get(sectionIndex) ?? pageIndex)
    const sectionPropertiesXml =
      section?.sectionPropertiesXml ??
      renderModel.value.metadata.sectionPropertiesXml
    const pageLayout = sectionPropertiesXml
      ? parseSectionLayout(sectionPropertiesXml)
      : baseDocumentLayout.value
    const variantOptions = { evenAndOddHeaders, documentPageIndex: pageIndex }
    // The variant these pages look for, so an authoring command can create
    // exactly that header/footer instead of always the "default" one.
    const variantReferenceType = resolveSectionVariantReferenceTypeForPage({
      sectionPropertiesXml,
      sectionPageIndex,
      documentPageIndex: pageIndex,
      evenAndOddHeaders,
    })
    const pageNumber = tracker.pageNumberForPage(pageIndex) ?? pageIndex + 1

    return {
      sectionIndex,
      pageLayout,
      pageNumber,
      pageLabel: tracker.pageLabelForPage(pageIndex) ?? String(pageNumber),
      parityBlankPage: tracker.isParityBlankPage(pageIndex),
      headerReferenceType: variantReferenceType,
      footerReferenceType: variantReferenceType,
      pageNumberFormat: tracker.pageNumberFormatForPage(pageIndex),
      pageBorders: parseSectionPageBorders(sectionPropertiesXml),
      columnLayout: {
        count: pages.value[pageIndex]?.columnCount ?? 1,
        gapPx: pages.value[pageIndex]?.columnGapPx ?? 0,
        widthsPx: pages.value[pageIndex]?.columnWidthsPx ?? [],
      },
      segments: pageNodeSegments.value[pageIndex] ?? [],
      columns: pageColumns.value[pageIndex] ?? [],
      headerSection: section
        ? selectSectionVariantForPage(
            section.headerSections,
            sectionPropertiesXml,
            sectionPageIndex,
            variantOptions
          )
        : undefined,
      footerSection: section
        ? selectSectionVariantForPage(
            section.footerSections,
            sectionPropertiesXml,
            sectionPageIndex,
            variantOptions
          )
        : undefined,
    }
  })
})

// The thumbnail rail reads the viewer's pagination result instead of running
// its own buildLayoutSnapshot on every model change.
if (props.controller) {
  watchEffect(() => {
    const controller = props.controller
    if (!controller) return
    const plan = paginationPlan.value
    const presentations = pagePresentations.value
    const registry = ensureDocxViewerPageSurfaceRegistry(controller)
    const sizes = new Map<number, { widthPx: number; heightPx: number }>()
    plan.pages.forEach((page, pageIndex) => {
      const layout = presentations[pageIndex]?.pageLayout ?? baseDocumentLayout.value
      sizes.set(pageIndex, {
        widthPx: layout.pageWidthPx ?? page.pageSizePx.width,
        heightPx: layout.pageHeightPx ?? page.pageSizePx.height,
      })
    })
    registry.layoutPageCount = plan.pages.length
    registry.layoutPageSizes = sizes
    registry.layoutSnapshot = plan.snapshot
    // The rail labels pages with the same per-section numbers the pages show.
    registry.pageNumberTracker = sectionPageNumbers.value
    notifyDocxViewerPageSurfaceListeners(controller)
  }, { flush: "post" })
}

// ── Theme / feature flags ──────────────────────────────────────────
const theme = computed(() => props.theme ?? props.controller?.documentTheme ?? "light")
// Boolean props default to `false`, so `null` is what marks "not passed":
// without it the controller fallback below could never apply.
const trackedChangesEnabled = computed(
  () => trackedChangesVisible.value
)
const commentsEnabled = computed(
  () => props.showComments ?? props.controller?.showComments ?? false
)
// Visual horizontal space the annotation gutter occupies at the current zoom.
// Visual horizontal space the annotation gutter occupies in screen px.
// 0 when annotations are hidden, so no space is reserved. The pane counter-
// scales, so this footprint does not follow the zoom factor.
const gutterVisualWidthPx = computed(() => isPrinting.value ? 0 : gutterScreenWidthPx.value)
const trackedChanges = computed<readonly DocxTrackedChange[]>(() => {
  // `showTrackedChanges` gates the collection even when a controller is
  // present: hidden revisions must not cost a model scan.
  if (!trackedChangesEnabled.value) return EMPTY_ANNOTATIONS
  const changes = props.controller?.trackedChanges ?? collectTrackedChangesFromModel(props.model)
  // Insertions and move-to content are already highlighted inline, so a gutter
  // card would show the same revision twice.
  return changes.filter((change) => !trackedChangeDuplicatesInlineMarkup(change))
})
const comments = computed<readonly DocxComment[]>(() => {
  if (!commentsEnabled.value) return EMPTY_ANNOTATIONS
  return props.controller?.comments ?? collectCommentsFromModel(props.model)
})

// Each annotation is owned by exactly one page: a change inside row 3 of a
// table spanning pages — or inside a paragraph split across pages — must not be
// repeated on every page that renders the node.
const trackedChangesByPage = computed<readonly (readonly DocxTrackedChange[])[]>(() =>
  groupGutterAnnotationsByPage(trackedChanges.value, pageNodeSegments.value)
)
const commentsByPage = computed<readonly (readonly DocxComment[])[]>(() =>
  groupGutterAnnotationsByPage(comments.value, pageNodeSegments.value)
)

function pageTrackedChanges(pageIndex: number): readonly DocxTrackedChange[] {
  return trackedChangesByPage.value[pageIndex] ?? EMPTY_ANNOTATIONS
}

function pageComments(pageIndex: number): readonly DocxComment[] {
  return commentsByPage.value[pageIndex] ?? EMPTY_ANNOTATIONS
}

// Capture comment anchors before any edit clears paragraph sourceXml, even
// while the comments gutter is hidden.
watch(() => props.model, (model) => primeCommentAnchors(model), { immediate: true })

// Notes are resolved from the run model, once per document: this is the single
// source of truth shared by in-text markers and the note bodies at the page
// bottom (see buildDocumentPageNotes).
const EMPTY_PAGE_NOTES: DocumentPageNotes = {
  footnotes: [],
  endnotes: [],
  footnoteNumbers: new Map(),
  endnoteNumbers: new Map(),
  footnoteLabels: new Map(),
  endnoteLabels: new Map(),
}

const pageNotes = computed<DocumentPageNotes[]>(() =>
  buildDocumentPageNotes(renderModel.value, pageNodeSegments.value)
)

function notesForPage(pageIndex: number): DocumentPageNotes {
  return pageNotes.value[pageIndex] ?? EMPTY_PAGE_NOTES
}

function pageFootnoteNotes(pageIndex: number): DocumentNoteDefinition[] {
  return notesForPage(pageIndex).footnotes.map((entry) => entry.note)
}

function pageEndnoteNotes(pageIndex: number): DocumentNoteDefinition[] {
  return notesForPage(pageIndex).endnotes.map((entry) => entry.note)
}

// ── Page heights ───────────────────────────────────────────────────
function pageLayoutKey(pageIndex: number): string {
  const layout = pagePresentations.value[pageIndex]?.pageLayout
  return layout ? `${layout.pageWidthPx}x${layout.pageHeightPx}` : "default"
}

function getPageHeight(pageIndex: number): number {
  const measured = measuredPageHeights.value[pageIndex]
  if (measured && measured.layoutKey === pageLayoutKey(pageIndex) && Number.isFinite(measured.heightPx) && measured.heightPx > 0) {
    return measured.heightPx
  }
  return pagePresentations.value[pageIndex]?.pageLayout.pageHeightPx ?? DEFAULT_DOC_PAGE_HEIGHT
}

// ── Page offsets ───────────────────────────────────────────────────
const viewMode = computed<DocxViewMode>(() => normalizeDocxViewMode(props.viewMode))
const fitMode = computed<DocxFitMode | null>(() => normalizeDocxFitMode(props.fitMode))

// One geometry for rendering, hit testing, snapping and page navigation. The
// continuous mode is one page per row, so the row math reduces to the page
// stack and every existing offset consumer keeps working unchanged.
const viewLayout = computed(() =>
  resolveDocxViewLayout({
    pageCount: pageCount.value,
    mode: viewMode.value,
    pageSize: (pageIndex) => ({
      widthPx: pagePresentations.value[pageIndex]?.pageLayout.pageWidthPx ?? DEFAULT_DOC_PAGE_WIDTH,
      heightPx: getPageHeight(pageIndex),
    }),
    zoom: zoomFactor.value,
    gapPx: DOC_PAGE_BREAK_GAP,
    topPaddingPx: PAGE_TOP_PADDING,
    spreadGapPx: SPREAD_COLUMN_GAP,
    gutterWidthPx: gutterVisualWidthPx.value,
    print: isPrinting.value,
  })
)

const pageOffsets = computed(() => viewLayout.value.pageOffsets)
const totalHeightPx = computed(() => viewLayout.value.totalHeightPx)

/** Row layout of a page, used by the wrapper styles and hit testing. */
function rowForPage(pageIndex: number): DocxPageRowLayout | undefined {
  const rowIndex = viewLayout.value.pageRows[pageIndex]
  return rowIndex === undefined ? undefined : viewLayout.value.rows[rowIndex]
}

// ── Visible page calculation ───────────────────────────────────────
interface PageRenderCache {
  entries: Map<number, VisiblePageEntry>
  revisions: Map<number, number>
}

const pageRenderCache: PageRenderCache = {
  entries: new Map(),
  revisions: new Map(),
}

function pageRenderRevision(pageIndex: number): number {
  return pageRenderCache.revisions.get(pageIndex) ?? 1
}

function sameNumberList(
  previous: readonly number[] | undefined,
  next: readonly number[] | undefined
): boolean {
  if (previous === next) return true
  if (!previous || !next || previous.length !== next.length) return false
  return previous.every((value, index) => value === next[index])
}

function samePageSegments(
  previous: readonly DocumentPageNodeSegment[],
  next: readonly DocumentPageNodeSegment[]
): boolean {
  if (previous === next) return true
  if (previous.length !== next.length) return false
  return (
    documentPageNodeSegmentsIdentityKey([...previous]) ===
    documentPageNodeSegmentsIdentityKey([...next])
  )
}

function samePageColumns(
  previous: readonly DocumentPageNodeSegment[][],
  next: readonly DocumentPageNodeSegment[][]
): boolean {
  if (previous === next) return true
  if (previous.length !== next.length) return false
  return previous.every((column, index) =>
    samePageSegments(column, next[index] ?? [])
  )
}

function samePageLayout(
  previous: ReturnType<typeof resolveDocumentLayout>,
  next: ReturnType<typeof resolveDocumentLayout>
): boolean {
  if (previous === next) return true
  if (!previous || !next) return false
  return (
    previous.pageWidthPx === next.pageWidthPx &&
    previous.pageHeightPx === next.pageHeightPx &&
    previous.marginsPx.top === next.marginsPx.top &&
    previous.marginsPx.right === next.marginsPx.right &&
    previous.marginsPx.bottom === next.marginsPx.bottom &&
    previous.marginsPx.left === next.marginsPx.left &&
    previous.headerDistancePx === next.headerDistancePx &&
    previous.footerDistancePx === next.footerDistancePx &&
    previous.docGridLinePitchPx === next.docGridLinePitchPx
  )
}

function sameColumnLayout(
  previous: VisiblePageEntry["columnLayout"],
  next: VisiblePageEntry["columnLayout"]
): boolean {
  return (
    previous === next ||
    (previous.count === next.count &&
      previous.gapPx === next.gapPx &&
      sameNumberList(previous.widthsPx, next.widthsPx))
  )
}

function sameSection(
  previous: HeaderSection | FooterSection | undefined,
  next: HeaderSection | FooterSection | undefined
): boolean {
  if (previous === next) return true
  if (!previous || !next) return false
  return (
    previous.partName === next.partName &&
    previous.referenceType === next.referenceType &&
    previous.nodes === next.nodes
  )
}

function sameNoteDefinitions(
  previous: readonly DocumentNoteDefinition[],
  next: readonly DocumentNoteDefinition[]
): boolean {
  if (previous === next) return true
  if (previous.length !== next.length) return false
  return previous.every(
    (note, index) =>
      note.id === next[index]?.id &&
      note.text === next[index]?.text &&
      note.nodes === next[index]?.nodes
  )
}

function sameLabels(
  previous: Map<number, string>,
  next: Map<number, string>
): boolean {
  if (previous === next) return true
  if (previous.size !== next.size) return false
  for (const [key, value] of previous) {
    if (next.get(key) !== value) return false
  }
  return true
}

function sameAnnotations<T>(previous: readonly T[], next: readonly T[]): boolean {
  if (previous === next) return true
  if (previous.length !== next.length) return false
  return previous.every((item, index) => item === next[index])
}

function samePageBorders(
  previous: VisiblePageEntry["pageBorders"],
  next: VisiblePageEntry["pageBorders"]
): boolean {
  if (previous === next) return true
  if (!previous || !next) return false
  return JSON.stringify(previous) === JSON.stringify(next)
}

function sameVisiblePageContent(
  previous: VisiblePageEntry,
  next: VisiblePageEntry
): boolean {
  return (
    previous.pageIndex === next.pageIndex &&
    previous.pageWidthPx === next.pageWidthPx &&
    previous.pageContentWidthPx === next.pageContentWidthPx &&
    previous.pageNumber === next.pageNumber &&
    previous.pageNumberFormat === next.pageNumberFormat &&
    previous.pageLabel === next.pageLabel &&
    previous.parityBlankPage === next.parityBlankPage &&
    previous.sectionPages === next.sectionPages &&
    previous.segmentationKey === next.segmentationKey &&
    samePageLayout(previous.pageLayout, next.pageLayout) &&
    samePageSegments(previous.pageNodeSegments, next.pageNodeSegments) &&
    samePageColumns(previous.columns, next.columns) &&
    sameColumnLayout(previous.columnLayout, next.columnLayout) &&
    sameSection(previous.headerSection, next.headerSection) &&
    sameSection(previous.footerSection, next.footerSection) &&
    sameNoteDefinitions(previous.footnotes, next.footnotes) &&
    sameNoteDefinitions(previous.endnotes, next.endnotes) &&
    sameLabels(previous.noteLabels.footnote, next.noteLabels.footnote) &&
    sameLabels(previous.noteLabels.endnote, next.noteLabels.endnote) &&
    samePageBorders(previous.pageBorders, next.pageBorders) &&
    sameAnnotations(previous.trackedChanges, next.trackedChanges) &&
    sameAnnotations(previous.comments, next.comments)
  )
}

function buildVisiblePageEntry(
  pageIndex: number,
  topPx: number,
  heightPx: number,
  previous?: VisiblePageEntry
): VisiblePageEntry {
  const presentation = pagePresentations.value[pageIndex]
  const presentedLayout = presentation?.pageLayout ?? baseDocumentLayout.value
  const pageLayout =
    previous && samePageLayout(previous.pageLayout, presentedLayout)
      ? previous.pageLayout
      : presentedLayout
  const pageWidthPx = pageLayout.pageWidthPx ?? DEFAULT_DOC_PAGE_WIDTH
  const segments = pageNodeSegments.value[pageIndex] ?? []
  const columns = pageColumns.value[pageIndex] ?? []
  const presentedColumnLayout = presentation?.columnLayout ?? {
    count: 1,
    gapPx: 0,
    widthsPx: [],
  }
  const columnLayout =
    previous &&
    sameColumnLayout(previous.columnLayout, presentedColumnLayout)
      ? previous.columnLayout
      : presentedColumnLayout
  const footnotes = pageFootnoteNotes(pageIndex)
  const endnotes = pageEndnoteNotes(pageIndex)
  const notes = notesForPage(pageIndex)
  const noteLabels =
    previous &&
    sameLabels(previous.noteLabels.footnote, notes.footnoteLabels) &&
    sameLabels(previous.noteLabels.endnote, notes.endnoteLabels)
      ? previous.noteLabels
      : { footnote: notes.footnoteLabels, endnote: notes.endnoteLabels }
  const pageBorders = presentation?.pageBorders
  const trackedChanges = pageTrackedChanges(pageIndex)
  const comments = pageComments(pageIndex)

  return {
    pageIndex,
    topPx,
    heightPx,
    rowIndex: viewLayout.value.pageRows[pageIndex] ?? 0,
    pageLayout,
    pageNodeSegments: segments,
    columns,
    columnLayout,
    pageWidthPx,
    pageContentWidthPx: Math.max(
      0,
      pageWidthPx -
        (pageLayout.marginsPx.left ?? 72) -
        (pageLayout.marginsPx.right ?? 72)
    ),
    pageNumber: presentation?.pageNumber ?? pageIndex + 1,
    pageNumberFormat: presentation?.pageNumberFormat,
    pageLabel: presentation?.pageLabel ?? String(pageIndex + 1),
    parityBlankPage: presentation?.parityBlankPage ?? false,
    sectionPages:
      presentation === undefined
        ? pageCount.value
        : sectionPageNumbers.value.sectionPagesForPage(pageIndex) ?? pageCount.value,
    sectionIndex: presentation?.sectionIndex ?? 0,
    headerReferenceType: presentation?.headerReferenceType ?? "default",
    footerReferenceType: presentation?.footerReferenceType ?? "default",
    headerSection: presentation?.headerSection,
    footerSection: presentation?.footerSection,
    footnotes:
      previous && sameNoteDefinitions(previous.footnotes, footnotes)
        ? previous.footnotes
        : footnotes,
    endnotes:
      previous && sameNoteDefinitions(previous.endnotes, endnotes)
        ? previous.endnotes
        : endnotes,
    noteLabels,
    pageBorders:
      previous && samePageBorders(previous.pageBorders, pageBorders)
        ? previous.pageBorders
        : pageBorders,
    trackedChanges:
      previous && sameAnnotations(previous.trackedChanges, trackedChanges)
        ? previous.trackedChanges
        : trackedChanges,
    comments:
      previous && sameAnnotations(previous.comments, comments)
        ? previous.comments
        : comments,
    segmentationKey:
      paginationPlan.value.segmentationKeys[pageIndex] ??
      documentPageNodeSegmentsIdentityKey(segments),
  }
}

const visiblePageEntries = computed<VisiblePageEntry[]>(() => {
  if (pageCount.value === 0) {
    pageRenderCache.entries = new Map()
    return []
  }

  const layout = viewLayout.value
  const viewportStart = isPrinting.value
    ? Number.NEGATIVE_INFINITY
    : Math.max(0, scrollTop.value - DOC_PAGE_BREAK_GAP * PAGE_OVERSCAN)
  const viewportEnd = isPrinting.value
    ? Number.POSITIVE_INFINITY
    : viewportStart + viewportHeight.value + DOC_PAGE_BREAK_GAP * PAGE_OVERSCAN * 2

  // A spread mounts both of its pages together; single-page mode mounts only
  // the page the reader currently rests on, so no neighbouring page leaks into
  // the viewport while the scroll offset snaps to a page boundary.
  const singleRowIndex = layout.mode === "single" && !isPrinting.value
    ? docxRowIndexAtOffset(layout.rows, scrollTop.value, viewportHeight.value)
    : -1
  const rows = isPrinting.value
    ? layout.rows
    : singleRowIndex >= 0
      ? [layout.rows[singleRowIndex]].filter((row) => row !== undefined)
      : docxRowsInViewport(layout.rows, { startPx: viewportStart, endPx: viewportEnd })

  const entries: VisiblePageEntry[] = []
  const nextEntryCache = new Map<number, VisiblePageEntry>()
  for (const row of rows) {
    for (const pageIndex of row.pageIndexes) {
      if (pageIndex >= pageCount.value) continue
      const pageTop = layout.pageOffsets[pageIndex]
      const pageHeight = layout.pageHeights[pageIndex]
      const previousEntry = pageRenderCache.entries.get(pageIndex)
      const nextEntry = buildVisiblePageEntry(pageIndex, pageTop, pageHeight, previousEntry)
      const sameContent =
        previousEntry !== undefined &&
        sameVisiblePageContent(previousEntry, nextEntry)
      const entry = sameContent
        ? previousEntry.topPx === pageTop &&
          previousEntry.heightPx === pageHeight &&
          previousEntry.rowIndex === nextEntry.rowIndex
          ? previousEntry
          : {
              ...previousEntry,
              topPx: pageTop,
              heightPx: pageHeight,
              rowIndex: nextEntry.rowIndex,
            }
        : nextEntry
      if (!sameContent) {
        pageRenderCache.revisions.set(
          pageIndex,
          (previousEntry ? pageRenderCache.revisions.get(pageIndex) ?? 1 : 0) + 1
        )
      }
      nextEntryCache.set(pageIndex, entry)
      entries.push(entry)
    }
  }

  pageRenderCache.entries = nextEntryCache
  return entries
})

// ── Styles ─────────────────────────────────────────────────────────
const rootStyle = computed<Record<string, string>>(() => {
  if (isPrinting.value) {
    const printStyle: Record<string, string> = {
      flex: "1",
      width: "100%",
      minWidth: "0",
      overflow: "visible",
      background: "#fff",
      padding: "0",
      outline: "none",
    }
    return printStyle
  }
  const screenStyle: Record<string, string> = {
    flex: "1",
    width: "100%",
    minWidth: "0",
    overflowY: "auto",
    overflowX: "auto",
    background: `var(--docx-surface-bg, ${theme.value === "dark" ? "#111827" : "#f4f4f5"})`,
    padding: "24px 0",
    outline: isDragOver.value ? "3px dashed #2563eb" : "none",
  }
  return screenStyle
})

const spacerStyle = computed<Record<string, string>>(() => {
  if (isPrinting.value) {
    return { height: "auto", minWidth: "0", position: "static", width: "100%" }
  }
  return {
    height: `${Math.max(0, totalHeightPx.value)}px`,
    // A spread row is as wide as both papers plus their gutters, so the widest
    // row decides the horizontal scroll width.
    minWidth: `${viewLayout.value.maxRowWidthPx + PAGE_SIDE_PADDING}px`,
    position: "relative",
    width: "100%",
  }
})

function pageWrapperStyle(entry: VisiblePageEntry): Record<string, string | number> {
  if (isPrinting.value) {
    return {
      position: "relative",
      top: "0",
      left: "0",
      transform: "none",
      height: `${entry.pageLayout.pageHeightPx}px`,
      width: `${entry.pageWidthPx}px`,
      margin: "0 auto",
      breakAfter: entry.pageIndex < pageCount.value - 1 ? "page" : "auto",
    }
  }
  // Wrapper width covers paper + gutter so translateX(-50%) centers the whole
  // footprint instead of just the paper (which would push the gutter off-right).
  // Offsets are scroll-space values that include the container's top padding;
  // wrappers are positioned inside the spacer, so subtract that padding again.
  // A spread row shifts each page horizontally around the row centre, so both
  // pages stay side by side while the row as a whole stays centered.
  const row = rowForPage(entry.pageIndex)
  const position = row ? row.pageIndexes.indexOf(entry.pageIndex) : -1
  const offsetPx = row && position >= 0
    ? Math.round(row.pageLeftsPx[position] + row.pageWidthsPx[position] / 2 - row.widthPx / 2)
    : 0
  return {
    position: "absolute",
    top: `${Math.max(0, entry.topPx - PAGE_TOP_PADDING)}px`,
    left: offsetPx === 0 ? "50%" : `calc(50% + ${offsetPx}px)`,
    transform: "translateX(-50%)",
    height: `${entry.heightPx}px`,
    width: `${entry.pageWidthPx * zoomFactor.value + gutterVisualWidthPx.value}px`,
  }
}

function pageContentStyle(entry: VisiblePageEntry): Record<string, string> {
  if (isPrinting.value) {
    return {
      height: `${entry.pageLayout.pageHeightPx}px`,
      transform: "none",
      transformOrigin: "left top",
      width: `${entry.pageWidthPx}px`,
    }
  }
  return {
    height: `${entry.pageLayout.pageHeightPx}px`,
    transform: `scale(${zoomFactor.value})`,
    transformOrigin: "left top",
    width: `${entry.pageWidthPx}px`,
  }
}

// ── Scroll handling ────────────────────────────────────────────────
let scrollTimer: ReturnType<typeof setTimeout> | null = null

function onScroll(): void {
  if (!scrollContainerRef.value) return
  scrollTop.value = scrollContainerRef.value.scrollTop
  viewportHeight.value = scrollContainerRef.value.clientHeight

  if (scrollTimer !== null) {
    clearTimeout(scrollTimer)
  }
  scrollTimer = setTimeout(() => {
    // Single-page mode keeps the reader on a page boundary: after the scroll
    // settles, a partial page is pulled back to the nearest page top.
    if (snapSinglePage()) return
    updateVisibleRange()
  }, SCROLL_MEASUREMENT_DEBOUNCE_MS)
}

/**
 * Snap the scroll offset to the nearest page in `single` mode.
 *
 * Returns true when the offset moved, in which case the following scroll event
 * publishes the new visible range instead of this pass.
 */
function snapSinglePage(): boolean {
  if (isPrinting.value || viewMode.value !== "single") return false
  const container = scrollContainerRef.value
  if (!container) return false
  const top = Math.max(0, container.scrollTop ?? scrollTop.value)
  const target = docxSnapTopForRow(viewLayout.value.rows, top, viewportHeight.value)
  if (target === undefined) return false
  const maxScroll = Math.max(0, totalHeightPx.value - viewportHeight.value)
  const clamped = Math.min(Math.max(0, target), maxScroll)
  if (Math.abs(top - clamped) <= SINGLE_PAGE_SNAP_TOLERANCE_PX) return false
  scrollTop.value = clamped
  scrollContainerTo(container, clamped, "auto")
  updateVisibleRange()
  return true
}

function updateVisibleRange(): void {
  const entries = visiblePageEntries.value
  if (entries.length === 0) {
    emit("visiblePageRange", { startPageIndex: 0, endPageIndex: -1 })
    return
  }
  emit("visiblePageRange", {
    startPageIndex: entries[0].pageIndex,
    endPageIndex: entries[entries.length - 1].pageIndex,
  })
}

// ── Bookmark navigation ────────────────────────────────────────────
const BOOKMARK_HIGHLIGHT_DURATION_MS = 2400
const BOOKMARK_NOTICE_DURATION_MS = 4000
const BOOKMARK_TARGET_LOOKUP_ATTEMPTS = 6

const bookmarkIndex = computed(() => createDocxBookmarkIndex(props.model))
const bookmarkTarget = ref<DocxBookmarkTargetHighlight | undefined>()
const bookmarkNotice = ref<string | undefined>()
const bookmarkNoticeStyle = computed<Record<string, string>>(() => ({
  alignItems: "center",
  background: theme.value === "dark" ? "#1f2937" : "#111827",
  borderRadius: "999px",
  boxShadow: "0 10px 25px rgba(15, 23, 42, 0.28)",
  color: "#ffffff",
  display: "flex",
  fontSize: "13px",
  gap: "8px",
  left: "50%",
  maxWidth: "min(560px, calc(100% - 48px))",
  padding: "10px 18px",
  pointerEvents: "none",
  position: "sticky",
  top: "16px",
  transform: "translateX(-50%)",
  width: "fit-content",
  zIndex: "20",
}))
let bookmarkRevealNonce = 0
let bookmarkHighlightTimer: ReturnType<typeof setTimeout> | undefined
let bookmarkNoticeTimer: ReturnType<typeof setTimeout> | undefined

function pageIndexForNodeIndex(nodeIndex: number): number | undefined {
  const pageIndex = pageNodeSegments.value.findIndex((segments) =>
    segments.some((segment) => segment.nodeIndex === nodeIndex)
  )
  return pageIndex >= 0 ? pageIndex : undefined
}

function scrollPageIntoView(pageIndex: number, behavior: ScrollBehavior = "smooth"): void {
  const container = scrollContainerRef.value
  if (!container || pageIndex < 0 || pageIndex >= pageCount.value) return
  const top = pageOffsets.value[pageIndex] ?? 0
  // The browser's scroll event is asynchronous, but a bookmark target page has
  // to mount before the DOM lookup below can find its paragraph. Publishing the
  // final offset immediately keeps virtualization in step with the scroll.
  scrollTop.value = top
  scrollContainerTo(container, top, behavior)
}

function scheduleBookmarkHighlightClear(): void {
  if (bookmarkHighlightTimer !== undefined) clearTimeout(bookmarkHighlightTimer)
  bookmarkHighlightTimer = setTimeout(() => {
    bookmarkHighlightTimer = undefined
    bookmarkTarget.value = undefined
  }, BOOKMARK_HIGHLIGHT_DURATION_MS)
}

function showBookmarkMissing(name: string): void {
  const label = name.trim() || "#"
  bookmarkNotice.value = formatOfficeMessage(
    m.value.docx.viewer.bookmarkMissing,
    { name: label }
  )
  if (bookmarkNoticeTimer !== undefined) clearTimeout(bookmarkNoticeTimer)
  bookmarkNoticeTimer = setTimeout(() => {
    bookmarkNoticeTimer = undefined
    bookmarkNotice.value = undefined
  }, BOOKMARK_NOTICE_DURATION_MS)
  emit("bookmarkMissing", label)
}

function bookmarkCellSelector(highlight: DocxBookmarkTargetHighlight): string | undefined {
  if (highlight.tableRowIndex === undefined || highlight.tableCellIndex === undefined) {
    return undefined
  }
  return (
    `[data-docx-table-cell-paragraph-host="true"]` +
    `[data-docx-table-node-index="${highlight.nodeIndex}"]` +
    `[data-docx-table-row-index="${highlight.tableRowIndex}"]` +
    `[data-docx-table-cell-index="${highlight.tableCellIndex}"]`
  )
}

function bookmarkTargetElement(
  highlight: DocxBookmarkTargetHighlight
): HTMLElement | undefined {
  const container = scrollContainerRef.value
  if (!container || typeof container.querySelector !== "function") return undefined
  // Query the cell before the table: a selector list returns whichever element
  // comes first in document order, and the table host is the cell's ancestor.
  const cellSelector = bookmarkCellSelector(highlight)
  const cell = cellSelector ? container.querySelector(cellSelector) : null
  if (cell) return cell as HTMLElement
  const paragraph = container.querySelector(
    `[data-docx-paragraph-node-index="${highlight.nodeIndex}"]`
  )
  if (paragraph) return paragraph as HTMLElement
  return (container.querySelector(
    `[data-docx-table-host="true"][data-docx-table-node-index="${highlight.nodeIndex}"]`
  ) as HTMLElement | null) ?? undefined
}

function bookmarkLookupFrame(): Promise<void> {
  return new Promise((resolve) => {
    const raf = (globalThis as { requestAnimationFrame?: (callback: () => void) => unknown })
      .requestAnimationFrame
    if (typeof raf === "function") raf(() => resolve())
    else setTimeout(resolve, 16)
  })
}

async function waitForBookmarkTarget(
  highlight: DocxBookmarkTargetHighlight
): Promise<HTMLElement | undefined> {
  for (let attempt = 0; attempt < BOOKMARK_TARGET_LOOKUP_ATTEMPTS; attempt += 1) {
    const element = bookmarkTargetElement(highlight)
    if (element) return element
    await nextTick()
    await bookmarkLookupFrame()
  }
  return bookmarkTargetElement(highlight)
}

function selectBookmarkTarget(target: DocxBookmarkTarget): void {
  const controller = props.controller
  if (!controller) return
  if (target.tableCell) {
    controller.selectTableCell(
      target.nodeIndex,
      target.tableCell.rowIndex,
      target.tableCell.cellIndex
    )
    return
  }
  controller.selectParagraph(target.nodeIndex)
}

function focusBookmarkTarget(element: HTMLElement): void {
  // A read-only surface has no caret; the highlight and scroll are the whole
  // navigation. An editable target paragraph takes the caret at its start.
  if (!props.editable) return
  element.focus?.()
}

/**
 * Scrolls to the bookmark named `name`, highlights its target and places the
 * caret there. Missing bookmarks surface a non-blocking notice instead of
 * throwing, so a stale table of contents cannot break the viewer.
 */
async function revealBookmark(name: string): Promise<DocxBookmarkRevealResult> {
  const requested = typeof name === "string" ? name.trim() : ""
  const target = bookmarkIndex.value.get(requested)
  if (!target) {
    const label = requested.replace(/^#/, "").trim()
    showBookmarkMissing(label)
    return { status: "missing", name: label }
  }

  const pageIndex = pageIndexForNodeIndex(target.nodeIndex)
  const highlight: DocxBookmarkTargetHighlight = {
    nodeIndex: target.nodeIndex,
    tableRowIndex: target.tableCell?.rowIndex,
    tableCellIndex: target.tableCell?.cellIndex,
    nonce: (bookmarkRevealNonce += 1),
  }
  bookmarkTarget.value = highlight
  scheduleBookmarkHighlightClear()
  selectBookmarkTarget(target)
  if (pageIndex !== undefined) {
    const pageVisible = visiblePageEntries.value.some(
      (entry) => entry.pageIndex === pageIndex
    )
    if (!pageVisible) {
      emit("requestPageReveal", pageIndex)
      // Instant page jump: the paragraph-level smooth scroll below finishes the
      // motion, and the target page stays mounted while we look it up.
      scrollPageIntoView(pageIndex, "auto")
    }
  }

  const element = await waitForBookmarkTarget(highlight)
  if (element) {
    element.scrollIntoView?.({ block: "center", behavior: "smooth" })
    focusBookmarkTarget(element)
  }

  return {
    status: "revealed",
    name: target.name,
    nodeIndex: target.nodeIndex,
    pageIndex,
    ...(target.tableCell
      ? {
          tableRowIndex: target.tableCell.rowIndex,
          tableCellIndex: target.tableCell.cellIndex,
        }
      : undefined),
    text: target.text,
  }
}

// ── Page measurement ───────────────────────────────────────────────
function onPageMeasure(pageIndex: number, event: PageMeasureEvent): void {
  if (Number.isFinite(event.heightPx) && event.heightPx > 0) {
    measuredPageHeights.value = {
      ...measuredPageHeights.value,
      [pageIndex]: { heightPx: event.heightPx, layoutKey: pageLayoutKey(pageIndex) },
    }
  }
}

function pageIdentityKey(pageIndex: number): string {
  return documentPageNodeSegmentsIdentityKey(pageNodeSegments.value[pageIndex] ?? [])
}

function schedulePaginationMeasurementFlush(): void {
  if (measurementFlushScheduled) return
  measurementFlushScheduled = true
  measurementFlushHandle = setTimeout(() => {
    measurementFlushScheduled = false
    measurementFlushHandle = undefined
    flushPaginationMeasurementFeedback()
  }, 0)
}

function onParagraphMeasure(nodeIndex: number, heightPx: number): void {
  if (!Number.isFinite(heightPx) || heightPx <= 0) return
  measuredParagraphHeightsPxByNodeIndex.value = {
    ...measuredParagraphHeightsPxByNodeIndex.value,
    [nodeIndex]: Math.round(heightPx),
  }
  schedulePaginationMeasurementFlush()
}

function onTableRowMeasure(nodeIndex: number, rowHeightsPx: number[]): void {
  if (!Number.isFinite(nodeIndex) || rowHeightsPx.length === 0) return
  const node = props.model.nodes[nodeIndex]
  if (!node || node.type !== "table" || node.rows.length === 0) return

  // A paginated table only renders the rows of the current page slice. Merge
  // those measurements with the rows measured earlier and only feed the table
  // back once every row has a real height; the segmenter treats a shorter
  // array as "no measurement" rather than filling the unseen rows from a
  // guess, and a sparse array would make it look complete.
  const previous = measuredTableRowHeightsPxByNodeIndex.value[nodeIndex] ?? []
  const merged = Array.from({ length: node.rows.length }, (_, rowIndex) => {
    const measured = rowHeightsPx[rowIndex]
    if (Number.isFinite(measured) && measured > 0) {
      return Math.max(1, Math.round(measured))
    }
    const earlier = previous[rowIndex]
    return Number.isFinite(earlier) && earlier > 0 ? Math.round(earlier) : 0
  })
  if (!merged.every((heightPx) => heightPx > 0)) return

  measuredTableRowHeightsPxByNodeIndex.value = {
    ...measuredTableRowHeightsPxByNodeIndex.value,
    [nodeIndex]: merged,
  }
  schedulePaginationMeasurementFlush()
}

function onPageContentMeasure(pageIndex: number, heightPx: number): void {
  if (!Number.isFinite(heightPx) || heightPx <= 0) return
  measuredPageContentHeights.value = {
    ...measuredPageContentHeights.value,
    [pageIndex]: {
      heightPx: Math.round(heightPx),
      identityKey: pageIdentityKey(pageIndex),
    },
  }
  schedulePaginationMeasurementFlush()
}

function flushPaginationMeasurementFeedback(): void {
  const keys = pageNodeSegments.value.map((_, pageIndex) => pageIdentityKey(pageIndex))
  if (keys.length === 0) return

  const paragraphHeightsPxByNodeIndex = { ...measuredParagraphHeightsPxByNodeIndex.value }
  const tableRowHeightsPxByNodeIndex = { ...measuredTableRowHeightsPxByNodeIndex.value }
  if (
    Object.keys(paragraphHeightsPxByNodeIndex).length === 0 &&
    Object.keys(tableRowHeightsPxByNodeIndex).length === 0
  ) {
    return
  }

  // Page content heights are diagnostic only: using them directly as page
  // budgets would let sparse pages shrink their own budget and spiral. They do
  // make convergence/oscillation sensitive to rendered overflow.
  const pageContentHeightsPx = keys.map((key, pageIndex) => {
    const measured = measuredPageContentHeights.value[pageIndex]
    return measured && measured.identityKey === key ? measured.heightPx : 0
  })

  const feedback: PaginationMeasurementFeedback = {
    paragraphHeightsPxByNodeIndex,
    tableRowHeightsPxByNodeIndex,
    pageIdentityKeys: keys,
  }
  const decision = paginationMeasurementLoop.observe({
    pageIdentityKeys: keys,
    pageContentHeightsPx,
    feedback,
  })

  paginationState.value =
    decision.action === "repaginate" ? "repaginating" : decision.action
  paginationIterations.value = decision.iteration
  paginationTargetReached.value = decision.targetReached

  if (decision.action === "repaginate" && decision.feedback) {
    // Capture the reader's page before offsets move. The watcher on
    // `pageOffsets` restores this exact position after the new plan renders.
    pendingPaginationScrollAnchor = resolvePaginationScrollAnchor(
      pageOffsets.value,
      scrollTop.value
    )
    paginationMeasurementFeedback.value = decision.feedback
  }
  measuredPageContentHeights.value = {}
}

// A new document (or a new pagination layout) invalidates every cached height:
// only mounted pages re-measure, so stale values would corrupt offsets for all
// virtualized pages until each one has been visited. It also starts a fresh
// convergence run with the new document's advisory page-count target.
watch(
  () => [
    props.model,
    props.controller?.documentLoadNonce,
    props.layoutOptions,
    effectiveRevisionDisplay.value,
    trackedChangesVisible.value,
  ] as const,
  () => {
    measuredPageHeights.value = {}
    measuredParagraphHeightsPxByNodeIndex.value = {}
    measuredTableRowHeightsPxByNodeIndex.value = {}
    measuredPageContentHeights.value = {}
    paginationMeasurementFeedback.value = undefined
    paginationMeasurementLoop = createPaginationMeasurementLoop()
    paginationState.value = "idle"
    paginationIterations.value = 0
    paginationTargetReached.value = false
    pendingPaginationScrollAnchor = undefined
  }
)

// Restore the anchored page after a measurement pass moved the page offsets.
watch(
  pageOffsets,
  (offsets) => {
    const anchor = pendingPaginationScrollAnchor
    if (!anchor) return
    pendingPaginationScrollAnchor = undefined
    const target = resolvePaginationScrollTop(offsets, anchor)
    if (target === undefined) return
    const container = scrollContainerRef.value
    if (container) container.scrollTop = target
    scrollTop.value = target
  },
  { flush: "post" }
)

// ── Drag-and-drop file import ──────────────────────────────────────
// `useFileDrop` (see the declaration above) owns the handlers, so dropped
// files follow the controller's import path and the outline always clears.

// ── Copy / cut ────────────────────────────────────────────────────
// The native contentEditable copy only carries the rendered DOM; writing the
// model runs here also gives consumers the internal payload and inline styles
// that survive a paste into Word. A false return leaves the event untouched so
// the browser's own copy still runs.
function onCopy(event: ClipboardEvent): void {
  if (!props.controller || event.defaultPrevented || !event.clipboardData) return
  if (props.controller.writeSelectionToClipboardData(event.clipboardData)) {
    event.preventDefault()
  }
}

function onCut(event: ClipboardEvent): void {
  if (!props.editable || !props.controller || event.defaultPrevented || !event.clipboardData) {
    return
  }
  if (!props.controller.writeSelectionToClipboardData(event.clipboardData)) return
  event.preventDefault()
  if (props.controller.selection.kind === "table-cell") {
    props.controller.clearTableSelection()
  } else {
    props.controller.deleteExpandedSelection()
  }
}

// ── Context menu ─────────────────────────────────────────────────────
/** Page under a point, in scroll space. Spread rows resolve the left/right page by x. */
function pageIndexAtClientPoint(clientX: number, clientY: number): number | undefined {
  const container = scrollContainerRef.value
  if (!container) return undefined
  const rect = container.getBoundingClientRect()
  const contentWidth = Math.max(
    container.clientWidth ?? 0,
    viewLayout.value.maxRowWidthPx + PAGE_SIDE_PADDING,
  )
  const x = clientX - rect.left + (container.scrollLeft ?? 0)
  const y = clientY - rect.top + container.scrollTop
  return docxPageAtPoint(viewLayout.value.rows, { x, y, contentWidth })
}

function emitContextMenuAt(clientX: number, clientY: number, target?: EventTarget | null): void {
  const container = scrollContainerRef.value
  if (!container) return
  const rect = container.getBoundingClientRect()
  const pageIndex = pageIndexAtClientPoint(clientX, clientY) ?? 0
  emit("contextMenu", {
    pageIndex,
    clientX,
    clientY,
    containerX: clientX - rect.left,
    containerY: clientY - rect.top,
    target,
  })
}

const touchGestures = createTouchGestureTracker({
  onLongPress: (pointer) => emitContextMenuAt(pointer.clientX, pointer.clientY, pointer.target as EventTarget | null),
})

function onTouchPointerDown(event: PointerEvent): void { touchGestures.pointerDown(event) }
function onTouchPointerMove(event: PointerEvent): void { touchGestures.pointerMove(event) }
function onTouchPointerUp(event: PointerEvent): void { touchGestures.pointerUp(event) }
function onTouchPointerCancel(event: PointerEvent): void { touchGestures.pointerCancel(event) }

function onContextMenu(event: MouseEvent): void {
  if (touchGestures.shouldSuppressContextMenu(event.clientX, event.clientY)) {
    event.preventDefault()
    return
  }
  emitContextMenuAt(event.clientX, event.clientY, event.target)
}

// ── Selection tracking ───────────────────────────────────────────────
function onMouseUp(): void {
  const sel = window.getSelection()
  if (!sel || sel.isCollapsed || !sel.toString().trim()) {
    emit("selectionChange", { kind: "none" })
    return
  }
  const text = sel.toString().trim()
  const container = scrollContainerRef.value
  if (!container) { emit("selectionChange", { kind: "text", text }); return }
  const range = sel.getRangeAt(0)
  const rect = range.getBoundingClientRect()
  const pageIndex = pageIndexAtClientPoint(rect.left, rect.top) ?? 0
  emit("selectionChange", { kind: "text", text, pageIndex })
}

// ── Lifecycle ──────────────────────────────────────────────────────
let containerObserver: ResizeObserver | undefined
let stopPrintDetection: () => void = () => {}

onMounted(() => {
  ensureDocumentFontInvalidation()
  stopFontMetricsSubscription = subscribeFontMetrics(() => {
    fontMetricsRevision.value = getFontMetricsRevision()
  })
  stopPrintDetection = setupPrintDetection()
  if (scrollContainerRef.value) {
    const height = scrollContainerRef.value.clientHeight
    if (Number.isFinite(height) && height > 0) viewportHeight.value = height
    const width = scrollContainerRef.value.clientWidth
    if (Number.isFinite(width) && width > 0) containerWidth.value = width
    scrollContainerRef.value.addEventListener("pointerdown", onTouchPointerDown, { passive: true })
    scrollContainerRef.value.addEventListener("pointermove", onTouchPointerMove, { passive: true })
    scrollContainerRef.value.addEventListener("pointerup", onTouchPointerUp)
    scrollContainerRef.value.addEventListener("pointercancel", onTouchPointerCancel)
  }
  const ResizeObserverConstructor = globalThis.ResizeObserver
  if (typeof ResizeObserverConstructor !== "function") {
    updateVisibleRange()
    return
  }
  containerObserver = new ResizeObserverConstructor((entries) => {
    for (const entry of entries) {
      viewportHeight.value = entry.contentRect.height
      containerWidth.value = entry.contentRect.width
    }
  })
  if (scrollContainerRef.value) {
    containerObserver.observe(scrollContainerRef.value)
  }
})

onUnmounted(() => {
  stopPrintDetection()
  stopFontMetricsSubscription?.()
  stopFontMetricsSubscription = undefined
  printPageStyle?.remove?.()
  printPageStyle = undefined
  containerObserver?.disconnect()
  containerObserver = undefined
  if (scrollTimer !== null) clearTimeout(scrollTimer)
  if (bookmarkHighlightTimer !== undefined) clearTimeout(bookmarkHighlightTimer)
  if (bookmarkNoticeTimer !== undefined) clearTimeout(bookmarkNoticeTimer)
  bookmarkHighlightTimer = undefined
  bookmarkNoticeTimer = undefined
  if (measurementFlushHandle !== undefined) clearTimeout(measurementFlushHandle)
  measurementFlushHandle = undefined
  touchGestures.reset()
  const container = scrollContainerRef.value
  container?.removeEventListener("pointerdown", onTouchPointerDown)
  container?.removeEventListener("pointermove", onTouchPointerMove)
  container?.removeEventListener("pointerup", onTouchPointerUp)
  container?.removeEventListener("pointercancel", onTouchPointerCancel)
})

// ── Expose ─────────────────────────────────────────────────────────
function scrollContainerTo(container: HTMLElement, top: number, behavior: ScrollBehavior): void {
  // Test/render environments may expose the host element without the full
  // HTMLElement scroll API; assigning scrollTop still keeps page navigation
  // observable instead of throwing on a missing scrollTo.
  if (typeof container.scrollTo === "function") {
    container.scrollTo({ top, behavior })
  } else {
    container.scrollTop = top
  }
}

defineExpose({
  scrollToPage(pageIndex: number): void {
    scrollPageIntoView(pageIndex)
  },
  /**
   * Reveals the bookmark named `name`: scrolls its page/paragraph into view,
   * highlights the target and places the caret. Missing bookmarks resolve to
   * `{ status: "missing" }` and show a non-blocking notice.
   */
  revealBookmark(name: string): Promise<DocxBookmarkRevealResult> {
    return revealBookmark(name)
  },
  scrollToNode(nodeIndex: number, options?: ScrollToOptions): void {
    const pageIndex = pageNodeSegments.value.findIndex((segments) =>
      segments.some((segment) => segment.nodeIndex === nodeIndex)
    )
    if (!scrollContainerRef.value || pageIndex < 0) return
    scrollContainerTo(scrollContainerRef.value, pageOffsets.value[pageIndex] ?? 0, options?.behavior ?? "smooth")
  },
  /** Page rendering a header/footer part or a note body, for search navigation. */
  searchTargetPage(target: DocxSearchPageTarget): number | undefined {
    return findDocxSearchTargetPage(target, pagePresentations.value, pageNotes.value)
  },
  get scrollContainer() {
    return scrollContainerRef.value
  },
})
</script>

<style scoped>
.docx-viewer-root {
  flex: 1;
  overflow-y: auto;
  overflow-x: hidden;
  background: #f3f4f6;
  touch-action: pan-x pan-y;
}

@media print {
  .docx-viewer-root {
    height: auto !important;
    overflow: visible !important;
    padding: 0 !important;
    background: #fff !important;
  }
}
</style>
