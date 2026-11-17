<template>
  <div
    ref="scrollContainerRef"
    class="docx-viewer-root"
    :data-docx-page-count="pageCount"
    :style="rootStyle"
    @scroll.passive="onScroll"
    @dragover.prevent="onDragOver"
    @dragleave="onDragLeave"
    @drop.prevent="onDrop"
    @contextmenu.prevent="onContextMenu"
    @mouseup="onMouseUp"
  >
    <div :style="spacerStyle">
      <div
        v-for="entry of visiblePageEntries"
        :key="`page-${entry.pageIndex}`"
        :data-docx-page-wrapper="true"
        :data-docx-page-index="entry.pageIndex"
        :style="pageWrapperStyle(entry)"
      >
        <div :style="pageContentStyle(entry)">
          <DocxPageSurface
            :page-index="entry.pageIndex"
            :page-layout="entry.pageLayout"
            :page-node-segments="entry.pageNodeSegments"
            :model="model"
            :controller="controller"
            :editable="editable"
            :zoom-factor="zoomFactor"
            :page-width-px="entry.pageWidthPx"
            :page-content-width-px="entry.pageContentWidthPx"
            :page-number="entry.pageNumber"
            :total-pages="pageCount"
            :page-number-format="entry.pageNumberFormat"
            :header-section="entry.headerSection"
            :footer-section="entry.footerSection"
            :tracked-changes-enabled="trackedChangesEnabled"
            :comments-enabled="commentsEnabled"
            :tracked-changes="pageTrackedChanges(entry.pageIndex)"
            :comments="pageComments(entry.pageIndex)"
            :footnotes="entry.footnotes"
            :endnotes="entry.endnotes"
            :footnote-labels="entry.noteLabels.footnote"
            :endnote-labels="entry.noteLabels.endnote"
            :note-marker-labels="entry.noteLabels"
            :page-background-color="documentBackgroundColor"
            :page-borders="entry.pageBorders ?? null"
            :theme="theme"
            @measure="onPageMeasure(entry.pageIndex, $event)"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, watch, watchEffect } from "vue"
import type {
  DocModel,
  DocxComment,
  DocxDocumentTheme,
  DocxEditorController,
  DocxTrackedChange,
  DocumentNoteDefinition,
  DocumentPageNodeSegment,
  DocumentPageNotes,
  DocumentPageBorders,
  FooterSection,
  HeaderSection,
  LayoutOptions,
} from "@arcships/docx-core"
import {
  buildLayoutSnapshot,
  buildDocumentPageNotes,
  collectCommentsFromModel,
  collectTrackedChangesFromModel,
  parseSectionLayout,
  parseSectionPageBorders,
  parseSectionPageNumberFormat,
  parseSectionPageNumberStartOverride,
  primeCommentAnchors,
  resolveDocumentBackgroundColor,
  resolveDocumentSectionsFromMetadata,
  resolveDocumentLayout,
  resolveSectionIndexForNodeIndex,
  groupGutterAnnotationsByPage,
  trackedChangeDuplicatesInlineMarkup,
  selectSectionVariantForPage,
  DEFAULT_DOC_PAGE_WIDTH,
  DEFAULT_DOC_PAGE_HEIGHT,
} from "@arcships/docx-core"
import DocxPageSurface from "./DocxPageSurface.vue"
import {
  DOCX_UNSUPPORTED_FILE_MESSAGE,
  isSupportedDocxFileName,
} from "../composables/docx-file-types"
import {
  ensureDocxViewerPageSurfaceRegistry,
  notifyDocxViewerPageSurfaceListeners,
} from "../composables/page-surface-registry"
import {
  annotationGutterReservedWidthPx,
} from "./gutter-metrics"

// ── Constants ──────────────────────────────────────────────────────
const DOC_PAGE_BREAK_GAP = 28
const PAGE_OVERSCAN = 2
const SCROLL_MEASUREMENT_DEBOUNCE_MS = 80
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
  /** When true, zoom auto-fits page width to container. Overrides zoomScale. */
  fitWidth?: boolean
  /** `null` follows the controller flag; `false` always hides. */
  showTrackedChanges?: boolean | null
  showComments?: boolean | null
}

interface VisiblePageEntry {
  pageIndex: number
  topPx: number
  heightPx: number
  pageLayout: ReturnType<typeof resolveDocumentLayout>
  pageNodeSegments: DocumentPageNodeSegment[]
  pageWidthPx: number
  pageContentWidthPx: number
  pageNumber: number
  pageNumberFormat?: string
  headerSection?: HeaderSection
  footerSection?: FooterSection
  footnotes: DocumentNoteDefinition[]
  endnotes: DocumentNoteDefinition[]
  noteLabels: { footnote: Map<number, string>; endnote: Map<number, string> }
  pageBorders?: DocumentPageBorders
}

interface PageMeasureEvent {
  heightPx: number
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<DocxViewerRootProps>(),
  { editable: false, zoomScale: 100, showTrackedChanges: null, showComments: null }
)

const emit = defineEmits<{
  pageCountChange: [count: number]
  visiblePageRange: [range: { startPageIndex: number; endPageIndex: number }]
  contextMenu: [ctx: { pageIndex: number; clientX: number; clientY: number; containerX: number; containerY: number; target?: EventTarget | null }]
  selectionChange: [sel: { kind: string; text?: string; pageIndex?: number }]
}>()

// ── State ──────────────────────────────────────────────────────────
const scrollContainerRef = ref<HTMLElement | undefined>()
const scrollTop = ref(0)
const viewportHeight = ref(800)
const containerWidth = ref(800)

// Track measured page heights (key: pageIndex). Each entry is tagged with the
// page layout it was measured for so layout changes invalidate stale values.
const measuredPageHeights = ref<Record<number, { heightPx: number; layoutKey: string }>>({})

// Drag-over state for file drops
const isDragOver = ref(false)

// ── Pagination ─────────────────────────────────────────────────────
const baseDocumentLayout = computed(() => resolveDocumentLayout(props.model))
const documentBackgroundColor = computed(() =>
  resolveDocumentBackgroundColor(props.model.metadata.documentBackgroundColor)
)

const resolvedLayoutOptions = computed<LayoutOptions>(() => ({
  pageWidth: baseDocumentLayout.value.pageWidthPx,
  pageHeight: baseDocumentLayout.value.pageHeightPx,
  margin: baseDocumentLayout.value.marginsPx.top,
  ...props.layoutOptions,
}))

const paginationPlan = computed(() => {
  const snapshot = buildLayoutSnapshot(props.model, resolvedLayoutOptions.value)
  const pageNodeSegments: DocumentPageNodeSegment[][] = snapshot.pages.map((page) =>
    page.blocks.flatMap((block): DocumentPageNodeSegment[] => {
      const nodeIndex = block.source?.nodeIndex
      if (!Number.isFinite(nodeIndex) || (nodeIndex as number) < 0) return []

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
      return [segment]
    })
  )

  return { pages: snapshot.pages, pageNodeSegments, snapshot }
})

const pages = computed(() => paginationPlan.value.pages)
const pageNodeSegments = computed(() => paginationPlan.value.pageNodeSegments)
const pageCount = computed(() => pages.value.length)
const FIT_ZOOM_MIN = 0.25
const FIT_ZOOM_MAX = 2
const FIT_ZOOM_SIDE_PAD = 48

// Horizontal space the annotation gutter reserves beside the paper, in screen
// px. The pane counter-scales itself (see gutter-metrics), so this footprint is
// constant at every zoom level. 0 when annotations are hidden.
const gutterScreenWidthPx = computed(() =>
  (
    (props.showTrackedChanges ?? props.controller?.showTrackedChanges ?? false) ||
    (props.showComments ?? props.controller?.showComments ?? false)
  )
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

const zoomFactor = computed(() => {
  if (props.fitWidth) return fitZoom.value
  return Math.min(2, Math.max(0.5, props.zoomScale / 100))
})

watch(pageCount, (count) => emit("pageCountChange", count), { immediate: true })

// ── Document layout ────────────────────────────────────────────────
const pagePresentations = computed(() => {
  const sections = resolveDocumentSectionsFromMetadata(props.model.metadata)
  const firstPageBySection = new Map<number, number>()
  const evenAndOddHeaders = props.model.metadata.compatibility?.evenAndOddHeaders ?? false
  let previousSectionIndex = 0
  let previousNodeIndex = 0
  let nextPageNumber = 1

  return pageNodeSegments.value.map((segments, pageIndex) => {
    const firstNodeIndex = segments[0]?.nodeIndex ?? previousNodeIndex
    previousNodeIndex = firstNodeIndex
    const sectionIndex = resolveSectionIndexForNodeIndex(
      sections,
      firstNodeIndex,
      previousSectionIndex
    )
    previousSectionIndex = sectionIndex
    if (!firstPageBySection.has(sectionIndex)) {
      firstPageBySection.set(sectionIndex, pageIndex)
      const sectionStart = parseSectionPageNumberStartOverride(
        sections[sectionIndex]?.sectionPropertiesXml
      )
      if (sectionStart !== undefined) nextPageNumber = sectionStart
    }

    const section = sections[sectionIndex]
    const sectionPageIndex = pageIndex - (firstPageBySection.get(sectionIndex) ?? pageIndex)
    const sectionPropertiesXml =
      section?.sectionPropertiesXml ?? props.model.metadata.sectionPropertiesXml
    const pageLayout = sectionPropertiesXml
      ? parseSectionLayout(sectionPropertiesXml)
      : baseDocumentLayout.value
    const variantOptions = { evenAndOddHeaders, documentPageIndex: pageIndex }
    const pageNumber = nextPageNumber
    nextPageNumber += 1

    return {
      pageLayout,
      pageNumber,
      pageNumberFormat: parseSectionPageNumberFormat(sectionPropertiesXml),
      pageBorders: parseSectionPageBorders(sectionPropertiesXml),
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

// ── Shared layout for thumbnail consumers ──────────────────────────
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
    notifyDocxViewerPageSurfaceListeners(controller)
  }, { flush: "post" })
}

// ── Theme / feature flags ──────────────────────────────────────────
const theme = computed(() => props.theme ?? props.controller?.documentTheme ?? "light")
// Boolean props default to `false`, so `null` is what marks "not passed":
// without it the controller fallback below could never apply.
const trackedChangesEnabled = computed(
  () => props.showTrackedChanges ?? props.controller?.showTrackedChanges ?? false
)
const commentsEnabled = computed(
  () => props.showComments ?? props.controller?.showComments ?? false
)
// Visual horizontal space the annotation gutter occupies at the current zoom.
// Visual horizontal space the annotation gutter occupies in screen px.
// 0 when annotations are hidden, so no space is reserved. The pane counter-
// scales, so this footprint does not follow the zoom factor.
const gutterVisualWidthPx = computed(() => gutterScreenWidthPx.value)
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
  buildDocumentPageNotes(props.model, pageNodeSegments.value)
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
const pageOffsets = computed(() => {
  const offsets: number[] = []
  let offset = PAGE_TOP_PADDING
  for (let i = 0; i < pageCount.value; i++) {
    offsets.push(offset)
    offset += getPageHeight(i) * zoomFactor.value + DOC_PAGE_BREAK_GAP
  }
  return offsets
})

const totalHeightPx = computed(() => {
  if (pageCount.value === 0) return 0
  const lastOffset = pageOffsets.value[pageOffsets.value.length - 1]
  return lastOffset + getPageHeight(pageCount.value - 1) * zoomFactor.value + DOC_PAGE_BREAK_GAP
})

// ── Visible page calculation ───────────────────────────────────────
const visiblePageEntries = computed<VisiblePageEntry[]>(() => {
  if (pageCount.value === 0) return []

  const viewportStart = Math.max(0, scrollTop.value - DOC_PAGE_BREAK_GAP * PAGE_OVERSCAN)
  const viewportEnd = viewportStart + viewportHeight.value + DOC_PAGE_BREAK_GAP * PAGE_OVERSCAN * 2

  const entries: VisiblePageEntry[] = []
  for (let i = 0; i < pageCount.value; i++) {
    const pageTop = pageOffsets.value[i]
    const pageHeight = getPageHeight(i) * zoomFactor.value
    const pageBottom = pageTop + pageHeight

    if (pageBottom >= viewportStart && pageTop <= viewportEnd) {
      entries.push({
        pageIndex: i,
        topPx: pageTop,
        heightPx: pageHeight,
        pageLayout: pagePresentations.value[i]?.pageLayout ?? baseDocumentLayout.value,
        pageNodeSegments: pageNodeSegments.value[i] ?? [],
        pageWidthPx:
          pagePresentations.value[i]?.pageLayout.pageWidthPx ?? DEFAULT_DOC_PAGE_WIDTH,
        pageContentWidthPx: Math.max(
          0,
          (pagePresentations.value[i]?.pageLayout.pageWidthPx ?? DEFAULT_DOC_PAGE_WIDTH) -
            (pagePresentations.value[i]?.pageLayout.marginsPx.left ?? 72) -
            (pagePresentations.value[i]?.pageLayout.marginsPx.right ?? 72)
        ),
        pageNumber: pagePresentations.value[i]?.pageNumber ?? i + 1,
        pageNumberFormat: pagePresentations.value[i]?.pageNumberFormat,
        headerSection: pagePresentations.value[i]?.headerSection,
        footerSection: pagePresentations.value[i]?.footerSection,
        footnotes: pageFootnoteNotes(i),
        endnotes: pageEndnoteNotes(i),
        noteLabels: {
          footnote: notesForPage(i).footnoteLabels,
          endnote: notesForPage(i).endnoteLabels,
        },
        pageBorders: pagePresentations.value[i]?.pageBorders,
      })
    }
  }

  return entries
})

// ── Styles ─────────────────────────────────────────────────────────
const rootStyle = computed(() => ({
  flex: "1",
  width: "100%",
  minWidth: "0",
  overflowY: "auto" as const,
  overflowX: "auto" as const,
  background: `var(--docx-surface-bg, ${theme.value === "dark" ? "#111827" : "#f4f4f5"})`,
  padding: "24px 0",
  outline: isDragOver.value ? "3px dashed #2563eb" : "none",
}))

const spacerStyle = computed(() => ({
  height: `${Math.max(0, totalHeightPx.value)}px`,
  minWidth: `${Math.max(
    0,
    ...pagePresentations.value.map(
      (entry) => entry.pageLayout.pageWidthPx * zoomFactor.value + gutterVisualWidthPx.value
    )
  ) + PAGE_SIDE_PADDING}px`,
  position: "relative" as const,
  width: "100%",
}))

function pageWrapperStyle(entry: VisiblePageEntry): Record<string, string | number> {
  // Wrapper width covers paper + gutter so translateX(-50%) centers the whole
  // footprint instead of just the paper (which would push the gutter off-right).
  // Offsets are scroll-space values that include the container's top padding;
  // wrappers are positioned inside the spacer, so subtract that padding again.
  return {
    position: "absolute",
    top: `${Math.max(0, entry.topPx - PAGE_TOP_PADDING)}px`,
    left: "50%",
    transform: "translateX(-50%)",
    height: `${entry.heightPx}px`,
    width: `${entry.pageWidthPx * zoomFactor.value + gutterVisualWidthPx.value}px`,
  }
}

function pageContentStyle(entry: VisiblePageEntry): Record<string, string> {
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
    updateVisibleRange()
  }, SCROLL_MEASUREMENT_DEBOUNCE_MS)
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

// ── Page measurement ───────────────────────────────────────────────
function onPageMeasure(pageIndex: number, event: PageMeasureEvent): void {
  if (Number.isFinite(event.heightPx) && event.heightPx > 0) {
    measuredPageHeights.value = {
      ...measuredPageHeights.value,
      [pageIndex]: { heightPx: event.heightPx, layoutKey: pageLayoutKey(pageIndex) },
    }
  }
}

// A new document (or a new pagination layout) invalidates every cached height:
// only mounted pages re-measure, so stale values would corrupt offsets for all
// virtualized pages until each one has been visited.
watch(
  () => [props.model, props.controller?.documentLoadNonce, props.layoutOptions] as const,
  () => { measuredPageHeights.value = {} }
)

// ── Drag-and-drop file import ──────────────────────────────────────
function onDragOver(event: DragEvent): void {
  if (!event.dataTransfer) return
  if (props.editable && props.controller && hasDocxFile(event.dataTransfer)) {
    // `dragover` fires continuously while inside the element; a counter would
    // only ever grow because `dragleave` fires once when the pointer leaves.
    isDragOver.value = true
    event.dataTransfer.dropEffect = "copy"
  }
}

function onDragLeave(event?: DragEvent): void {
  const container = scrollContainerRef.value
  const related = event?.relatedTarget as Node | null | undefined
  if (related && container && typeof container.contains === "function" && container.contains(related)) {
    return
  }
  isDragOver.value = false
}

async function onDrop(event: DragEvent): Promise<void> {
  isDragOver.value = false
  if (!props.editable || !props.controller) return

  const files = event.dataTransfer?.files
  if (!files || files.length === 0) return

  for (let i = 0; i < files.length; i++) {
    if (isSupportedDocxFileName(files[i].name)) {
      await props.controller.importDocxFile(files[i])
      return
    }
  }

  props.controller.setStatus(DOCX_UNSUPPORTED_FILE_MESSAGE)
}

// ── Context menu ─────────────────────────────────────────────────────
function onContextMenu(event: MouseEvent): void {
  const container = scrollContainerRef.value
  if (!container) return
  const rect = container.getBoundingClientRect()
  const clickY = event.clientY - rect.top + container.scrollTop
  let pageIndex = 0
  for (let i = 0; i < pageCount.value; i++) {
    const top = pageOffsets.value[i]
    const bottom = top + getPageHeight(i) * zoomFactor.value + DOC_PAGE_BREAK_GAP
    if (clickY >= top && clickY < bottom) { pageIndex = i; break }
  }
  emit("contextMenu", {
    pageIndex,
    clientX: event.clientX,
    clientY: event.clientY,
    containerX: event.clientX - rect.left,
    containerY: event.clientY - rect.top,
    target: event.target,
  })
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
  const clickY = rect.top - container.getBoundingClientRect().top + container.scrollTop
  let pageIndex = 0
  for (let i = 0; i < pageCount.value; i++) {
    const top = pageOffsets.value[i]
    const bottom = top + getPageHeight(i) * zoomFactor.value + DOC_PAGE_BREAK_GAP
    if (clickY >= top && clickY < bottom) { pageIndex = i; break }
  }
  emit("selectionChange", { kind: "text", text, pageIndex })
}

function hasDocxFile(dataTransfer: DataTransfer): boolean {
  if (dataTransfer.types.includes("Files")) {
    for (let i = 0; i < dataTransfer.files.length; i++) {
      if (isSupportedDocxFileName(dataTransfer.files[i].name)) return true
    }
  }
  return false
}

// ── Lifecycle ──────────────────────────────────────────────────────
onMounted(() => {
  if (scrollContainerRef.value) {
    const height = scrollContainerRef.value.clientHeight
    if (Number.isFinite(height) && height > 0) viewportHeight.value = height
    const width = scrollContainerRef.value.clientWidth
    if (Number.isFinite(width) && width > 0) containerWidth.value = width
  }
  const ResizeObserverConstructor = globalThis.ResizeObserver
  if (typeof ResizeObserverConstructor !== "function") {
    updateVisibleRange()
    return
  }
  const observer = new ResizeObserverConstructor((entries) => {
    for (const entry of entries) {
      viewportHeight.value = entry.contentRect.height
      containerWidth.value = entry.contentRect.width
    }
  })
  if (scrollContainerRef.value) {
    observer.observe(scrollContainerRef.value)
  }

  onUnmounted(() => {
    observer.disconnect()
    if (scrollTimer !== null) clearTimeout(scrollTimer)
  })
})

// ── Expose ─────────────────────────────────────────────────────────
defineExpose({
  scrollToPage(pageIndex: number): void {
    if (!scrollContainerRef.value || pageIndex < 0 || pageIndex >= pageCount.value)
      return
    const targetTop = pageOffsets.value[pageIndex] ?? 0
    scrollContainerRef.value.scrollTo({ top: targetTop, behavior: "smooth" })
  },
  scrollToNode(nodeIndex: number, options?: ScrollToOptions): void {
    const pageIndex = pageNodeSegments.value.findIndex((segments) =>
      segments.some((segment) => segment.nodeIndex === nodeIndex)
    )
    if (!scrollContainerRef.value || pageIndex < 0) return
    const targetTop = pageOffsets.value[pageIndex] ?? 0
    scrollContainerRef.value.scrollTo({
      top: targetTop,
      behavior: options?.behavior ?? "smooth",
    })
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
}
</style>
