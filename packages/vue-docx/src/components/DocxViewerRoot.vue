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
            :page-width-px="entry.pageWidthPx"
            :page-content-width-px="entry.pageContentWidthPx"
            :page-number="entry.pageNumber"
            :total-pages="pageCount"
            :page-number-format="entry.pageNumberFormat"
            :header-section="entry.headerSection"
            :footer-section="entry.footerSection"
            :tracked-changes-enabled="trackedChangesEnabled"
            :comments-enabled="commentsEnabled"
            :tracked-changes="trackedChanges"
            :comments="comments"
            :footnotes="pageFootnotes(entry.pageIndex)"
            :endnotes="pageEndnotes(entry.pageIndex)"
            :search-query="searchQuery"
            :active-search-node-index="activeSearchNodeIndex"
            :theme="theme"
            @measure="onPageMeasure(entry.pageIndex, $event)"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, watch } from "vue"
import type {
  DocModel,
  DocxComment,
  DocxDocumentTheme,
  DocxEditorController,
  DocxTrackedChange,
  DocumentNoteDefinition,
  DocumentPageNodeSegment,
  FooterSection,
  HeaderSection,
  LayoutOptions,
} from "@arcships/docx-core"
import {
  buildLayoutSnapshot,
  collectCommentsFromModel,
  collectTrackedChangesFromModel,
  parseSectionLayout,
  parseSectionPageNumberFormat,
  parseSectionPageNumberStartOverride,
  resolveDocumentSectionsFromMetadata,
  resolveDocumentLayout,
  resolveSectionIndexForNodeIndex,
  nodeReferencedNoteIds,
  selectSectionVariantForPage,
  DEFAULT_DOC_PAGE_WIDTH,
  DEFAULT_DOC_PAGE_HEIGHT,
} from "@arcships/docx-core"
import DocxPageSurface from "./DocxPageSurface.vue"

// ── Constants ──────────────────────────────────────────────────────
const DOC_PAGE_BREAK_GAP = 28
const PAGE_OVERSCAN = 2
const SCROLL_MEASUREMENT_DEBOUNCE_MS = 80
// Gutter (tracked changes / comments) sits beside the paper, inside the scaled
// content. It must enter the horizontal footprint so the paper stays centered
// and horizontal scroll can reach it.
const GUTTER_GAP_PX = 16
const GUTTER_WIDTH_PX = 240
const PAGE_SIDE_PADDING = 48

// ── Types ──────────────────────────────────────────────────────────
export interface DocxViewerRootProps {
  model: DocModel
  controller?: DocxEditorController
  editable?: boolean
  layoutOptions?: LayoutOptions
  theme?: DocxDocumentTheme
  zoomScale?: number
  showTrackedChanges?: boolean
  showComments?: boolean
  searchQuery?: string
  activeSearchNodeIndex?: number
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
}

interface PageMeasureEvent {
  heightPx: number
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<DocxViewerRootProps>(),
  { editable: false, zoomScale: 100 }
)

const emit = defineEmits<{
  pageCountChange: [count: number]
  visiblePageRange: [range: { startPageIndex: number; endPageIndex: number }]
}>()

// ── State ──────────────────────────────────────────────────────────
const scrollContainerRef = ref<HTMLElement | undefined>()
const scrollTop = ref(0)
const viewportHeight = ref(800)

// Track measured page heights (key: pageIndex, value: px)
const measuredPageHeights = ref<Record<number, number>>({})

// Drag-over state for file drops
const dragOverDepth = ref(0)
const isDragOver = computed(() => dragOverDepth.value > 0)

// ── Pagination ─────────────────────────────────────────────────────
const baseDocumentLayout = computed(() => resolveDocumentLayout(props.model))

const resolvedLayoutOptions = computed<LayoutOptions>(() => ({
  pageWidth: baseDocumentLayout.value.pageWidthPx,
  pageHeight: baseDocumentLayout.value.pageHeightPx,
  margin: baseDocumentLayout.value.marginsPx.top,
  ...props.layoutOptions,
}))

const paginationPlan = computed(() => {
  const snapshot = buildLayoutSnapshot(props.model, resolvedLayoutOptions.value)
  const pageNodeSegments: DocumentPageNodeSegment[][] = snapshot.pages.map((page) =>
    page.blocks.flatMap((block) => {
      const nodeIndex = block.source?.nodeIndex
      return Number.isFinite(nodeIndex) && (nodeIndex as number) >= 0
        ? [{ nodeIndex: Math.round(nodeIndex as number) }]
        : []
    })
  )

  return { pages: snapshot.pages, pageNodeSegments }
})

const pages = computed(() => paginationPlan.value.pages)
const pageNodeSegments = computed(() => paginationPlan.value.pageNodeSegments)
const pageCount = computed(() => pages.value.length)
const zoomFactor = computed(() => Math.min(2, Math.max(0.5, props.zoomScale / 100)))

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

// ── Theme / feature flags ──────────────────────────────────────────
const theme = computed(() => props.theme ?? props.controller?.documentTheme ?? "light")
const trackedChangesEnabled = computed(() => props.showTrackedChanges ?? props.controller?.showTrackedChanges ?? false)
const commentsEnabled = computed(() => props.showComments ?? props.controller?.showComments ?? false)
// Visual horizontal space the annotation gutter occupies at the current zoom.
// 0 when annotations are hidden, so no space is reserved.
const gutterVisualWidthPx = computed(() =>
  (trackedChangesEnabled.value || commentsEnabled.value)
    ? (GUTTER_GAP_PX + GUTTER_WIDTH_PX) * zoomFactor.value
    : 0
)
const trackedChanges = computed<readonly DocxTrackedChange[]>(() =>
  props.controller?.trackedChanges ?? (trackedChangesEnabled.value ? collectTrackedChangesFromModel(props.model) : [])
)
const comments = computed<readonly DocxComment[]>(() =>
  props.controller?.comments ?? (commentsEnabled.value ? collectCommentsFromModel(props.model) : [])
)

function notesForPage(pageIndex: number, kind: "footnote" | "endnote"): DocumentNoteDefinition[] {
  const definitions = kind === "footnote"
    ? props.model.metadata.footnotes ?? []
    : props.model.metadata.endnotes ?? []
  if (!definitions.length) return []
  if (kind === "endnote") return pageIndex === pageCount.value - 1 ? definitions : []
  const byId = new Map(definitions.map((note) => [note.id, note]))
  const ids: number[] = []
  const seen = new Set<number>()
  for (const segment of pageNodeSegments.value[pageIndex] ?? []) {
    const node = props.model.nodes[segment.nodeIndex]
    if (!node) continue
    for (const id of nodeReferencedNoteIds(node, kind, segment.tableRowRange, segment.paragraphLineRange)) {
      if (!seen.has(id)) { seen.add(id); ids.push(id) }
    }
  }
  return ids.map((id) => byId.get(id)).filter((note): note is DocumentNoteDefinition => Boolean(note))
}

function pageFootnotes(pageIndex: number): DocumentNoteDefinition[] { return notesForPage(pageIndex, "footnote") }
function pageEndnotes(pageIndex: number): DocumentNoteDefinition[] { return notesForPage(pageIndex, "endnote") }

// ── Page heights ───────────────────────────────────────────────────
function getPageHeight(pageIndex: number): number {
  const measured = measuredPageHeights.value[pageIndex]
  if (Number.isFinite(measured) && (measured as number) > 0) {
    return measured as number
  }
  return pagePresentations.value[pageIndex]?.pageLayout.pageHeightPx ?? DEFAULT_DOC_PAGE_HEIGHT
}

// ── Page offsets ───────────────────────────────────────────────────
const pageOffsets = computed(() => {
  const offsets: number[] = []
  let offset = 0
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
  return {
    position: "absolute",
    top: `${entry.topPx}px`,
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
      [pageIndex]: event.heightPx,
    }
  }
}

// ── Drag-and-drop file import ──────────────────────────────────────
function onDragOver(event: DragEvent): void {
  if (!event.dataTransfer) return
  if (props.editable && props.controller && hasDocxFile(event.dataTransfer)) {
    dragOverDepth.value++
    event.dataTransfer.dropEffect = "copy"
  }
}

function onDragLeave(): void {
  dragOverDepth.value = Math.max(0, dragOverDepth.value - 1)
}

async function onDrop(event: DragEvent): Promise<void> {
  dragOverDepth.value = 0
  if (!props.editable || !props.controller) return
  const file = event.dataTransfer?.files?.[0]
  if (file && file.name.endsWith(".docx")) {
    await props.controller.importDocxFile(file)
  }
}

function hasDocxFile(dataTransfer: DataTransfer): boolean {
  if (dataTransfer.types.includes("Files")) {
    for (let i = 0; i < dataTransfer.files.length; i++) {
      if (dataTransfer.files[i].name.endsWith(".docx")) return true
    }
  }
  return false
}

// ── Lifecycle ──────────────────────────────────────────────────────
onMounted(() => {
  if (scrollContainerRef.value) {
    const height = scrollContainerRef.value.clientHeight
    if (Number.isFinite(height) && height > 0) viewportHeight.value = height
  }
  const ResizeObserverConstructor = globalThis.ResizeObserver
  if (typeof ResizeObserverConstructor !== "function") {
    updateVisibleRange()
    return
  }
  const observer = new ResizeObserverConstructor((entries) => {
    for (const entry of entries) {
      viewportHeight.value = entry.contentRect.height
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
  scrollToNode(nodeIndex: number): void {
    const pageIndex = pageNodeSegments.value.findIndex((segments) =>
      segments.some((segment) => segment.nodeIndex === nodeIndex)
    )
    if (!scrollContainerRef.value || pageIndex < 0) return
    const targetTop = pageOffsets.value[pageIndex] ?? 0
    scrollContainerRef.value.scrollTo({ top: targetTop, behavior: "smooth" })
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
