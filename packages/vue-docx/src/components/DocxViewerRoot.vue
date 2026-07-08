<template>
  <div
    ref="scrollContainerRef"
    class="docx-viewer-root"
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
        <DocxPageSurface
          :page-index="entry.pageIndex"
          :page-layout="entry.pageLayout"
          :page-node-segments="entry.pageNodeSegments"
          :controller="controller"
          :editable="editable"
          :page-width-px="pageWidthPx"
          :page-content-width-px="pageContentWidthPx"
          :tracked-changes-enabled="trackedChangesEnabled"
          :comments-enabled="commentsEnabled"
          :theme="theme"
          @measure="onPageMeasure(entry.pageIndex, $event)"
        />
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, inject, watch, shallowRef } from "vue"
import type { DocxEditorController } from "@extend-ai/docx-core"
import {
  layoutDocument,
  resolveDocumentLayout,
  DEFAULT_DOC_PAGE_WIDTH,
  DEFAULT_DOC_PAGE_HEIGHT,
} from "@extend-ai/docx-core"
import type {
  LayoutPage,
  LayoutBlock,
  DocumentPageNodeSegment,
} from "@extend-ai/docx-core"
import DocxPageSurface from "./DocxPageSurface.vue"

// ── Constants ──────────────────────────────────────────────────────
const DOC_PAGE_BREAK_GAP = 28
const PAGE_OVERSCAN = 2
const SCROLL_MEASUREMENT_DEBOUNCE_MS = 80

// ── Types ──────────────────────────────────────────────────────────
export interface DocxViewerRootProps {
  controller: DocxEditorController
  editable?: boolean
}

interface VisiblePageEntry {
  pageIndex: number
  topPx: number
  heightPx: number
  pageLayout: ReturnType<typeof resolveDocumentLayout>
  pageNodeSegments: DocumentPageNodeSegment[]
  blocks: LayoutBlock[]
}

interface PageMeasureEvent {
  heightPx: number
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<DocxViewerRootProps>(),
  { editable: true }
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
const paginationPlan = computed(() => {
  const model = props.controller.model
  if (!model) return { pages: [], pageNodeSegments: [] as DocumentPageNodeSegment[][] }

  const documentLayout = resolveDocumentLayout(model)
  const layoutResult = layoutDocument(model, {
    pageWidth: documentLayout.pageWidthPx,
    pageHeight: documentLayout.pageHeightPx,
    margin: documentLayout.marginsPx?.top ?? 72,
  })

  // Build page node segments from layout result
  const pageNodeSegments: DocumentPageNodeSegment[][] = layoutResult.map((page) =>
    page.blocks.map((block, idx) => ({
      nodeIndex: idx,
      ...(block.kind === "table"
        ? { tableRowRange: undefined }
        : {}),
    })) as DocumentPageNodeSegment[]
  )

  return { pages: layoutResult, pageNodeSegments }
})

const pages = computed(() => paginationPlan.value.pages)
const pageNodeSegments = computed(() => paginationPlan.value.pageNodeSegments)
const pageCount = computed(() => pages.value.length)

watch(pageCount, (count) => {
  emit("pageCountChange", count)
})

// ── Document layout ────────────────────────────────────────────────
const documentLayout = computed(() => {
  const model = props.controller.model
  if (!model) return { pageWidthPx: DEFAULT_DOC_PAGE_WIDTH, pageHeightPx: DEFAULT_DOC_PAGE_HEIGHT, marginsPx: { top: 72, bottom: 72, left: 72, right: 72 }, headerDistancePx: 0, footerDistancePx: 0 }
  return resolveDocumentLayout(model)
})

const pageWidthPx = computed(() => documentLayout.value.pageWidthPx)
const pageContentWidthPx = computed(
  () => documentLayout.value.pageWidthPx - (documentLayout.value.marginsPx?.left ?? 72) - (documentLayout.value.marginsPx?.right ?? 72)
)

// ── Theme / feature flags ──────────────────────────────────────────
const theme = computed(() => props.controller.documentTheme)
const trackedChangesEnabled = computed(() => props.controller.showTrackedChanges)
const commentsEnabled = computed(() => props.controller.showComments)

// ── Page heights ───────────────────────────────────────────────────
function getPageHeight(pageIndex: number): number {
  const measured = measuredPageHeights.value[pageIndex]
  if (Number.isFinite(measured) && (measured as number) > 0) {
    return measured as number
  }
  return documentLayout.value.pageHeightPx
}

// ── Page offsets ───────────────────────────────────────────────────
const pageOffsets = computed(() => {
  const offsets: number[] = []
  let offset = 0
  for (let i = 0; i < pageCount.value; i++) {
    offsets.push(offset)
    offset += getPageHeight(i) + DOC_PAGE_BREAK_GAP
  }
  return offsets
})

const totalHeightPx = computed(() => {
  if (pageCount.value === 0) return 0
  const lastOffset = pageOffsets.value[pageOffsets.value.length - 1]
  return lastOffset + getPageHeight(pageCount.value - 1) + DOC_PAGE_BREAK_GAP
})

// ── Visible page calculation ───────────────────────────────────────
const visiblePageEntries = computed<VisiblePageEntry[]>(() => {
  if (pageCount.value === 0) return []

  const viewportStart = Math.max(0, scrollTop.value - DOC_PAGE_BREAK_GAP * PAGE_OVERSCAN)
  const viewportEnd = viewportStart + viewportHeight.value + DOC_PAGE_BREAK_GAP * PAGE_OVERSCAN * 2

  const entries: VisiblePageEntry[] = []
  for (let i = 0; i < pageCount.value; i++) {
    const pageTop = pageOffsets.value[i]
    const pageHeight = getPageHeight(i)
    const pageBottom = pageTop + pageHeight

    if (pageBottom >= viewportStart && pageTop <= viewportEnd) {
      entries.push({
        pageIndex: i,
        topPx: pageTop,
        heightPx: pageHeight,
        pageLayout: documentLayout.value,
        pageNodeSegments: pageNodeSegments.value[i] ?? [],
        blocks: pages.value[i]?.blocks ?? [],
      })
    }
  }

  return entries
})

// ── Styles ─────────────────────────────────────────────────────────
const rootStyle = computed(() => ({
  flex: "1",
  overflowY: "auto" as const,
  overflowX: "hidden" as const,
  background: "#f3f4f6",
  padding: "24px 0",
  outline: isDragOver.value ? "3px dashed #2563eb" : "none",
}))

const spacerStyle = computed(() => ({
  height: `${Math.max(0, totalHeightPx.value)}px`,
  position: "relative" as const,
  width: "100%",
}))

function pageWrapperStyle(entry: VisiblePageEntry): Record<string, string | number> {
  return {
    position: "absolute",
    top: `${entry.topPx}px`,
    left: "50%",
    transform: "translateX(-50%)",
    width: `${pageWidthPx.value}px`,
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
  if (props.editable && hasDocxFile(event.dataTransfer)) {
    dragOverDepth.value++
    event.dataTransfer.dropEffect = "copy"
  }
}

function onDragLeave(): void {
  dragOverDepth.value = Math.max(0, dragOverDepth.value - 1)
}

async function onDrop(event: DragEvent): Promise<void> {
  dragOverDepth.value = 0
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
    viewportHeight.value = scrollContainerRef.value.clientHeight
  }
  const observer = new ResizeObserver((entries) => {
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
