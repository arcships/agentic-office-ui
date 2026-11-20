<template>
  <DocumentViewerSidebar
    class="docx-thumbnail-panel"
    :class="isDark ? 'docx-thumbnail-panel--dark' : undefined"
    data-testid="docx-thumbnail-panel"
    :open="!collapsed"
    :mode="mode"
    :breakpoint="breakpoint"
    :width="160"
    flush
    :label="m.docx.thumbnail.pages"
    :data-office-theme="isDark ? 'dark' : 'light'"
    data-office-density="comfortable"
    @update:open="onSidebarOpen"
  >
    <div
      ref="listRef"
      class="docx-thumbnail-list"
      role="listbox"
      :aria-label="m.docx.thumbnail.pages"
      :aria-activedescendant="activeDescendantId"
      :aria-busy="isLoading || undefined"
      :data-page-count="pageCount"
      tabindex="0"
      @keydown="onListKeydown"
      @scroll.passive="onThumbnailScroll"
    >
      <div
        v-if="isLoading"
        class="docx-thumbnail-skeleton"
        role="status"
        :aria-label="m.docx.thumbnail.loading"
        data-testid="docx-thumbnail-skeleton"
      >
        <div class="docx-thumbnail-skeleton__card">
          <span class="docx-thumbnail-skeleton__pulse" />
        </div>
        <div class="docx-thumbnail-skeleton__caption" />
      </div>
      <div v-else class="docx-thumbnail-track">
        <div
          v-if="topSpacerHeight > 0"
          class="docx-thumbnail-spacer"
          :style="{ height: `${topSpacerHeight}px` }"
          aria-hidden="true"
        />
        <div
          v-for="pageIndex of renderedPageRange"
          :key="pageIndex"
          :id="optionId(pageIndex)"
          class="docx-thumbnail-item"
          :class="{
            'docx-thumbnail-item--active': pageIndex + 1 === currentPage,
            'docx-thumbnail-item--blank': isParityBlankPage(pageIndex),
          }"
          role="option"
          data-testid="docx-thumbnail"
          :aria-current="pageIndex + 1 === currentPage ? 'page' : undefined"
          :aria-label="`${m.common.page.label} ${pageLabel(pageIndex)}`"
          :aria-posinset="pageIndex + 1"
          :aria-selected="pageIndex + 1 === currentPage"
          :aria-setsize="pageCount"
          :data-thumbnail-page-index="pageIndex"
          :data-thumbnail-page-label="pageLabel(pageIndex)"
          :data-thumbnail-parity-blank="isParityBlankPage(pageIndex) ? 'true' : 'false'"
          @click="onSelectPage(pageIndex)"
        >
          <FileThumbnail
            class="docx-thumbnail-card"
            :class="{ 'docx-thumbnail-card--active': pageIndex + 1 === currentPage }"
            size="92px"
            :file="thumbnailFile(pageIndex)"
            :preview-aspect-ratio="cardAspectRatio(pageIndex)"
            :is-loading="thumbnailLoading(pageIndex)"
            :has-error="thumbnailError(pageIndex)"
            preview-class-name="docx-thumbnail-preview"
            :alt="`${m.common.page.label} ${pageLabel(pageIndex)}`"
          >
            <template #preview>
              <canvas
                :ref="(el) => setCanvasRef(pageIndex, el as HTMLCanvasElement)"
                :width="thumbnailCanvasPixelWidth()"
                :height="thumbnailCanvasPixelHeight(pageIndex)"
                :style="thumbnailCanvasStyle(pageIndex)"
                class="docx-thumbnail-canvas"
              />
            </template>
          </FileThumbnail>
          <span class="docx-thumbnail-label" data-testid="docx-thumbnail-label">
            {{ pageLabel(pageIndex) }}
          </span>
        </div>
        <div
          v-if="bottomSpacerHeight > 0"
          class="docx-thumbnail-spacer"
          :style="{ height: `${bottomSpacerHeight}px` }"
          aria-hidden="true"
        />
      </div>
    </div>
  </DocumentViewerSidebar>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, onMounted, onUnmounted, watch } from "vue"
import type {
  DocModel,
  DocNode,
  DocxEditorController,
  DocxSectionPageNumberTracker,
  DocumentNoteDefinition,
  LayoutOptions,
  LayoutSnapshot,
  LayoutSnapshotPage,
  LayoutSnapshotParagraphFragment,
  LayoutSnapshotRun,
  LayoutSnapshotTableFragment,
} from "@arcships/docx-core"
import {
  buildLayoutSnapshot,
  createSectionPageNumberTracker,
  nodeReferencedNoteIds,
  parseSectionLayout,
  resolveDocumentBackgroundColor,
  resolveDocumentLayout,
  resolveDocumentSectionsFromMetadata,
  resolveSectionPropertiesXmlForNodeIndex,
  selectSectionVariantForPage,
  pagePatternPaintStyle,
  paintPagePatternOnCanvas,
} from "@arcships/docx-core"
import {
  ensureDocxViewerPageSurfaceRegistry,
  subscribeDocxViewerPageSurfaces,
} from "../composables/page-surface-registry"
import { DocumentViewerSidebar, FileThumbnail } from "@arcships/vue-ui"
import { useDocxMessages } from "../i18n"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    controller?: DocxEditorController
    model?: DocModel
    layoutOptions?: LayoutOptions
    currentPage?: number
    totalPages?: number
    defaultCollapsed?: boolean
    showToggle?: boolean
    isDark?: boolean
    /**
     * Layout strategy forwarded to `DocumentViewerSidebar`. `auto` resolves to
     * the upstream 768px inline breakpoint.
     */
    mode?: "auto" | "inline" | "overlay"
    /** Container width at which `auto` switches to the inline rail. */
    breakpoint?: number
  }>(),
  {
    currentPage: 1,
    totalPages: 1,
    defaultCollapsed: false,
    showToggle: true,
    isDark: false,
    mode: "auto",
    breakpoint: 768,
  }
)

const m = useDocxMessages()

const emit = defineEmits<{
  selectPage: [pageIndex: number]
}>()

// ── State ──────────────────────────────────────────────────────────
const collapsed = ref(props.defaultCollapsed)
const visibleStartPage = ref(0)
const listRef = ref<HTMLElement | undefined>()
const registryRevision = ref(0)
const listboxId = `docx-thumbnails-${Math.random().toString(36).slice(2, 8)}`

/**
 * When the panel is paired with an editor controller, the viewer owns
 * pagination. Subscribe to its registry instead of running another layout pass
 * on every model change.
 */
watch(
  () => props.controller,
  (controller, _previous, onCleanup) => {
    if (!controller) return
    registryRevision.value += 1
    const unsubscribe = subscribeDocxViewerPageSurfaces(controller, () => {
      registryRevision.value += 1
    })
    onCleanup(unsubscribe)
  },
  { immediate: true }
)

const resolvedModel = computed(() => props.model ?? props.controller?.model)
const documentLayout = computed(() => {
  const model = resolvedModel.value
  return model ? resolveDocumentLayout(model) : undefined
})
const documentBackgroundColor = computed(() =>
  resolveDocumentBackgroundColor(resolvedModel.value?.metadata.documentBackgroundColor)
)
const documentBackgroundPattern = computed(
  () => resolvedModel.value?.metadata.documentBackgroundPattern
)
const documentPagePatternStyle = computed(() =>
  pagePatternPaintStyle(documentBackgroundPattern.value)
)
const resolvedLayoutOptions = computed<LayoutOptions>(() => ({
  pageWidth: documentLayout.value?.pageWidthPx ?? 816,
  pageHeight: documentLayout.value?.pageHeightPx ?? 1056,
  margin: documentLayout.value?.marginsPx.top ?? 72,
  ...props.layoutOptions,
}))
const sharedSnapshot = computed<LayoutSnapshot | undefined>(() => {
  // Track layout publications from the viewer.
  registryRevision.value
  const controller = props.controller
  return controller ? ensureDocxViewerPageSurfaceRegistry(controller).layoutSnapshot : undefined
})
const layoutSnapshot = computed<LayoutSnapshot | undefined>(() => {
  if (sharedSnapshot.value) return sharedSnapshot.value
  // A controller without a published viewer snapshot must not duplicate the
  // viewer's pagination work; wait for the shared snapshot instead.
  if (props.controller) return undefined
  const model = resolvedModel.value
  return model ? buildLayoutSnapshot(model, resolvedLayoutOptions.value) : undefined
})
const sharedPageSizes = computed(() => {
  registryRevision.value
  const controller = props.controller
  return controller
    ? ensureDocxViewerPageSurfaceRegistry(controller).layoutPageSizes
    : undefined
})
const pageCount = computed(() => {
  if (props.controller) {
    registryRevision.value
    const count = ensureDocxViewerPageSurfaceRegistry(props.controller).layoutPageCount
    if (count > 0) return count
  }
  return layoutSnapshot.value?.pages.length ?? Math.max(0, props.totalPages ?? 0)
})

interface PagePresentation {
  headerNodes: DocNode[]
  footerNodes: DocNode[]
  footnotes: DocumentNoteDefinition[]
  endnotes: DocumentNoteDefinition[]
}

/**
 * Per-section page numbering for the rail's labels.
 *
 * The viewer publishes its tracker so the rail shows exactly the numbers the
 * pages themselves show; a standalone panel (no controller) resolves the same
 * tracker from its own pagination.
 */
const pageNumberTracker = computed<DocxSectionPageNumberTracker | undefined>(() => {
  const model = resolvedModel.value
  const snapshot = layoutSnapshot.value
  if (!model || !snapshot) return undefined
  registryRevision.value
  if (props.controller) {
    const shared = ensureDocxViewerPageSurfaceRegistry(props.controller).pageNumberTracker
    if (shared) return shared
  }
  return createSectionPageNumberTracker({
    sections: resolveDocumentSectionsFromMetadata(model.metadata).map((section) => ({
      startNodeIndex: section.startNodeIndex,
      // A section without its own `w:sectPr` inherits the document's.
      sectionPropertiesXml:
        section.sectionPropertiesXml ?? model.metadata.sectionPropertiesXml,
    })),
    pages: snapshot.pages.map((page) => ({ startNodeIndex: pageFirstNodeIndex(page) })),
    nodeCount: model.nodes.length,
  })
})

function pageLabel(pageIndex: number): string {
  return pageNumberTracker.value?.pageLabelForPage(pageIndex) ?? String(pageIndex + 1)
}

function isParityBlankPage(pageIndex: number): boolean {
  return pageNumberTracker.value?.isParityBlankPage(pageIndex) ?? false
}

const pagePresentations = computed<PagePresentation[]>(() => {
  const model = resolvedModel.value
  const snapshot = layoutSnapshot.value
  if (!model || !snapshot) return []

  const metadata = model.metadata
  const sections = resolveDocumentSectionsFromMetadata(metadata)
  const evenAndOddHeaders = metadata.compatibility?.evenAndOddHeaders ?? false
  const footnotesById = new Map((metadata.footnotes ?? []).map((note) => [note.id, note]))
  const firstPageBySection = new Map<number, number>()
  const tracker = pageNumberTracker.value

  return snapshot.pages.map((page, pageIndex) => {
    const sectionIndex = tracker?.sectionIndexForPage(pageIndex) ?? 0
    if (!firstPageBySection.has(sectionIndex)) {
      firstPageBySection.set(sectionIndex, pageIndex)
    }

    const section = sections[sectionIndex]
    const sectionPageIndex =
      pageIndex - (firstPageBySection.get(sectionIndex) ?? pageIndex)
    const sectionPropertiesXml =
      section?.sectionPropertiesXml ?? metadata.sectionPropertiesXml
    const headerSection = section
      ? selectSectionVariantForPage(
          section.headerSections,
          sectionPropertiesXml,
          sectionPageIndex,
          { evenAndOddHeaders, documentPageIndex: pageIndex }
        )
      : undefined
    const footerSection = section
      ? selectSectionVariantForPage(
          section.footerSections,
          sectionPropertiesXml,
          sectionPageIndex,
          { evenAndOddHeaders, documentPageIndex: pageIndex }
        )
      : undefined

    const referencedNoteIds: number[] = []
    const seenNoteIds = new Set<number>()
    for (const block of page.blocks) {
      const nodeIndex = block.source?.nodeIndex
      if (nodeIndex === undefined || nodeIndex < 0) continue
      const node = model.nodes[nodeIndex]
      if (!node) continue
      for (const id of nodeReferencedNoteIds(node, "footnote")) {
        if (!seenNoteIds.has(id)) {
          seenNoteIds.add(id)
          referencedNoteIds.push(id)
        }
      }
    }

    return {
      headerNodes: headerSection?.nodes ?? [],
      footerNodes: footerSection?.nodes ?? [],
      footnotes: referencedNoteIds
        .map((id) => footnotesById.get(id))
        .filter((note): note is DocumentNoteDefinition => Boolean(note)),
      endnotes:
        pageIndex === snapshot.pages.length - 1
          ? [...(metadata.endnotes ?? [])]
          : [],
    }
  })
})

// ── Thumbnail dimensions ───────────────────────────────────────────
/**
 * Upstream DOCX thumbnail geometry: a `FileThumbnail` card with `w-[92px]`
 * inside an option box with `p-2` (8px), `gap-2` (8px) and one `text-xs` line
 * (16px); every row reserves `pb-3` (12px) for the next card, and the list pads
 * the virtual track with 16px on both edges.
 */
const THUMBNAIL_WIDTH = 92
const THUMBNAIL_OPTION_PADDING = 8
const THUMBNAIL_OPTION_GAP = 8
const THUMBNAIL_LABEL_HEIGHT = 16
const THUMBNAIL_ROW_GAP = 12
const THUMBNAIL_LIST_PADDING = 16
const THUMBNAIL_ITEM_EXTRA_HEIGHT =
  THUMBNAIL_OPTION_PADDING * 2 +
  THUMBNAIL_OPTION_GAP +
  THUMBNAIL_LABEL_HEIGHT +
  THUMBNAIL_ROW_GAP
const DEFAULT_PAGE_SIZE = { width: 816, height: 1056 }
const DOCX_THUMBNAIL_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

interface ThumbnailPageMetric {
  width: number
  height: number
  scale: number
  /** Offset of this item in the rail's scroll space. */
  top: number
}

function pageFirstNodeIndex(page: LayoutSnapshotPage): number | undefined {
  for (const block of page.blocks) {
    const nodeIndex = block.source?.nodeIndex
    if (Number.isFinite(nodeIndex) && (nodeIndex as number) >= 0) return Math.round(nodeIndex as number)
  }
  return undefined
}

// When the panel owns pagination (the viewer has no controller), resolve each
// page's section layout so mixed portrait/landscape documents keep one aspect
// ratio per page instead of squashing every page to the first scheme.
const derivedPageLayouts = computed(() => {
  const snapshot = layoutSnapshot.value
  const model = resolvedModel.value
  if (!snapshot || !model) return undefined
  const sections = resolveDocumentSectionsFromMetadata(model.metadata)
  const layouts = new Map<number, ReturnType<typeof parseSectionLayout>>()
  let previousNodeIndex = 0
  for (const page of snapshot.pages) {
    const firstNodeIndex = pageFirstNodeIndex(page) ?? previousNodeIndex
    previousNodeIndex = firstNodeIndex
    const sectionPropertiesXml = resolveSectionPropertiesXmlForNodeIndex(
      sections,
      firstNodeIndex,
      model.metadata.sectionPropertiesXml,
    )
    if (!sectionPropertiesXml) continue
    layouts.set(page.index, parseSectionLayout(sectionPropertiesXml))
  }
  return layouts
})

function pageSizeFor(pageIndex: number): { width: number; height: number } {
  const layoutSize = sharedPageSizes.value?.get(pageIndex)
  if (layoutSize && layoutSize.widthPx > 0 && layoutSize.heightPx > 0) {
    return { width: layoutSize.widthPx, height: layoutSize.heightPx }
  }
  const derived = derivedPageLayouts.value?.get(pageIndex)
  if (derived && derived.pageWidthPx > 0 && derived.pageHeightPx > 0) {
    return { width: derived.pageWidthPx, height: derived.pageHeightPx }
  }
  const pageSize = layoutSnapshot.value?.pages[pageIndex]?.pageSizePx
  if (pageSize && pageSize.width > 0 && pageSize.height > 0) return pageSize
  const fallback = documentLayout.value
  if (fallback && fallback.pageWidthPx > 0 && fallback.pageHeightPx > 0) {
    return { width: fallback.pageWidthPx, height: fallback.pageHeightPx }
  }
  return DEFAULT_PAGE_SIZE
}

const pageMetrics = computed<ThumbnailPageMetric[]>(() => {
  const metrics: ThumbnailPageMetric[] = []
  let top = 0
  for (let pageIndex = 0; pageIndex < pageCount.value; pageIndex++) {
    const pageSize = pageSizeFor(pageIndex)
    const height = Math.max(1, Math.round(THUMBNAIL_WIDTH * (pageSize.height / pageSize.width)))
    metrics.push({
      width: THUMBNAIL_WIDTH,
      height,
      scale: THUMBNAIL_WIDTH / pageSize.width,
      top,
    })
    top += height + THUMBNAIL_ITEM_EXTRA_HEIGHT
  }
  return metrics
})
const railContentHeight = computed(() => {
  const metrics = pageMetrics.value
  if (metrics.length === 0) return 0
  const last = metrics[metrics.length - 1]
  return last.top + last.height + THUMBNAIL_ITEM_EXTRA_HEIGHT
})

function readDevicePixelRatio(): number {
  return typeof window === "undefined"
    ? 1
    : Math.min(3, Math.max(1, window.devicePixelRatio || 1))
}

const devicePixelRatio = ref(readDevicePixelRatio())
let dprMediaQuery: MediaQueryList | undefined

function bindDevicePixelRatioListener(): void {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") return
  dprMediaQuery?.removeEventListener?.("change", syncDevicePixelRatio)
  dprMediaQuery = window.matchMedia(
    `(resolution: ${window.devicePixelRatio || 1}dppx)`
  )
  dprMediaQuery.addEventListener?.("change", syncDevicePixelRatio)
}

function syncDevicePixelRatio(): void {
  devicePixelRatio.value = readDevicePixelRatio()
  bindDevicePixelRatioListener()
}

onMounted(() => {
  if (typeof window === "undefined") return
  window.addEventListener("resize", syncDevicePixelRatio)
  bindDevicePixelRatioListener()
})

onUnmounted(() => {
  if (typeof window === "undefined") return
  window.removeEventListener("resize", syncDevicePixelRatio)
  dprMediaQuery?.removeEventListener?.("change", syncDevicePixelRatio)
  dprMediaQuery = undefined
})

function pageMetric(pageIndex: number): ThumbnailPageMetric {
  const fallbackHeight = Math.max(
    1,
    Math.round(THUMBNAIL_WIDTH * (DEFAULT_PAGE_SIZE.height / DEFAULT_PAGE_SIZE.width)),
  )
  return pageMetrics.value[pageIndex] ?? {
    width: THUMBNAIL_WIDTH,
    height: fallbackHeight,
    scale: THUMBNAIL_WIDTH / DEFAULT_PAGE_SIZE.width,
    top: pageIndex * (fallbackHeight + THUMBNAIL_ITEM_EXTRA_HEIGHT),
  }
}

function thumbnailCanvasWrapperStyle(pageIndex: number): Record<string, string> {
  const metric = pageMetric(pageIndex)
  return { width: `${metric.width}px`, height: `${metric.height}px` }
}

function thumbnailCanvasStyle(pageIndex: number): Record<string, string> {
  return thumbnailCanvasWrapperStyle(pageIndex)
}

function thumbnailCanvasPixelWidth(): number {
  return Math.max(1, Math.round(THUMBNAIL_WIDTH * devicePixelRatio.value))
}

function thumbnailCanvasPixelHeight(pageIndex: number): number {
  return Math.max(1, Math.round(pageMetric(pageIndex).height * devicePixelRatio.value))
}

// ── Upstream option/card helpers ───────────────────────────────────
const isLoading = computed(() => pageCount.value < 1)

function optionId(pageIndex: number): string {
  return `${listboxId}-page-${pageIndex + 1}`
}

const activeDescendantId = computed(() => {
  if (pageCount.value < 1) return undefined
  const pageIndex = Math.min(pageCount.value - 1, Math.max(0, (props.currentPage ?? 1) - 1))
  return optionId(pageIndex)
})

function cardAspectRatio(pageIndex: number): number {
  const pageSize = pageSizeFor(pageIndex)
  return pageSize.width / pageSize.height
}

function thumbnailFile(pageIndex: number) {
  return {
    name: `${m.value.common.page.label} ${pageLabel(pageIndex)}`,
    type: DOCX_THUMBNAIL_MIME,
  }
}

function thumbnailLoading(pageIndex: number): boolean {
  return thumbnailStates.value[pageIndex] !== "ready"
}

function thumbnailError(pageIndex: number): boolean {
  return thumbnailStates.value[pageIndex] === "error"
}

function setThumbnailState(pageIndex: number, state: string): void {
  if (thumbnailStates.value[pageIndex] === state) return
  thumbnailStates.value = { ...thumbnailStates.value, [pageIndex]: state }
}

// ── Page range for rendering ───────────────────────────────────────
const VISIBLE_COUNT = 20
const renderWindow = computed(() => {
  const total = pageCount.value
  if (total <= 0) return { start: 0, end: 0 }
  const firstVisible = Math.min(Math.max(0, visibleStartPage.value), total - 1)
  const start = Math.max(0, firstVisible - 5)
  return { start, end: Math.min(total, start + VISIBLE_COUNT) }
})
const renderedPageRange = computed(() => {
  const range: number[] = []
  for (let i = renderWindow.value.start; i < renderWindow.value.end; i++) {
    range.push(i)
  }
  return range
})
const topSpacerHeight = computed(() => pageMetrics.value[renderWindow.value.start]?.top ?? 0)
const bottomSpacerHeight = computed(() => {
  const window = renderWindow.value
  const metrics = pageMetrics.value
  if (window.end >= metrics.length) return 0
  return Math.max(0, railContentHeight.value - (metrics[window.end]?.top ?? railContentHeight.value))
})

// ── Canvas refs ────────────────────────────────────────────────────
const canvasRefs = new Map<number, HTMLCanvasElement>()
/** Reactive mirror of each canvas' `data-thumbnail-state` for the card overlay. */
const thumbnailStates = ref<Record<number, string>>({})

function setCanvasRef(pageIndex: number, el: HTMLCanvasElement | null): void {
  if (!el) {
    canvasRefs.delete(pageIndex)
    if (pageIndex in thumbnailStates.value) {
      const next = { ...thumbnailStates.value }
      delete next[pageIndex]
      thumbnailStates.value = next
    }
    return
  }
  if (canvasRefs.get(pageIndex) === el) return
  canvasRefs.set(pageIndex, el)
  scheduleThumbnailRender()
}

// ── Thumbnail rendering ────────────────────────────────────────────
let renderScheduled = false

function scheduleThumbnailRender(): void {
  if (renderScheduled) return
  renderScheduled = true
  void nextTick(() => {
    renderScheduled = false
    void renderVisibleThumbnails()
  })
}

watch(
  [
    renderedPageRange,
    layoutSnapshot,
    pageMetrics,
    () => props.controller?.documentLoadNonce,
    devicePixelRatio,
  ],
  scheduleThumbnailRender,
  { immediate: true, flush: "post" }
)

// Keep the rail in sync with the page the main viewer is showing. Without this
// the rail stays at page 1 and never renders the active page's window.
watch(
  [() => props.currentPage, pageCount],
  () => { syncRailToCurrentPage() },
  { flush: "post", immediate: true }
)

let railScrollIsProgrammatic = false

function syncRailToCurrentPage(): void {
  const total = pageCount.value
  if (total <= 0) return
  const pageIndex = Math.min(total - 1, Math.max(0, (props.currentPage ?? 1) - 1))
  const metric = pageMetrics.value[pageIndex]
  if (!metric) return
  const desiredWindowStart = Math.max(0, pageIndex - 5)
  if (visibleStartPage.value !== desiredWindowStart) {
    visibleStartPage.value = desiredWindowStart
  }
  const list = listRef.value
  if (!list || typeof list.scrollTo !== "function") return
  const viewport = typeof list.clientHeight === "number" && list.clientHeight > 0
    ? list.clientHeight
    : metric.height
  const target = Math.max(0, metric.top - Math.max(0, (viewport - metric.height) / 2))
  if (Math.abs((list.scrollTop ?? 0) - target) <= 1) return
  railScrollIsProgrammatic = true
  list.scrollTo({ top: target, behavior: "auto" })
  void nextTick(() => { railScrollIsProgrammatic = false })
}

type ThumbnailTextRun = Extract<LayoutSnapshotRun, { kind: "text" }>
type ThumbnailImageRun = Extract<LayoutSnapshotRun, { kind: "image" }>
type ThumbnailImageState =
  | { status: "loading" }
  | { status: "error" }
  | { status: "loaded"; image: HTMLImageElement }

const imageStates = new Map<string, ThumbnailImageState>()

function requestThumbnailImage(src: string): ThumbnailImageState {
  const existing = imageStates.get(src)
  if (existing) return existing
  if (typeof Image === "undefined") {
    const unavailable: ThumbnailImageState = { status: "error" }
    imageStates.set(src, unavailable)
    return unavailable
  }

  const loading: ThumbnailImageState = { status: "loading" }
  imageStates.set(src, loading)
  const image = new Image()
  image.onload = () => {
    imageStates.set(src, { status: "loaded", image })
    scheduleThumbnailRender()
  }
  image.onerror = () => {
    imageStates.set(src, { status: "error" })
    scheduleThumbnailRender()
  }
  image.src = src
  return loading
}

// Scale of the page currently being rasterized. Pages can mix portrait and
// landscape sections, so each thumbnail carries its own scale.
let renderScale = 1

function canvasColor(value: string | undefined, fallback: string): string {
  const color = value?.trim()
  if (!color || color.toLowerCase() === "auto") return fallback
  return /^[0-9a-f]{6}$/i.test(color) ? `#${color}` : color
}

function headingScale(level: LayoutSnapshotParagraphFragment["headingLevel"]): number {
  if (!level) return 1
  return ({ 1: 2.15, 2: 1.75, 3: 1.45, 4: 1.28, 5: 1.15, 6: 1.05 } as const)[level]
}

function runFontSize(
  run: ThumbnailTextRun,
  headingLevel?: LayoutSnapshotParagraphFragment["headingLevel"]
): number {
  const naturalSize = (run.style?.fontSizePt ?? 12) * (4 / 3) * headingScale(headingLevel)
  return Math.max(naturalSize, 3.5 / renderScale)
}

function applyRunStyle(
  ctx: CanvasRenderingContext2D,
  run: ThumbnailTextRun,
  headingLevel?: LayoutSnapshotParagraphFragment["headingLevel"]
): number {
  const size = runFontSize(run, headingLevel)
  const style = run.style?.italic ? "italic " : ""
  const weight = run.style?.bold || headingLevel ? "700 " : ""
  ctx.font = `${style}${weight}${size}px ${run.style?.fontFamily || "Calibri, Arial, sans-serif"}`
  ctx.fillStyle = canvasColor(run.style?.color, "#374151")
  return size
}

function drawTextRuns(
  ctx: CanvasRenderingContext2D,
  runs: readonly ThumbnailTextRun[],
  frame: { x: number; y: number; width: number; height: number },
  headingLevel?: LayoutSnapshotParagraphFragment["headingLevel"],
  maxLines?: number
): void {
  if (runs.length === 0 || frame.width <= 0 || frame.height <= 0) return
  const largestFont = runs.reduce(
    (largest, run) => Math.max(largest, runFontSize(run, headingLevel)),
    12
  )
  const lineHeight = Math.max(16, largestFont * 1.2)
  const resolvedMaxLines = Math.max(
    1,
    Number.isFinite(maxLines)
      ? Math.floor(maxLines as number)
      : Math.floor(frame.height / lineHeight)
  )
  let lineIndex = 0
  let cursorX = frame.x
  const right = frame.x + frame.width

  ctx.save()
  ctx.beginPath()
  ctx.rect(frame.x, frame.y, frame.width, frame.height)
  ctx.clip()

  outer: for (const run of runs) {
    const tokens = run.text.match(/\r\n|\n|\t|[^\S\r\n\t]+|[^\s\r\n\t]+/g) ?? []
    for (const rawToken of tokens) {
      if (rawToken === "\r\n" || rawToken === "\n") {
        lineIndex++
        cursorX = frame.x
        if (lineIndex >= resolvedMaxLines) break outer
        continue
      }
      const token = rawToken === "\t" ? "    " : rawToken
      const fontSize = applyRunStyle(ctx, run, headingLevel)
      const tokenWidth = ctx.measureText(token).width
      if (cursorX > frame.x && cursorX + tokenWidth > right && token.trim()) {
        lineIndex++
        cursorX = frame.x
        if (lineIndex >= resolvedMaxLines) break outer
      }
      if (cursorX === frame.x && !token.trim()) continue
      const baseline = frame.y + lineIndex * lineHeight + Math.min(lineHeight * 0.8, fontSize)
      ctx.fillText(token, cursorX, baseline)
      cursorX += tokenWidth
    }
  }
  ctx.restore()
}

function pageMargins(page: LayoutSnapshotPage): { top: number; right: number; bottom: number; left: number } {
  const explicitMargin = props.layoutOptions?.margin
  if (Number.isFinite(explicitMargin)) {
    const margin = Math.max(0, Number(explicitMargin))
    return { top: margin, right: margin, bottom: margin, left: margin }
  }
  const derived = derivedPageLayouts.value?.get(page.index)
  return derived?.marginsPx ?? page.marginsPx
}

function contentFrame(
  page: LayoutSnapshotPage,
  sourceFrame: { y: number; height: number }
): { x: number; y: number; width: number; height: number } {
  const margins = pageMargins(page)
  return {
    x: margins.left,
    y: margins.top + sourceFrame.y - page.marginsPx.top,
    width: Math.max(1, page.pageSizePx.width - margins.left - margins.right),
    height: Math.max(1, sourceFrame.height),
  }
}

function drawImagePlaceholder(
  ctx: CanvasRenderingContext2D,
  frame: { x: number; y: number; width: number; height: number }
): void {
  ctx.fillStyle = "#e5e7eb"
  ctx.fillRect(frame.x, frame.y, frame.width, frame.height)
  ctx.strokeStyle = "#9ca3af"
  ctx.lineWidth = Math.max(1, 0.75 / renderScale)
  ctx.strokeRect(frame.x, frame.y, frame.width, frame.height)
  ctx.beginPath()
  ctx.moveTo(frame.x, frame.y)
  ctx.lineTo(frame.x + frame.width, frame.y + frame.height)
  ctx.moveTo(frame.x + frame.width, frame.y)
  ctx.lineTo(frame.x, frame.y + frame.height)
  ctx.stroke()
}

function drawImageRun(
  ctx: CanvasRenderingContext2D,
  image: ThumbnailImageRun,
  frame: { x: number; y: number; width: number; height: number }
): void {
  const maxWidth = Math.max(1, Math.min(frame.width, image.widthPx ?? frame.width))
  const maxHeight = Math.max(1, Math.min(frame.height, image.heightPx ?? frame.height))
  const src = image.src?.trim()
  if (src) {
    const state = requestThumbnailImage(src)
    if (state.status === "loaded") {
      const naturalWidth = state.image.naturalWidth || state.image.width || maxWidth
      const naturalHeight = state.image.naturalHeight || state.image.height || maxHeight
      if (naturalWidth > 0 && naturalHeight > 0) {
        const scale = Math.min(maxWidth / naturalWidth, maxHeight / naturalHeight)
        const width = Math.max(1, naturalWidth * scale)
        const height = Math.max(1, naturalHeight * scale)
        const x = frame.x + Math.max(0, (frame.width - width) / 2)
        const y = frame.y + Math.max(0, (frame.height - height) / 2)
        try {
          ctx.drawImage(state.image, x, y, width, height)
          return
        } catch {
          // Fall through to the placeholder for sources the canvas rejects.
        }
      }
    }
  }
  drawImagePlaceholder(ctx, {
    x: frame.x,
    y: frame.y,
    width: maxWidth,
    height: maxHeight,
  })
}

function drawParagraph(
  ctx: CanvasRenderingContext2D,
  block: LayoutSnapshotParagraphFragment,
  page: LayoutSnapshotPage
): void {
  const frame = contentFrame(page, block.framePx)
  const textRuns = block.runs.filter((run): run is ThumbnailTextRun => run.kind === "text")
  drawTextRuns(ctx, textRuns, frame, block.headingLevel)
  for (const run of block.runs) {
    if (run.kind === "image") drawImageRun(ctx, run, frame)
  }
}

function drawTable(
  ctx: CanvasRenderingContext2D,
  block: LayoutSnapshotTableFragment,
  page: LayoutSnapshotPage
): void {
  const frame = contentFrame(page, block.framePx)
  const rows = block.rows
  const rowHeight = frame.height / Math.max(1, rows.length)
  const lineWidth = Math.max(1, 0.75 / renderScale)

  if (rows.length === 0) {
    ctx.strokeStyle = "#9ca3af"
    ctx.lineWidth = lineWidth
    ctx.strokeRect(frame.x, frame.y, frame.width, frame.height)
    return
  }

  rows.forEach((row, rowIndex) => {
    const columnCount = Math.max(
      1,
      row.cells.reduce((sum, cell) => sum + Math.max(1, cell.colSpan ?? 1), 0)
    )
    let columnCursor = 0
    row.cells.forEach((cell) => {
      const colSpan = Math.max(1, cell.colSpan ?? 1)
      const cellX = frame.x + frame.width * (columnCursor / columnCount)
      const cellWidth = frame.width * (colSpan / columnCount)
      const cellY = frame.y + rowIndex * rowHeight
      ctx.fillStyle = canvasColor(cell.backgroundColor ?? row.backgroundColor, "#ffffff")
      ctx.fillRect(cellX, cellY, cellWidth, rowHeight)
      ctx.strokeStyle = "#9ca3af"
      ctx.lineWidth = lineWidth
      ctx.strokeRect(cellX, cellY, cellWidth, rowHeight)

      const padding = Math.max(4, 1.5 / renderScale)
      const textRuns = cell.paragraphs.flatMap((paragraph) =>
        paragraph.runs.filter((run): run is ThumbnailTextRun => run.kind === "text")
      )
      const cellFrame = {
        x: cellX + padding,
        y: cellY + padding,
        width: Math.max(1, cellWidth - padding * 2),
        height: Math.max(1, rowHeight - padding * 2),
      }
      drawTextRuns(ctx, textRuns, cellFrame)
      for (const paragraph of cell.paragraphs) {
        for (const run of paragraph.runs) {
          if (run.kind === "image") drawImageRun(ctx, run, cellFrame)
        }
      }
      columnCursor += colSpan
    })
  })
}

function paragraphPlainText(
  paragraph: Extract<DocNode, { type: "paragraph" }>
): string {
  return paragraph.children
    .map((child) => (child.type === "text" ? child.text : ""))
    .join("")
}

function collectNodeTextLines(nodes: readonly DocNode[]): string[] {
  const lines: string[] = []
  for (const node of nodes) {
    if (node.type === "paragraph") {
      const text = paragraphPlainText(node).trim()
      if (text) lines.push(text)
      continue
    }
    for (const row of node.rows) {
      const cells = row.cells
        .map((cell) => collectNodeTextLines(cell.nodes).join(" "))
        .filter((text) => text.trim())
      if (cells.length > 0) lines.push(cells.join(" | "))
    }
  }
  return lines
}

function textRunsFromLines(
  lines: readonly string[],
  fontSizePt: number
): ThumbnailTextRun[] {
  return lines.map((text, index) => ({
    kind: "text" as const,
    id: `line-${index}`,
    text: index === 0 ? text : `\n${text}`,
    style: { fontSizePt },
  }))
}

function drawPageChrome(
  ctx: CanvasRenderingContext2D,
  page: LayoutSnapshotPage,
  presentation: PagePresentation | undefined
): void {
  if (!presentation) return
  const margins = pageMargins(page)
  const width = Math.max(1, page.pageSizePx.width - margins.left - margins.right)
  const pageHeight = page.pageSizePx.height
  const chromeHeight = Math.max(8, margins.top * 0.7)

  const headerRuns = textRunsFromLines(
    collectNodeTextLines(presentation.headerNodes),
    9
  )
  if (headerRuns.length > 0) {
    drawTextRuns(ctx, headerRuns, {
      x: margins.left,
      y: Math.max(1, margins.top * 0.12),
      width,
      height: chromeHeight,
    })
  }

  const footerRuns = textRunsFromLines(
    collectNodeTextLines(presentation.footerNodes),
    9
  )
  const footerHeight = Math.max(8, margins.bottom * 0.7)
  const footerY = Math.max(
    margins.top,
    pageHeight - margins.bottom * 0.85 - footerHeight * 0.2
  )
  if (footerRuns.length > 0) {
    drawTextRuns(ctx, footerRuns, {
      x: margins.left,
      y: footerY,
      width,
      height: footerHeight,
    })
  }

  const noteText = (
    note: DocumentNoteDefinition,
    index: number
  ): string => {
    const text =
      note.text?.trim() || collectNodeTextLines(note.nodes ?? []).join(" ").trim()
    return text ? `[${index + 1}] ${text}` : ""
  }
  const noteLines = [
    ...presentation.footnotes.map(noteText),
    ...presentation.endnotes.map(noteText),
  ].filter((line) => line.trim().length > 0)
  if (noteLines.length === 0) return

  const bodyBottom = page.blocks.reduce((bottom, block) => {
    const frame = contentFrame(page, block.framePx)
    return Math.max(bottom, frame.y + frame.height)
  }, margins.top)
  const noteTop = Math.min(bodyBottom + 8, footerY - 8)
  const noteBottom = footerRuns.length > 0 ? footerY - 2 : pageHeight - margins.bottom * 0.15
  const noteHeight = Math.max(1, noteBottom - noteTop)
  if (noteHeight < 6) return
  drawTextRuns(ctx, textRunsFromLines(noteLines, 7), {
    x: margins.left,
    y: noteTop,
    width,
    height: noteHeight,
  })
}

function renderVisibleThumbnails(): void {
  const snapshot = layoutSnapshot.value
  if (!snapshot) return

  const dpr = devicePixelRatio.value

  for (const pageIndex of renderedPageRange.value) {
    const canvas = canvasRefs.get(pageIndex)
    if (!canvas) continue

    const metric = pageMetrics.value[pageIndex]
    if (!metric) continue
    const page = snapshot.pages[pageIndex]
    if (!page) {
      canvas.dataset.thumbnailState = "unavailable"
      canvas.dataset.thumbnailContentBlocks = "0"
      setThumbnailState(pageIndex, "unavailable")
      continue
    }

    const ctx = canvas.getContext("2d")
    if (!ctx) {
      canvas.dataset.thumbnailState = "error"
      setThumbnailState(pageIndex, "error")
      continue
    }
    canvas.dataset.thumbnailState = "rendering"

    try {
      renderScale = metric.scale
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      ctx.clearRect(0, 0, canvas.width, canvas.height)
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
      ctx.fillStyle =
        documentPagePatternStyle.value.backgroundColor ??
        documentBackgroundColor.value ??
        "#ffffff"
      ctx.fillRect(0, 0, metric.width, metric.height)
      paintPagePatternOnCanvas(ctx, documentBackgroundPattern.value, {
        widthPx: metric.width,
        heightPx: metric.height,
      })
      ctx.strokeStyle = "#d4d4d4"
      ctx.lineWidth = 1
      ctx.strokeRect(0.5, 0.5, metric.width - 1, metric.height - 1)

      ctx.save()
      ctx.scale(metric.scale, metric.scale)
      for (const block of page.blocks) {
        if (block.kind === "paragraph") drawParagraph(ctx, block, page)
        else drawTable(ctx, block, page)
      }
      drawPageChrome(ctx, page, pagePresentations.value[pageIndex])
      ctx.restore()

      canvas.dataset.thumbnailState = "ready"
      canvas.dataset.thumbnailContentBlocks = String(page.blocks.length)
      setThumbnailState(pageIndex, "ready")
    } catch {
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      canvas.dataset.thumbnailState = "error"
      canvas.dataset.thumbnailContentBlocks = "0"
      setThumbnailState(pageIndex, "error")
    }
  }
}

// ── Scroll handling ────────────────────────────────────────────────
function pageIndexAtScrollTop(scrollTop: number): number {
  const metrics = pageMetrics.value
  if (metrics.length === 0) return 0
  let low = 0
  let high = metrics.length - 1
  while (low < high) {
    const middle = Math.floor((low + high + 1) / 2)
    if (metrics[middle].top <= scrollTop) low = middle
    else high = middle - 1
  }
  return low
}

function onThumbnailScroll(event: Event): void {
  if (railScrollIsProgrammatic) return
  const target = event.target as HTMLElement
  visibleStartPage.value = pageIndexAtScrollTop(target.scrollTop)
}

/** Upstream listbox keys: ArrowUp/Down step one page, Home/End jump. */
function onListKeydown(event: KeyboardEvent): void {
  const total = pageCount.value
  if (total < 1) return
  const current = Math.min(total, Math.max(1, props.currentPage ?? 1))
  let next: number | null = null
  if (event.key === "ArrowDown") next = Math.min(total, current + 1)
  else if (event.key === "ArrowUp") next = Math.max(1, current - 1)
  else if (event.key === "Home") next = 1
  else if (event.key === "End") next = total
  if (next === null) return
  event.preventDefault()
  if (next !== (props.currentPage ?? 1)) emit("selectPage", next - 1)
}

// ── Select page ────────────────────────────────────────────────────
function onSelectPage(pageIndex: number): void {
  emit("selectPage", pageIndex)
}

function onSidebarOpen(open: boolean): void {
  collapsed.value = !open
}
</script>

<style scoped>
/*
 * Upstream Extend geometry (docx-viewer.tsx `DocxThumbnailSidebarList`):
 * a 160px `w-40` rail, no header, 16px list padding, `12px` side inset per
 * option row, a `w-[92px]` page card, `gap-2` before a `text-xs` page label and
 * a `pb-3` row gap.
 */
.docx-thumbnail-panel {
  --office-ui-sidebar-width: 10rem;
  box-sizing: border-box;
  /* Upstream's rail inherits the toolbar foreground; the shared shell sets the
     sidebar-specific token, so restore the inherited colour here. */
  color: var(--office-foreground);
}
.docx-thumbnail-panel--dark {
  color-scheme: dark;
}
.docx-thumbnail-panel :deep(.office-document-sidebar__body) {
  overflow: hidden;
  padding: 0;
}
.docx-thumbnail-list {
  box-sizing: border-box;
  block-size: 100%;
  overflow-y: auto;
  padding: 0;
  outline: none;
}
/* Upstream offsets the virtual track by `DOCX_THUMBNAIL_LIST_PADDING` (16px)
   on both edges; the loading skeleton is not padded. */
.docx-thumbnail-track {
  box-sizing: border-box;
  padding-block: 16px;
}
.docx-thumbnail-list:focus-visible {
  outline: none;
}
.docx-thumbnail-spacer {
  inline-size: 1px;
  pointer-events: none;
}
.docx-thumbnail-item {
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 8px;
  inline-size: auto;
  margin-inline: 12px;
  margin-block-end: 12px;
  padding: 8px;
  border-radius: var(--office-radius-md);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: 16px;
  cursor: default;
  outline: none;
  user-select: none;
  transition: box-shadow 150ms ease;
}
.docx-thumbnail-item:hover {
  background: var(--office-sidebar-accent);
}
.docx-thumbnail-item--active {
  background: var(--office-sidebar-accent);
  color: var(--office-foreground);
}
.docx-thumbnail-item:focus-visible {
  box-shadow:
    0 0 0 2px var(--office-ring),
    0 0 0 3px var(--office-background);
}
.docx-thumbnail-card {
  flex: none;
  inline-size: 92px;
  border: 0;
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  box-shadow: var(--office-shadow-xs);
}
.docx-thumbnail-card--active {
  box-shadow: var(--office-shadow-sm);
}
/* `preview-class-name` lands inside the FileThumbnail root, so restyling it
   needs a deep selector: white page sheet at the card radius. */
.docx-thumbnail-panel :deep(.docx-thumbnail-preview) {
  border-radius: var(--office-radius-md);
  background: #fff;
}
/* A blank page a section's parity inserted stays an empty sheet; the dashed
   edge keeps it from reading as a page that failed to render. */
.docx-thumbnail-panel .docx-thumbnail-item--blank :deep(.docx-thumbnail-preview) {
  outline: 1px dashed var(--office-border);
  outline-offset: -1px;
}
.docx-thumbnail-canvas {
  display: block;
}
.docx-thumbnail-label {
  display: block;
  max-inline-size: 100%;
  overflow: hidden;
  color: inherit;
  font-size: var(--office-text-xs);
  line-height: 16px;
  text-align: center;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.docx-thumbnail-skeleton {
  box-sizing: border-box;
  padding: 16px;
  color: var(--office-muted-foreground);
}
.docx-thumbnail-skeleton__card {
  block-size: 112px;
  inline-size: 80px;
  margin-inline: auto;
  overflow: hidden;
  border-radius: var(--office-radius-md);
  background: var(--office-background);
  box-shadow: var(--office-shadow-xs);
}
.docx-thumbnail-skeleton__pulse {
  display: block;
  inline-size: 100%;
  block-size: 100%;
  background: var(--office-muted);
}
.docx-thumbnail-skeleton__caption {
  block-size: 12px;
  inline-size: 40px;
  margin-block-start: 12px;
  margin-inline: auto;
  border-radius: 9999px;
  background: var(--office-muted);
}

@media (prefers-reduced-motion: reduce) {
  .docx-thumbnail-item {
    transition: none;
  }
}

@media print {
  .docx-thumbnail-panel {
    display: none !important;
  }
}
</style>
