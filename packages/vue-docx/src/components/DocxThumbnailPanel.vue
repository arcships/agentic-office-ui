<template>
  <div class="docx-thumbnail-panel" data-testid="docx-thumbnail-panel" :class="{ 'docx-thumbnail-panel--collapsed': collapsed }">
    <!-- Toggle button -->
    <button v-if="showToggle" class="docx-thumbnail-toggle" @click="collapsed = !collapsed" :title="collapsed ? 'Show thumbnails' : 'Hide thumbnails'">
      {{ collapsed ? '▶' : '◀' }}
    </button>

    <!-- Thumbnails container -->
    <div v-if="!collapsed" class="docx-thumbnail-container">
      <div class="docx-thumbnail-header">
        <span class="docx-thumbnail-title">Pages</span>
        <span class="docx-thumbnail-count">{{ pageCount }}</span>
      </div>

      <div ref="listRef" class="docx-thumbnail-list" @scroll.passive="onThumbnailScroll">
        <div
          v-if="topSpacerHeight > 0"
          class="docx-thumbnail-spacer"
          :style="{ height: `${topSpacerHeight}px` }"
          aria-hidden="true"
        />
        <div
          v-for="pageIndex of renderedPageRange"
          :key="pageIndex"
          class="docx-thumbnail-item"
          :class="{ 'docx-thumbnail-item--active': pageIndex + 1 === currentPage }"
          data-testid="docx-thumbnail"
          :aria-current="pageIndex + 1 === currentPage ? 'page' : undefined"
          :data-thumbnail-page-index="pageIndex"
          @click="onSelectPage(pageIndex)"
        >
          <!-- Thumbnail canvas -->
          <div class="docx-thumbnail-canvas-wrapper" :style="thumbnailCanvasWrapperStyle(pageIndex)">
            <canvas
              :ref="(el) => setCanvasRef(pageIndex, el as HTMLCanvasElement)"
              :width="thumbnailCanvasPixelWidth()"
              :height="thumbnailCanvasPixelHeight(pageIndex)"
              :style="thumbnailCanvasStyle(pageIndex)"
              class="docx-thumbnail-canvas"
            />
          </div>

          <!-- Page number label -->
          <div class="docx-thumbnail-label">
            {{ pageIndex + 1 }}
          </div>
        </div>
        <div
          v-if="bottomSpacerHeight > 0"
          class="docx-thumbnail-spacer"
          :style="{ height: `${bottomSpacerHeight}px` }"
          aria-hidden="true"
        />
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, watch } from "vue"
import type {
  DocModel,
  DocxEditorController,
  LayoutOptions,
  LayoutSnapshot,
  LayoutSnapshotPage,
  LayoutSnapshotParagraphFragment,
  LayoutSnapshotRun,
  LayoutSnapshotTableFragment,
} from "@arcships/docx-core"
import { buildLayoutSnapshot, parseSectionLayout, resolveDocumentLayout, resolveDocumentSectionsFromMetadata, resolveSectionPropertiesXmlForNodeIndex } from "@arcships/docx-core"
import {
  ensureDocxViewerPageSurfaceRegistry,
  subscribeDocxViewerPageSurfaces,
} from "../composables/page-surface-registry"

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
  }>(),
  { currentPage: 1, totalPages: 1, defaultCollapsed: false, showToggle: true }
)

const emit = defineEmits<{
  selectPage: [pageIndex: number]
}>()

// ── State ──────────────────────────────────────────────────────────
const collapsed = ref(props.defaultCollapsed)
const visibleStartPage = ref(0)
const listRef = ref<HTMLElement | undefined>()
const registryRevision = ref(0)

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

// ── Thumbnail dimensions ───────────────────────────────────────────
const THUMBNAIL_WIDTH = 140
const THUMBNAIL_ITEM_EXTRA_HEIGHT = 36
const DEFAULT_PAGE_SIZE = { width: 816, height: 1056 }

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

const devicePixelRatio = typeof window !== "undefined"
  ? Math.min(3, Math.max(1, window.devicePixelRatio || 1))
  : 1

function pageMetric(pageIndex: number): ThumbnailPageMetric {
  return pageMetrics.value[pageIndex] ?? {
    width: THUMBNAIL_WIDTH,
    height: 181,
    scale: THUMBNAIL_WIDTH / DEFAULT_PAGE_SIZE.width,
    top: pageIndex * (181 + THUMBNAIL_ITEM_EXTRA_HEIGHT),
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
  return Math.max(1, Math.round(THUMBNAIL_WIDTH * devicePixelRatio))
}

function thumbnailCanvasPixelHeight(pageIndex: number): number {
  return Math.max(1, Math.round(pageMetric(pageIndex).height * devicePixelRatio))
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

function setCanvasRef(pageIndex: number, el: HTMLCanvasElement | null): void {
  if (!el) {
    canvasRefs.delete(pageIndex)
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
  [renderedPageRange, layoutSnapshot, pageMetrics, () => props.controller?.documentLoadNonce],
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
  maxLines = 1
): void {
  if (runs.length === 0 || frame.width <= 0 || frame.height <= 0) return
  const largestFont = runs.reduce(
    (largest, run) => Math.max(largest, runFontSize(run, headingLevel)),
    12
  )
  const lineHeight = Math.max(16, largestFont * 1.2)
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
        if (lineIndex >= maxLines) break outer
        continue
      }
      const token = rawToken === "\t" ? "    " : rawToken
      const fontSize = applyRunStyle(ctx, run, headingLevel)
      const tokenWidth = ctx.measureText(token).width
      if (cursorX > frame.x && cursorX + tokenWidth > right && token.trim()) {
        lineIndex++
        cursorX = frame.x
        if (lineIndex >= maxLines) break outer
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
  block: LayoutSnapshotParagraphFragment,
  page: LayoutSnapshotPage
): void {
  const image = block.runs.find((run) => run.kind === "image")
  if (!image) return
  const frame = contentFrame(page, block.framePx)
  const width = Math.min(frame.width, Math.max(24, image.widthPx ?? 96))
  const height = Math.max(
    18,
    Math.min(image.heightPx ?? block.framePx.height, block.framePx.height || 96)
  )
  ctx.fillStyle = "#e5e7eb"
  ctx.fillRect(frame.x, frame.y, width, height)
  ctx.strokeStyle = "#9ca3af"
  ctx.lineWidth = Math.max(1, 0.75 / renderScale)
  ctx.strokeRect(frame.x, frame.y, width, height)
  ctx.beginPath()
  ctx.moveTo(frame.x, frame.y)
  ctx.lineTo(frame.x + width, frame.y + height)
  ctx.moveTo(frame.x + width, frame.y)
  ctx.lineTo(frame.x, frame.y + height)
  ctx.stroke()
}

function drawParagraph(
  ctx: CanvasRenderingContext2D,
  block: LayoutSnapshotParagraphFragment,
  page: LayoutSnapshotPage
): void {
  const frame = contentFrame(page, block.framePx)
  const textRuns = block.runs.filter((run): run is ThumbnailTextRun => run.kind === "text")
  drawTextRuns(ctx, textRuns, frame, block.headingLevel, 1)
  drawImagePlaceholder(ctx, block, page)
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
      drawTextRuns(
        ctx,
        textRuns,
        {
          x: cellX + padding,
          y: cellY + padding,
          width: Math.max(1, cellWidth - padding * 2),
          height: Math.max(1, rowHeight - padding * 2),
        },
        undefined,
        1
      )
      columnCursor += colSpan
    })
  })
}

function renderVisibleThumbnails(): void {
  const snapshot = layoutSnapshot.value
  if (!snapshot) return

  const dpr = devicePixelRatio

  for (const pageIndex of renderedPageRange.value) {
    const canvas = canvasRefs.get(pageIndex)
    if (!canvas) continue

    const metric = pageMetrics.value[pageIndex]
    if (!metric) continue
    const page = snapshot.pages[pageIndex]
    if (!page) {
      canvas.dataset.thumbnailState = "unavailable"
      canvas.dataset.thumbnailContentBlocks = "0"
      continue
    }

    const ctx = canvas.getContext("2d")
    if (!ctx) {
      canvas.dataset.thumbnailState = "error"
      continue
    }
    canvas.dataset.thumbnailState = "rendering"

    try {
      renderScale = metric.scale
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      ctx.clearRect(0, 0, canvas.width, canvas.height)
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
      ctx.fillStyle = "#ffffff"
      ctx.fillRect(0, 0, metric.width, metric.height)
      ctx.strokeStyle = "#d4d4d4"
      ctx.lineWidth = 1
      ctx.strokeRect(0.5, 0.5, metric.width - 1, metric.height - 1)

      ctx.save()
      ctx.scale(metric.scale, metric.scale)
      for (const block of page.blocks) {
        if (block.kind === "paragraph") drawParagraph(ctx, block, page)
        else drawTable(ctx, block, page)
      }
      ctx.restore()

      canvas.dataset.thumbnailState = "ready"
      canvas.dataset.thumbnailContentBlocks = String(page.blocks.length)
    } catch {
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      canvas.dataset.thumbnailState = "error"
      canvas.dataset.thumbnailContentBlocks = "0"
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

// ── Select page ────────────────────────────────────────────────────
function onSelectPage(pageIndex: number): void {
  emit("selectPage", pageIndex)
}
</script>

<style scoped>
.docx-thumbnail-panel {
  display: flex;
  flex-shrink: 0;
  background: #f9fafb;
  border-right: 1px solid #e5e7eb;
  position: relative;
  transition: width 150ms ease;
}
.docx-thumbnail-panel--collapsed {
  width: 28px;
}
.docx-thumbnail-toggle {
  position: absolute;
  right: 4px;
  top: 8px;
  width: 20px;
  height: 20px;
  border: none;
  background: #e5e7eb;
  border-radius: 3px;
  cursor: pointer;
  font-size: 10px;
  color: #6b7280;
  z-index: 1;
  display: flex;
  align-items: center;
  justify-content: center;
}
.docx-thumbnail-toggle:hover {
  background: #d1d5db;
}
.docx-thumbnail-container {
  width: 170px;
  display: flex;
  flex-direction: column;
}
.docx-thumbnail-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 8px 12px;
  border-bottom: 1px solid #e5e7eb;
}
.docx-thumbnail-title {
  font-size: 12px;
  font-weight: 600;
  color: #374151;
}
.docx-thumbnail-count {
  font-size: 11px;
  color: #6b7280;
  background: #e5e7eb;
  border-radius: 10px;
  padding: 1px 8px;
}
.docx-thumbnail-list {
  flex: 1;
  overflow-y: auto;
  padding: 8px;
}
.docx-thumbnail-spacer {
  width: 1px;
  pointer-events: none;
}
.docx-thumbnail-item {
  margin-bottom: 12px;
  cursor: pointer;
  border-radius: 4px;
  padding: 4px;
  transition: background 100ms ease;
}
.docx-thumbnail-item:hover {
  background: #f3f4f6;
}
.docx-thumbnail-item--active {
  background: #eff6ff;
  outline: 2px solid #3b82f6;
  outline-offset: -2px;
}
.docx-thumbnail-canvas-wrapper {
  overflow: hidden;
  border-radius: 2px;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
  background: #fff;
}
.docx-thumbnail-canvas {
  display: block;
}
.docx-thumbnail-label {
  text-align: center;
  font-size: 11px;
  line-height: 14px;
  color: #6b7280;
  margin-top: 2px;
}
</style>
