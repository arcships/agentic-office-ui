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

      <div class="docx-thumbnail-list" @scroll.passive="onThumbnailScroll">
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
          <div class="docx-thumbnail-canvas-wrapper" :style="thumbnailCanvasWrapperStyle">
            <canvas
              :ref="(el) => setCanvasRef(pageIndex, el as HTMLCanvasElement)"
              :width="thumbnailPixelWidth"
              :height="thumbnailPixelHeight"
              :style="thumbnailCanvasStyle"
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
  LayoutSnapshotPage,
  LayoutSnapshotParagraphFragment,
  LayoutSnapshotRun,
  LayoutSnapshotTableFragment,
} from "@arcships/docx-core"
import { buildLayoutSnapshot, resolveDocumentLayout } from "@arcships/docx-core"

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
const layoutSnapshot = computed(() => {
  const model = resolvedModel.value
  return model ? buildLayoutSnapshot(model, resolvedLayoutOptions.value) : undefined
})
const pageCount = computed(() =>
  layoutSnapshot.value?.pages.length ?? Math.max(0, props.totalPages ?? 0)
)

// ── Thumbnail dimensions ───────────────────────────────────────────
const thumbnailWidth = 140
const thumbnailHeight = computed(() => {
  const pageSize = layoutSnapshot.value?.resolvedDocument.layout.pageSizePx
  if (!pageSize) return 181
  const ratio = pageSize.height / pageSize.width
  return Math.round(thumbnailWidth * ratio)
})
const devicePixelRatio = typeof window !== "undefined"
  ? Math.min(3, Math.max(1, window.devicePixelRatio || 1))
  : 1
const thumbnailScale = computed(() => {
  const sourceWidth = layoutSnapshot.value?.resolvedDocument.layout.pageSizePx.width ?? 816
  return thumbnailWidth / sourceWidth
})
const thumbnailPixelWidth = computed(() =>
  Math.max(1, Math.round(thumbnailWidth * devicePixelRatio))
)
const thumbnailPixelHeight = computed(() =>
  Math.max(1, Math.round(thumbnailHeight.value * devicePixelRatio))
)

const thumbnailCanvasWrapperStyle = computed(() => ({
  width: `${thumbnailWidth}px`,
  height: `${thumbnailHeight.value}px`,
}))

const thumbnailCanvasStyle = computed(() => ({
  width: `${thumbnailWidth}px`,
  height: `${thumbnailHeight.value}px`,
}))

// ── Page range for rendering ───────────────────────────────────────
const VISIBLE_COUNT = 20
const thumbnailItemHeight = computed(() => thumbnailHeight.value + 36)
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
const topSpacerHeight = computed(() => renderWindow.value.start * thumbnailItemHeight.value)
const bottomSpacerHeight = computed(
  () => Math.max(0, pageCount.value - renderWindow.value.end) * thumbnailItemHeight.value
)

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
  [renderedPageRange, layoutSnapshot, () => props.controller?.documentLoadNonce],
  scheduleThumbnailRender,
  { immediate: true, flush: "post" }
)

type ThumbnailTextRun = Extract<LayoutSnapshotRun, { kind: "text" }>

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
  return Math.max(naturalSize, 3.5 / thumbnailScale.value)
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

function pageMargins(
  page: LayoutSnapshotPage
): { top: number; right: number; bottom: number; left: number } {
  const explicitMargin = props.layoutOptions?.margin
  if (Number.isFinite(explicitMargin)) {
    const margin = Math.max(0, Number(explicitMargin))
    return { top: margin, right: margin, bottom: margin, left: margin }
  }
  const margins = documentLayout.value?.marginsPx
  return margins
    ? { top: margins.top, right: margins.right, bottom: margins.bottom, left: margins.left }
    : page.marginsPx
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
  ctx.lineWidth = Math.max(1, 0.75 / thumbnailScale.value)
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
  const lineWidth = Math.max(1, 0.75 / thumbnailScale.value)

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

      const padding = Math.max(4, 1.5 / thumbnailScale.value)
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

  const scale = thumbnailScale.value
  const dpr = devicePixelRatio

  for (const pageIndex of renderedPageRange.value) {
    const canvas = canvasRefs.get(pageIndex)
    if (!canvas) continue

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
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      ctx.clearRect(0, 0, canvas.width, canvas.height)
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
      ctx.fillStyle = "#ffffff"
      ctx.fillRect(0, 0, thumbnailWidth, thumbnailHeight.value)
      ctx.strokeStyle = "#d4d4d4"
      ctx.lineWidth = 1
      ctx.strokeRect(0.5, 0.5, thumbnailWidth - 1, thumbnailHeight.value - 1)

      ctx.save()
      ctx.scale(scale, scale)
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
function onThumbnailScroll(event: Event): void {
  const target = event.target as HTMLElement
  const firstVisibleIndex = Math.floor(target.scrollTop / thumbnailItemHeight.value)
  visibleStartPage.value = firstVisibleIndex
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
