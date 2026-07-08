<template>
  <div class="docx-thumbnail-panel" :class="{ 'docx-thumbnail-panel--collapsed': collapsed }">
    <!-- Toggle button -->
    <button class="docx-thumbnail-toggle" @click="collapsed = !collapsed" :title="collapsed ? 'Show thumbnails' : 'Hide thumbnails'">
      {{ collapsed ? '▶' : '◀' }}
    </button>

    <!-- Thumbnails container -->
    <div v-if="!collapsed" class="docx-thumbnail-container">
      <div class="docx-thumbnail-header">
        <span class="docx-thumbnail-title">Pages</span>
        <span class="docx-thumbnail-count">{{ totalPages }}</span>
      </div>

      <div class="docx-thumbnail-list" @scroll.passive="onThumbnailScroll">
        <div
          v-for="pageIndex of renderedPageRange"
          :key="pageIndex"
          class="docx-thumbnail-item"
          :class="{ 'docx-thumbnail-item--active': pageIndex + 1 === currentPage }"
          :data-thumbnail-page-index="pageIndex"
          @click="onSelectPage(pageIndex)"
        >
          <!-- Thumbnail canvas -->
          <div class="docx-thumbnail-canvas-wrapper" :style="thumbnailCanvasWrapperStyle">
            <canvas
              :ref="(el) => setCanvasRef(pageIndex, el as HTMLCanvasElement)"
              :width="thumbnailWidth * devicePixelRatio"
              :height="thumbnailHeight * devicePixelRatio"
              :style="thumbnailCanvasStyle"
              class="docx-thumbnail-canvas"
            />
          </div>

          <!-- Page number label -->
          <div class="docx-thumbnail-label">
            {{ pageIndex + 1 }}
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onMounted, onUnmounted } from "vue"
import type { DocxEditorController } from "@extend-ai/docx-core"
import { layoutDocument, resolveDocumentLayout } from "@extend-ai/docx-core"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    controller: DocxEditorController
    currentPage?: number
    totalPages?: number
  }>(),
  { currentPage: 1, totalPages: 1 }
)

const emit = defineEmits<{
  selectPage: [pageIndex: number]
}>()

// ── State ──────────────────────────────────────────────────────────
const collapsed = ref(true)
const visibleStartPage = ref(0)

// ── Thumbnail dimensions ───────────────────────────────────────────
const thumbnailWidth = 140
const thumbnailHeight = computed(() => {
  const docLayout = resolveDocumentLayout(props.controller.model)
  const ratio = docLayout.pageHeightPx / docLayout.pageWidthPx
  return Math.round(thumbnailWidth * ratio)
})
const devicePixelRatio = ref(typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1)
const thumbnailScale = computed(
  () => thumbnailWidth / resolveDocumentLayout(props.controller.model).pageWidthPx
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
const renderedPageRange = computed(() => {
  const start = Math.max(0, visibleStartPage.value - 5)
  const end = Math.min((props.totalPages ?? 1), start + VISIBLE_COUNT)
  const range: number[] = []
  for (let i = start; i < end; i++) {
    range.push(i)
  }
  return range
})

// ── Canvas refs ────────────────────────────────────────────────────
const canvasRefs = ref<Record<number, HTMLCanvasElement | null>>({})

function setCanvasRef(pageIndex: number, el: HTMLCanvasElement | null): void {
  canvasRefs.value = { ...canvasRefs.value, [pageIndex]: el }
}

// ── Thumbnail rendering ────────────────────────────────────────────
watch(
  [renderedPageRange, () => props.controller.model, () => props.controller.documentLoadNonce],
  () => {
    renderVisibleThumbnails()
  },
  { immediate: false }
)

async function renderVisibleThumbnails(): Promise<void> {
  const model = props.controller.model
  if (!model) return

  const pages = layoutDocument(model, {})
  const docLayout = resolveDocumentLayout(model)
  const scale = thumbnailScale.value
  const dpr = devicePixelRatio.value

  for (const pageIndex of renderedPageRange.value) {
    const canvas = canvasRefs.value[pageIndex]
    if (!canvas) continue

    const page = pages[pageIndex]
    if (!page) continue

    const ctx = canvas.getContext("2d")
    if (!ctx) continue

    // Clear canvas
    ctx.clearRect(0, 0, canvas.width, canvas.height)

    // Draw white page background
    ctx.fillStyle = "#ffffff"
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    ctx.strokeStyle = "#d4d4d4"
    ctx.lineWidth = 1 * dpr
    ctx.strokeRect(0, 0, canvas.width, canvas.height)

    // Draw content area
    const margin = (docLayout.marginsPx?.top ?? 72) * scale
    ctx.save()
    ctx.translate(margin, margin)

    // Draw each block at thumbnail scale
    let yOffset = 0
    for (const block of page.blocks) {
      if (block.kind === "paragraph") {
        // Draw text as lines
        ctx.font = `${10 * scale * dpr}px Calibri, sans-serif`
        ctx.fillStyle = "#374151"

        const text = block.runs
          .filter((r) => r.kind === "text")
          .map((r) => r.text)
          .join("")

        // Word wrap
        const maxWidth = (docLayout.pageWidthPx - (docLayout.marginsPx?.left ?? 72) * 2) * scale * dpr
        const words = text.split(" ")
        let line = ""
        const lineHeight = 12 * scale * dpr
        const lines: string[] = []

        for (const word of words) {
          const testLine = line ? `${line} ${word}` : word
          const metrics = ctx.measureText(testLine)
          if (metrics.width > maxWidth && line) {
            lines.push(line)
            line = word
          } else {
            line = testLine
          }
        }
        if (line) lines.push(line)

        // Draw lines
        for (let i = 0; i < Math.min(lines.length, 10); i++) {
          ctx.fillText(lines[i], 0, yOffset + lineHeight * (i + 0.8))
        }

        yOffset += Math.max(1, lines.length) * lineHeight + 4 * scale * dpr
      } else if (block.kind === "table") {
        // Draw simplified table representation
        const rowCount = Math.min(block.rows.length, 10)
        const cellCount = block.rows[0]?.cells.length ?? 2
        const cellW = (docLayout.pageWidthPx - (docLayout.marginsPx?.left ?? 72) * 2) * scale * dpr / cellCount
        const cellH = 8 * scale * dpr

        for (let ri = 0; ri < rowCount; ri++) {
          for (let ci = 0; ci < cellCount; ci++) {
            ctx.strokeStyle = "#e5e7eb"
            ctx.lineWidth = 0.5 * dpr
            ctx.strokeRect(ci * cellW, yOffset + ri * cellH, cellW, cellH)
            ctx.fillStyle = "#f9fafb"
            ctx.fillRect(ci * cellW, yOffset + ri * cellH, cellW, cellH)
          }
        }

        yOffset += rowCount * cellH + 4 * scale * dpr
      }
    }

    ctx.restore()
  }
}

// ── Scroll handling ────────────────────────────────────────────────
function onThumbnailScroll(event: Event): void {
  const target = event.target as HTMLElement
  const itemHeight = thumbnailHeight.value + 36 // card height + label + margins
  const firstVisibleIndex = Math.floor(target.scrollTop / itemHeight)
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
  color: #6b7280;
  margin-top: 2px;
}
</style>
