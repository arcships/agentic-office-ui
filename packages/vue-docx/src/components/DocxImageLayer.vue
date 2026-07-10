<template>
  <div v-if="floatingImages.length > 0" class="docx-image-layer" :style="layerStyle">
    <div
      v-for="image of floatingImages"
      :key="image.key"
      class="docx-floating-image-container"
      :class="{
        'docx-floating-image--selected': image.selected,
      }"
      :style="imageContainerStyle(image)"
      @mousedown="onImageMouseDown(image, $event)"
    >
      <!-- Resize handles -->
      <template v-if="image.selected && editable">
        <div
          v-for="handle of resizeHandles"
          :key="handle.position"
          class="docx-image-resize-handle"
          :class="`docx-image-resize-handle--${handle.position}`"
          :style="handle.style"
          @mousedown.prevent.stop="onResizeStart(image, handle.position, $event)"
        />
      </template>

      <img
        v-if="image.src"
        :src="image.src"
        :alt="image.alt ?? 'Image'"
        :draggable="false"
        :style="imageStyle(image)"
        class="docx-floating-image"
      />

      <!-- Fallback for unsupported images -->
      <div
        v-else
        class="docx-floating-image-fallback"
        :style="{ width: `${image.widthPx}px`, height: `${image.heightPx}px` }"
      >
        {{ image.fallbackLabel ?? "Image" }}
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onUnmounted } from "vue"
import type {
  DocxEditorController,
  DocxImageLocation,
  DocumentPageNodeSegment,
} from "@extend-ai/docx-core"
import { resolveRenderableImageSource } from "@extend-ai/docx-core"

// ── Types ──────────────────────────────────────────────────────────
interface FloatingImageEntry {
  key: string
  nodeIndex: number
  childIndex: number
  src?: string
  alt?: string
  widthPx: number
  heightPx: number
  xPx: number
  yPx: number
  selected: boolean
  fallbackLabel?: string
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    pageIndex: number
    pageWidthPx: number
    pageHeightPx: number
    controller: DocxEditorController
    pageNodeSegments: DocumentPageNodeSegment[]
    editable?: boolean
  }>(),
  { editable: false }
)

// ── State ──────────────────────────────────────────────────────────
const draggingImage = ref<FloatingImageEntry | null>(null)
const dragStartX = ref(0)
const dragStartY = ref(0)
const dragStartPosX = ref(0)
const dragStartPosY = ref(0)
const dragPreview = ref<{ key: string; xPx: number; yPx: number } | null>(null)

const resizingImage = ref<FloatingImageEntry | null>(null)
const resizeStartSize = ref<{ w: number; h: number }>({ w: 0, h: 0 })
const resizeStartPos = ref<{ x: number; y: number }>({ x: 0, y: 0 })
const resizePosition = ref("se")
const resizePreview = ref<{ key: string; widthPx: number; heightPx: number } | null>(null)
let pointerListenersAttached = false

// ── Computed ───────────────────────────────────────────────────────
const floatingImages = computed<FloatingImageEntry[]>(() => {
  const model = props.controller.model
  if (!model) return []

  const images: FloatingImageEntry[] = []
  const pageNodeIndexes = new Set(
    props.pageNodeSegments.map((segment) => segment.nodeIndex)
  )
  model.nodes.forEach((node, nodeIndex) => {
    if (!pageNodeIndexes.has(nodeIndex)) return
    if (node.type !== "paragraph") return
    node.children.forEach((child, childIndex) => {
      if (child.type !== "image") return
      if (!(child as any).floating) return // Only floating images

      const run = child as any
      const src = resolveRenderableImageSource(run) ?? undefined
      const selection = props.controller.selection
      images.push({
        key: `img-${nodeIndex}-${childIndex}`,
        nodeIndex,
        childIndex,
        src,
        alt: run.alt,
        widthPx: run.widthPx ?? 120,
        heightPx: run.heightPx ?? 80,
        xPx: run.floating?.xPx ?? 0,
        yPx: run.floating?.yPx ?? 0,
        selected: selection.kind === "paragraph" && selection.nodeIndex === nodeIndex,
        fallbackLabel: undefined,
      })
    })
  })

  return images
})

const layerStyle = computed(() => ({
  position: "absolute" as const,
  left: "0",
  top: "0",
  width: `${props.pageWidthPx}px`,
  height: `${props.pageHeightPx}px`,
  pointerEvents: "none" as const,
  zIndex: "20",
}))

const resizeHandles = [
  { position: "nw", style: { top: "-4px", left: "-4px", cursor: "nw-resize" } },
  { position: "ne", style: { top: "-4px", right: "-4px", cursor: "ne-resize" } },
  { position: "sw", style: { bottom: "-4px", left: "-4px", cursor: "sw-resize" } },
  { position: "se", style: { bottom: "-4px", right: "-4px", cursor: "se-resize" } },
]

// ── Style helpers ──────────────────────────────────────────────────
function imageContainerStyle(image: FloatingImageEntry): Record<string, string> {
  const drag = dragPreview.value?.key === image.key ? dragPreview.value : undefined
  const resize = resizePreview.value?.key === image.key ? resizePreview.value : undefined
  return {
    position: "absolute",
    left: `${drag?.xPx ?? image.xPx}px`,
    top: `${drag?.yPx ?? image.yPx}px`,
    width: `${resize?.widthPx ?? image.widthPx}px`,
    height: `${resize?.heightPx ?? image.heightPx}px`,
    pointerEvents: "auto",
    cursor: "move",
    outline: image.selected ? "2px solid #3b82f6" : "none",
    outlineOffset: "2px",
  }
}

function imageStyle(image: FloatingImageEntry): Record<string, string> {
  return {
    width: "100%",
    height: "100%",
    objectFit: "contain" as const,
  }
}

// ── Mouse handlers ─────────────────────────────────────────────────
function onImageMouseDown(image: FloatingImageEntry, event: MouseEvent): void {
  if (!props.editable) return
  event.preventDefault()
  props.controller.selectParagraph(image.nodeIndex)
  draggingImage.value = image
  dragStartX.value = event.clientX
  dragStartY.value = event.clientY
  dragStartPosX.value = image.xPx
  dragStartPosY.value = image.yPx
  dragPreview.value = { key: image.key, xPx: image.xPx, yPx: image.yPx }
  attachPointerListeners()
}

function onResizeStart(
  image: FloatingImageEntry,
  position: string,
  event: MouseEvent
): void {
  if (!props.editable) return
  resizingImage.value = image
  resizeStartSize.value = { w: image.widthPx, h: image.heightPx }
  resizeStartPos.value = { x: event.clientX, y: event.clientY }
  resizePosition.value = position
  resizePreview.value = {
    key: image.key,
    widthPx: image.widthPx,
    heightPx: image.heightPx,
  }
  attachPointerListeners()
}

function imageLocation(image: FloatingImageEntry): DocxImageLocation {
  return {
    kind: "paragraph",
    nodeIndex: image.nodeIndex,
    childIndex: image.childIndex,
  }
}

function updatePointerPreview(event: MouseEvent): void {
  if (draggingImage.value) {
    const image = draggingImage.value
    const xPx = Math.round(dragStartPosX.value + event.clientX - dragStartX.value)
    const yPx = Math.round(dragStartPosY.value + event.clientY - dragStartY.value)
    dragPreview.value = {
      key: image.key,
      xPx: Math.max(-image.widthPx + 16, Math.min(props.pageWidthPx - 16, xPx)),
      yPx: Math.max(-image.heightPx + 16, Math.min(props.pageHeightPx - 16, yPx)),
    }
  }

  if (resizingImage.value) {
    const image = resizingImage.value
    const dx = event.clientX - resizeStartPos.value.x
    const dy = event.clientY - resizeStartPos.value.y
    const widthDelta = resizePosition.value.includes("w") ? -dx : dx
    const heightDelta = resizePosition.value.includes("n") ? -dy : dy
    resizePreview.value = {
      key: image.key,
      widthPx: Math.max(16, Math.round(resizeStartSize.value.w + widthDelta)),
      heightPx: Math.max(16, Math.round(resizeStartSize.value.h + heightDelta)),
    }
  }
}

function onDocumentMouseMove(event: MouseEvent): void {
  updatePointerPreview(event)
}

function onDocumentMouseUp(event: MouseEvent): void {
  updatePointerPreview(event)
  finishPointerInteraction(true)
}

function attachPointerListeners(): void {
  if (pointerListenersAttached) return
  document.addEventListener("mousemove", onDocumentMouseMove)
  document.addEventListener("mouseup", onDocumentMouseUp)
  pointerListenersAttached = true
}

function detachPointerListeners(): void {
  if (!pointerListenersAttached) return
  document.removeEventListener("mousemove", onDocumentMouseMove)
  document.removeEventListener("mouseup", onDocumentMouseUp)
  pointerListenersAttached = false
}

function finishPointerInteraction(commit: boolean): void {
  const dragged = draggingImage.value
  const moved = dragPreview.value
  const resized = resizingImage.value
  const size = resizePreview.value

  detachPointerListeners()
  draggingImage.value = null
  resizingImage.value = null
  dragPreview.value = null
  resizePreview.value = null

  if (!commit || !props.editable) return
  if (dragged && moved?.key === dragged.key) {
    props.controller.moveFloatingImage(imageLocation(dragged), {
      xPx: moved.xPx,
      yPx: moved.yPx,
    })
  } else if (resized && size?.key === resized.key) {
    props.controller.resizeImage(
      imageLocation(resized),
      size.widthPx,
      size.heightPx,
    )
  }
}

watch(() => props.editable, (editable) => {
  if (!editable) finishPointerInteraction(false)
})

onUnmounted(() => finishPointerInteraction(false))
</script>

<style scoped>
.docx-image-layer {
  overflow: visible;
}
.docx-floating-image-container {
  position: absolute;
}
.docx-floating-image {
  display: block;
}
.docx-floating-image-fallback {
  display: flex;
  align-items: center;
  justify-content: center;
  border: 1px dashed #d1d5db;
  background: #f9fafb;
  color: #6b7280;
  font-size: 10px;
  font-weight: 700;
  text-transform: lowercase;
  border-radius: 3px;
}
.docx-image-resize-handle {
  position: absolute;
  width: 8px;
  height: 8px;
  background: #3b82f6;
  border: 1px solid #fff;
  border-radius: 1px;
  z-index: 21;
}
</style>
