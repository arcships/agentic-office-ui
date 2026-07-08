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
import { ref, computed } from "vue"
import type {
  DocxEditorController,
  DocxImageLocation,
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

const resizingImage = ref<FloatingImageEntry | null>(null)
const resizeStartSize = ref<{ w: number; h: number }>({ w: 0, h: 0 })
const resizeStartPos = ref<{ x: number; y: number }>({ x: 0, y: 0 })

// ── Computed ───────────────────────────────────────────────────────
const floatingImages = computed<FloatingImageEntry[]>(() => {
  const model = props.controller.model
  if (!model) return []

  const images: FloatingImageEntry[] = []
  model.nodes.forEach((node, nodeIndex) => {
    if (node.type !== "paragraph") return
    node.children.forEach((child, childIndex) => {
      if (child.type !== "image") return
      if (!(child as any).floating) return // Only floating images

      const run = child as any
      const src = resolveRenderableImageSource(run) ?? undefined
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
        selected: false,
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
  return {
    position: "absolute",
    left: `${image.xPx}px`,
    top: `${image.yPx}px`,
    width: `${image.widthPx}px`,
    height: `${image.heightPx}px`,
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
  draggingImage.value = image
  dragStartX.value = event.clientX
  dragStartY.value = event.clientY
  dragStartPosX.value = image.xPx
  dragStartPosY.value = image.yPx

  const onMouseMove = (e: MouseEvent) => {
    if (!draggingImage.value) return
    const dx = e.clientX - dragStartX.value
    const dy = e.clientY - dragStartY.value
    // Update image position in-place (simplified)
  }

  const onMouseUp = () => {
    draggingImage.value = null
    document.removeEventListener("mousemove", onMouseMove)
    document.removeEventListener("mouseup", onMouseUp)
  }

  document.addEventListener("mousemove", onMouseMove)
  document.addEventListener("mouseup", onMouseUp)
}

function onResizeStart(
  image: FloatingImageEntry,
  _position: string,
  event: MouseEvent
): void {
  resizingImage.value = image
  resizeStartSize.value = { w: image.widthPx, h: image.heightPx }
  resizeStartPos.value = { x: event.clientX, y: event.clientY }

  const onMouseMove = (e: MouseEvent) => {
    if (!resizingImage.value) return
    const dx = e.clientX - resizeStartPos.value.x
    const dy = e.clientY - resizeStartPos.value.y
    // Update image size in-place (simplified)
  }

  const onMouseUp = () => {
    if (resizingImage.value) {
      props.controller.resizeImage(
        { kind: "paragraph", nodeIndex: resizingImage.value.nodeIndex, childIndex: resizingImage.value.childIndex },
        resizeStartSize.value.w,
        resizeStartSize.value.h
      )
    }
    resizingImage.value = null
    document.removeEventListener("mousemove", onMouseMove)
    document.removeEventListener("mouseup", onMouseUp)
  }

  document.addEventListener("mousemove", onMouseMove)
  document.addEventListener("mouseup", onMouseUp)
}
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
