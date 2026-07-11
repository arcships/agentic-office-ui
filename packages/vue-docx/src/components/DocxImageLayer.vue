<template>
  <div ref="imageLayerEl" v-if="floatingImages.length > 0" class="docx-image-layer" :style="layerStyle">
    <div
      v-for="image of floatingImages"
      :key="image.key"
      class="docx-floating-image-container"
      data-testid="docx-floating-image"
      :data-image-node-index="image.nodeIndex"
      :data-image-child-index="image.childIndex"
      :data-image-layout="image.flowWrapped ? 'wrapped' : 'absolute'"
      :class="{
        'docx-floating-image--selected': image.selected,
      }"
      :style="imageContainerStyle(image)"
      @mousedown="onImageMouseDown(image, $event)"
    >
      <div
        v-if="image.selected && editable"
        class="docx-image-wrap-toolbar"
        data-testid="docx-image-wrap-toolbar"
        @mousedown.stop
      >
        <button type="button" data-testid="docx-image-wrap-inline" title="Inline with text" @click.stop="setWrapMode(image, 'inline')">Inline</button>
        <button type="button" data-testid="docx-image-wrap-square" title="Wrap text around image" @click.stop="setWrapMode(image, 'square')">Square</button>
        <button type="button" data-testid="docx-image-wrap-behind" title="Behind text" @click.stop="setWrapMode(image, 'behindText')">Behind</button>
        <button type="button" data-testid="docx-image-wrap-front" title="In front of text" @click.stop="setWrapMode(image, 'inFrontOfText')">Front</button>
      </div>

      <!-- Resize handles -->
      <template v-if="image.selected && editable">
        <div
          v-for="handle of resizeHandles"
          :key="handle.position"
          class="docx-image-resize-handle"
          data-testid="docx-image-resize-handle"
          :class="`docx-image-resize-handle--${handle.position}`"
          :style="handle.style"
          @mousedown.prevent.stop="onResizeStart(image, handle.position, $event)"
        />
      </template>

      <img
        v-if="!image.flowWrapped && image.src && !failedImageKeys.has(image.key)"
        :src="image.src"
        :alt="image.alt ?? 'Image'"
        :draggable="false"
        :style="imageStyle(image)"
        class="docx-floating-image"
        loading="lazy"
        decoding="async"
        data-image-state="loading"
        @load="onImageLoad"
        @error="onImageError(image, $event)"
      />

      <!-- Fallback for unsupported images -->
      <div
        v-else-if="!image.flowWrapped"
        class="docx-floating-image-fallback"
        :data-image-state="failedImageKeys.has(image.key) ? 'error' : 'unsupported'"
        :style="{ width: `${image.widthPx}px`, height: `${image.heightPx}px` }"
      >
        {{ image.fallbackLabel ?? "Image" }}
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, nextTick, onUnmounted } from "vue"
import type {
  DocxEditorController,
  DocxImageLocation,
  DocxImageWrapMode,
  DocumentPageNodeSegment,
  ImageRunNode,
} from "@arcships/docx-core"
import {
  resolveRenderableImageSource,
  shouldRenderWrappedFloatingImage,
} from "@arcships/docx-core"

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
  flowWrapped: boolean
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
const failedImageKeys = ref(new Set<string>())
const imageLayerEl = ref<HTMLElement | null>(null)
const layoutRevision = ref(0)

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

      const run = child as ImageRunNode
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
        flowWrapped: shouldRenderWrappedFloatingImage(run),
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
  void layoutRevision.value
  const drag = dragPreview.value?.key === image.key ? dragPreview.value : undefined
  const resize = resizePreview.value?.key === image.key ? resizePreview.value : undefined
  const flowBox = image.flowWrapped ? wrappedImageBox(image) : undefined
  const dragDeltaX = drag ? drag.xPx - image.xPx : 0
  const dragDeltaY = drag ? drag.yPx - image.yPx : 0
  const leftPx = flowBox ? flowBox.left + dragDeltaX : (drag?.xPx ?? image.xPx)
  const topPx = flowBox ? flowBox.top + dragDeltaY : (drag?.yPx ?? image.yPx)
  return {
    position: "absolute",
    left: `${leftPx}px`,
    top: `${topPx}px`,
    width: `${resize?.widthPx ?? flowBox?.width ?? image.widthPx}px`,
    height: `${resize?.heightPx ?? flowBox?.height ?? image.heightPx}px`,
    pointerEvents: "auto",
    cursor: "move",
    outline: image.selected ? "2px solid #3b82f6" : "none",
    outlineOffset: "2px",
  }
}

function wrappedImageBox(
  image: FloatingImageEntry
): { left: number; top: number; width: number; height: number } | undefined {
  const layer = imageLayerEl.value
  const surface = layer?.closest<HTMLElement>("[data-docx-page-surface='true']")
  if (!surface) return undefined
  const paragraph = surface.querySelector<HTMLElement>(
    `[data-docx-paragraph-node-index="${image.nodeIndex}"]`
  )
  const wrapped = paragraph?.querySelector<HTMLElement>(
    `[data-docx-image-child-index="${image.childIndex}"][data-docx-image-layout="wrapped"]`
  )
  if (!wrapped) return undefined
  const surfaceRect = surface.getBoundingClientRect()
  const wrappedRect = wrapped.getBoundingClientRect()
  const scale = surface.offsetWidth > 0
    ? surfaceRect.width / surface.offsetWidth
    : 1
  const safeScale = Number.isFinite(scale) && scale > 0 ? scale : 1
  return {
    left: (wrappedRect.left - surfaceRect.left) / safeScale,
    top: (wrappedRect.top - surfaceRect.top) / safeScale,
    width: wrappedRect.width / safeScale,
    height: wrappedRect.height / safeScale,
  }
}

function imageStyle(image: FloatingImageEntry): Record<string, string> {
  return {
    width: "100%",
    height: "100%",
    objectFit: "contain" as const,
  }
}

function onImageLoad(event: Event): void {
  const target = event.currentTarget as HTMLImageElement | null
  if (target) target.dataset.imageState = "ready"
}

function onImageError(image: FloatingImageEntry, event: Event): void {
  const target = event.currentTarget as HTMLImageElement | null
  if (target) target.dataset.imageState = "error"
  const next = new Set(failedImageKeys.value)
  next.add(image.key)
  failedImageKeys.value = next
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

function setWrapMode(image: FloatingImageEntry, mode: DocxImageWrapMode): void {
  if (!props.editable) return
  props.controller.setImageWrapMode(imageLocation(image), mode, {
    xPx: image.xPx,
    yPx: image.yPx,
  })
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

watch(
  () => floatingImages.value.map((image) => [
    image.key,
    image.flowWrapped,
    image.xPx,
    image.yPx,
    image.widthPx,
    image.heightPx,
  ].join(":")).join("|"),
  () => {
    void nextTick(() => {
      layoutRevision.value += 1
    })
  },
  { immediate: true }
)

onUnmounted(() => {
  finishPointerInteraction(false)
  failedImageKeys.value.clear()
})
</script>

<style scoped>
.docx-image-layer {
  overflow: visible;
}
.docx-floating-image-container {
  position: absolute;
}
.docx-image-wrap-toolbar {
  align-items: center;
  background: #ffffff;
  border: 1px solid #d4d4d8;
  border-radius: 7px;
  box-shadow: 0 5px 16px rgba(15, 23, 42, 0.16);
  display: flex;
  gap: 2px;
  left: 0;
  padding: 3px;
  position: absolute;
  top: -38px;
  white-space: nowrap;
  z-index: 22;
}
.docx-image-wrap-toolbar button {
  background: transparent;
  border: 0;
  border-radius: 4px;
  color: #3f3f46;
  cursor: pointer;
  font: inherit;
  font-size: 11px;
  padding: 4px 6px;
}
.docx-image-wrap-toolbar button:hover,
.docx-image-wrap-toolbar button:focus-visible {
  background: #f4f4f5;
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
