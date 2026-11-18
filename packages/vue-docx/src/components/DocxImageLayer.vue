<template>
  <div ref="imageLayerEl" v-if="floatingImages.length > 0" class="docx-image-layer-root">
    <div
      v-for="layer of imageLayers"
      :key="layer.key"
      class="docx-image-layer"
      :class="`docx-image-layer--${layer.key}`"
      :data-docx-image-layer="layer.key"
      :style="layerStyle(layer.key)"
    >
      <div
        v-for="image of layer.images"
        :key="image.key"
        class="docx-floating-image-container"
        data-testid="docx-floating-image"
        :data-image-node-index="image.nodeIndex"
        :data-image-child-index="image.childIndex"
        :data-image-layout="image.flowWrapped ? 'wrapped' : 'absolute'"
        :data-image-behind="image.behindDocument ? 'true' : undefined"
        :class="{
          'docx-floating-image--selected': image.selected,
        }"
        :style="imageContainerStyle(image)"
        @mousedown="onImageMouseDown(image, $event)"
      >



        <div
          v-if="image.selected && editable && !image.behindDocument && !image.lockedAnchor"
          class="docx-image-wrap-toolbar"
          data-testid="docx-image-wrap-toolbar"
          data-docx-thumbnail-exclude="true"
          @mousedown.stop
        >
          <button type="button" data-testid="docx-image-wrap-inline" :title="m.docx.image.inlineTitle" @click.stop="setWrapMode(image, 'inline')">{{ m.docx.image.inline }}</button>
          <button type="button" data-testid="docx-image-wrap-square" :title="m.docx.image.squareTitle" @click.stop="setWrapMode(image, 'square')">{{ m.docx.image.square }}</button>
          <button type="button" data-testid="docx-image-wrap-behind" :title="m.docx.image.behindTitle" @click.stop="setWrapMode(image, 'behindText')">{{ m.docx.image.behind }}</button>
          <button type="button" data-testid="docx-image-wrap-front" :title="m.docx.image.frontTitle" @click.stop="setWrapMode(image, 'inFrontOfText')">{{ m.docx.image.front }}</button>
        </div>

        <!-- Resize handles -->
        <template v-if="image.selected && editable && !image.behindDocument && !image.lockedAnchor">
          <div
            v-for="handle of resizeHandles"
            :key="handle.position"
            class="docx-image-resize-handle"
            data-testid="docx-image-resize-handle"
            data-image-resize-handle="true"
            :class="`docx-image-resize-handle--${handle.position}`"
            :style="handle.style"
            @mousedown.prevent.stop="onResizeStart(image, handle.position, $event)"
          />
        </template>

        <img
          v-if="!image.flowWrapped && image.src && !failedImageSourceKeys.has(image.failureKey)"
          :src="image.src"
          :alt="image.alt ?? m.docx.image.alt"
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
          :data-image-state="failedImageSourceKeys.has(image.failureKey) ? 'error' : 'unsupported'"
          :style="{ width: `${image.widthPx}px`, height: `${image.heightPx}px` }"
        >
          {{ image.fallbackLabel ?? m.docx.image.alt }}
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, nextTick, onMounted, onUnmounted, onUpdated } from "vue"
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
  unsupportedImageFallbackLabel,
} from "@arcships/docx-core"
import { useDocxMessages } from "../i18n"

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
  behindDocument: boolean
  selected: boolean
  lockedAnchor: boolean
  failureKey: string
  fallbackLabel?: string
}

interface WrappedImageBox {
  left: number
  top: number
  width: number
  height: number
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
    zoomFactor?: number
    /** Image keys (`nodeIndex:childIndex`) rendered by the page cover layer. */
    skipImageKeys?: readonly string[]
  }>(),
  { editable: false, zoomFactor: 1 }
)

const m = useDocxMessages()

// ── State ──────────────────────────────────────────────────────────
const draggingImage = ref<FloatingImageEntry | null>(null)
const dragStartX = ref(0)
const dragStartY = ref(0)
const dragStartPosX = ref(0)
const dragStartPosY = ref(0)
const dragPreview = ref<{ key: string; xPx: number; yPx: number } | null>(null)
const failedImageSourceKeys = ref(new Set<string>())
const imageLayerEl = ref<HTMLElement | null>(null)
const wrappedBoxes = ref(new Map<string, WrappedImageBox>())

const resizingImage = ref<FloatingImageEntry | null>(null)
const resizeStartSize = ref<{ w: number; h: number }>({ w: 0, h: 0 })
const resizeStartPos = ref<{ x: number; y: number }>({ x: 0, y: 0 })
const resizePosition = ref("se")
const resizePreview = ref<{ key: string; widthPx: number; heightPx: number } | null>(null)
let pointerListenersAttached = false
let resizeObserver: ResizeObserver | null = null
let wrappedBoxRefreshScheduled = false

function imageSourceFailureKey(src: string | undefined): string {
  if (!src) return "missing"
  let hash = 2166136261
  for (let index = 0; index < src.length; index += 1) {
    hash ^= src.charCodeAt(index)
    hash = Math.imul(hash, 16777619)
  }
  return `${src.length}:${(hash >>> 0).toString(36)}`
}

// ── Computed ───────────────────────────────────────────────────────
const floatingImages = computed<FloatingImageEntry[]>(() => {
  const model = props.controller.model
  if (!model) return []

  const images: FloatingImageEntry[] = []
  const skippedImageKeys = new Set(props.skipImageKeys ?? [])
  const pageNodeIndexes = new Set(
    props.pageNodeSegments.map((segment) => segment.nodeIndex)
  )
  model.nodes.forEach((node, nodeIndex) => {
    if (!pageNodeIndexes.has(nodeIndex)) return
    if (node.type !== "paragraph") return
    node.children.forEach((child, childIndex) => {
      if (child.type !== "image") return
      if (!(child as any).floating) return // Only floating images
      if (skippedImageKeys.has(`${nodeIndex}:${childIndex}`)) return

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
        behindDocument: run.floating?.behindDocument === true,
        selected: selection.kind === "paragraph" && selection.nodeIndex === nodeIndex,
        lockedAnchor: run.floating?.lockedAnchor === true,
        failureKey: imageSourceFailureKey(src),
        // A picture the pipeline cannot paint gets the same label as the
        // in-flow placeholder; a node without any source keeps its alt text.
        fallbackLabel:
          !src && run.src
            ? unsupportedImageFallbackLabel(run, run.widthPx, run.heightPx)
            : undefined,
      })
    })
  })

  return images
})

const imageLayers = computed(() => {
  const behind = floatingImages.value.filter((image) => image.behindDocument)
  const front = floatingImages.value.filter((image) => !image.behindDocument)
  return [
    { key: "behind" as const, images: behind },
    { key: "front" as const, images: front },
  ].filter((layer) => layer.images.length > 0)
})

function layerStyle(kind: "behind" | "front") {
  return {
    position: "absolute" as const,
    left: "0",
    top: "0",
    width: `${props.pageWidthPx}px`,
    height: `${props.pageHeightPx}px`,
    pointerEvents: "none" as const,
    zIndex: kind === "behind" ? "-1" : "20",
  }
}

const resizeHandles = [
  { position: "nw", style: { top: "-4px", left: "-4px", cursor: "nw-resize" } },
  { position: "ne", style: { top: "-4px", right: "-4px", cursor: "ne-resize" } },
  { position: "sw", style: { bottom: "-4px", left: "-4px", cursor: "sw-resize" } },
  { position: "se", style: { bottom: "-4px", right: "-4px", cursor: "se-resize" } },
]

// Pointer deltas are screen pixels; the page content is CSS-scaled by the
// viewer. Convert them back to page coordinates before mutating the model.
const pointerScale = computed(() => {
  const factor = props.zoomFactor
  return Number.isFinite(factor) && factor > 0 ? factor : 1
})

// ── Style helpers ──────────────────────────────────────────────────
function imageContainerStyle(image: FloatingImageEntry): Record<string, string> {
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
    pointerEvents: image.behindDocument ? "none" : "auto",
    cursor: image.lockedAnchor ? "default" : image.behindDocument ? "default" : "move",
    outline: image.selected && !image.behindDocument ? "2px solid #3b82f6" : "none",
    outlineOffset: "2px",
  }
}

/**
 * Box of the in-flow (wrapped) image inside its anchor paragraph, in page
 * coordinates. Measured from the DOM and cached in `wrappedBoxes`: a wrapped
 * image moves whenever text above it reflows (edits, font loading, resize,
 * zoom), so the transparent drag target is re-derived from layout instead of
 * being frozen at the position it had on the render that created it.
 */
function wrappedImageBox(image: FloatingImageEntry): WrappedImageBox | undefined {
  return wrappedBoxes.value.get(image.key)
}

function pageSurfaceElement(): HTMLElement | undefined {
  const layer = imageLayerEl.value
  if (!layer || typeof layer.closest !== "function") return undefined
  return layer.closest<HTMLElement>("[data-docx-page-surface='true']") ?? undefined
}

function measureScale(surface: HTMLElement, surfaceRect: DOMRect): number {
  const scale =
    surface.offsetWidth > 0 ? surfaceRect.width / surface.offsetWidth : 1
  return Number.isFinite(scale) && scale > 0 ? scale : 1
}

function measureWrappedBoxes(): void {
  const surface = pageSurfaceElement()
  const next = new Map<string, WrappedImageBox>()
  if (
    surface &&
    typeof surface.querySelector === "function" &&
    typeof surface.getBoundingClientRect === "function"
  ) {
    const surfaceRect = surface.getBoundingClientRect()
    const scale = measureScale(surface, surfaceRect)
    for (const image of floatingImages.value) {
      if (!image.flowWrapped) continue
      const paragraph = surface.querySelector<HTMLElement>(
        `[data-docx-paragraph-node-index="${image.nodeIndex}"]`
      )
      const wrapped = paragraph?.querySelector<HTMLElement>(
        `[data-docx-image-child-index="${image.childIndex}"][data-docx-image-layout="wrapped"]`
      )
      if (!wrapped || typeof wrapped.getBoundingClientRect !== "function") continue
      const wrappedRect = wrapped.getBoundingClientRect()
      next.set(image.key, {
        left: (wrappedRect.left - surfaceRect.left) / scale,
        top: (wrappedRect.top - surfaceRect.top) / scale,
        width: wrappedRect.width / scale,
        height: wrappedRect.height / scale,
      })
    }
  }

  if (sameWrappedBoxes(next, wrappedBoxes.value)) return
  wrappedBoxes.value = next
}

function sameWrappedBoxes(
  next: Map<string, WrappedImageBox>,
  current: Map<string, WrappedImageBox>
): boolean {
  if (next.size !== current.size) return false
  for (const [key, box] of next) {
    const previous = current.get(key)
    if (!previous) return false
    if (
      Math.abs(previous.left - box.left) > 0.5 ||
      Math.abs(previous.top - box.top) > 0.5 ||
      Math.abs(previous.width - box.width) > 0.5 ||
      Math.abs(previous.height - box.height) > 0.5
    ) {
      return false
    }
  }
  return true
}

function scheduleWrappedBoxRefresh(): void {
  if (wrappedBoxRefreshScheduled) return
  wrappedBoxRefreshScheduled = true
  void nextTick(() => {
    wrappedBoxRefreshScheduled = false
    measureWrappedBoxes()
  })
}

function observeWrappedLayout(): void {
  const ResizeObserverConstructor = globalThis.ResizeObserver
  if (typeof ResizeObserverConstructor !== "function") return
  const surface = pageSurfaceElement()
  if (!surface) return
  resizeObserver ??= new ResizeObserverConstructor(() => {
    scheduleWrappedBoxRefresh()
  })
  resizeObserver.observe(surface)
  for (const image of floatingImages.value) {
    if (!image.flowWrapped) continue
    const paragraph = surface.querySelector<HTMLElement>(
      `[data-docx-paragraph-node-index="${image.nodeIndex}"]`
    )
    if (paragraph) resizeObserver.observe(paragraph)
  }
}

function onWindowResize(): void {
  scheduleWrappedBoxRefresh()
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
  scheduleWrappedBoxRefresh()
}

function onImageError(image: FloatingImageEntry, event: Event): void {
  const target = event.currentTarget as HTMLImageElement | null
  if (target) target.dataset.imageState = "error"
  const next = new Set(failedImageSourceKeys.value)
  next.add(image.failureKey)
  failedImageSourceKeys.value = next
}

// ── Mouse handlers ─────────────────────────────────────────────────
function onImageMouseDown(image: FloatingImageEntry, event: MouseEvent): void {
  if (!props.editable || image.behindDocument) return
  event.preventDefault()
  props.controller.selectParagraph(image.nodeIndex)
  if (image.lockedAnchor) return
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
  if (!props.editable || image.lockedAnchor) return
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
  const scale = pointerScale.value
  if (draggingImage.value) {
    const image = draggingImage.value
    const xPx = Math.round(dragStartPosX.value + (event.clientX - dragStartX.value) / scale)
    const yPx = Math.round(dragStartPosY.value + (event.clientY - dragStartY.value) / scale)
    dragPreview.value = {
      key: image.key,
      xPx: Math.max(-image.widthPx + 16, Math.min(props.pageWidthPx - 16, xPx)),
      yPx: Math.max(-image.heightPx + 16, Math.min(props.pageHeightPx - 16, yPx)),
    }
  }

  if (resizingImage.value) {
    const image = resizingImage.value
    const dx = (event.clientX - resizeStartPos.value.x) / scale
    const dy = (event.clientY - resizeStartPos.value.y) / scale
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

const imageGroupsKey = computed(() =>
  floatingImages.value.map((image) => [
    image.key,
    image.flowWrapped,
    image.behindDocument,
    image.xPx,
    image.yPx,
    image.widthPx,
    image.heightPx,
  ].join(":")).join("|")
)

// A new document model can reuse the same node/child position. Failure state
// belongs to the image source, not to that position, so reset it on switch.
watch(() => props.controller.model, () => {
  if (failedImageSourceKeys.value.size > 0) {
    failedImageSourceKeys.value = new Set<string>()
  }
})

// Zoom does not resize the surface's layout box (the viewer applies a CSS
// transform), so the hit box has to be re-measured when the factor changes.
watch(() => props.zoomFactor, () => scheduleWrappedBoxRefresh())

watch(
  imageGroupsKey,
  () => {
    void nextTick(() => {
      scheduleWrappedBoxRefresh()
      observeWrappedLayout()
    })
  },
  { immediate: true }
)

onMounted(() => {
  observeWrappedLayout()
  if (typeof window !== "undefined") {
    window.addEventListener("resize", onWindowResize)
  }
  // Late font loading reflows the text above a wrapped image.
  const fonts = (globalThis as { document?: Document }).document?.fonts
  if (fonts && typeof fonts.ready?.then === "function") {
    void fonts.ready.then(() => scheduleWrappedBoxRefresh())
  }
  scheduleWrappedBoxRefresh()
})

// Any re-render can move the wrapped image (text edit, font swap, decode), so
// re-measure once the update has been flushed.
onUpdated(() => {
  scheduleWrappedBoxRefresh()
  observeWrappedLayout()
})

onUnmounted(() => {
  finishPointerInteraction(false)
  failedImageSourceKeys.value.clear()
  if (typeof window !== "undefined") {
    window.removeEventListener("resize", onWindowResize)
  }
  resizeObserver?.disconnect()
  resizeObserver = null
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
