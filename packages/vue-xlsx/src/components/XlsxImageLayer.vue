<template>
  <div class="xlsx-image-layer" data-testid="xlsx-image-layer" :style="layerStyle">
    <div
      v-for="pane in panes"
      :key="pane.key"
      class="xlsx-image-layer__pane"
      :data-grid-pane="pane.key"
      :style="pane.style"
    >
      <div
        v-for="item in pane.items"
        :key="item.image.id"
        class="xlsx-image-layer__item"
        data-testid="xlsx-image-item"
        data-drawing-kind="image"
        :data-drawing-id="item.image.id"
        tabindex="0"
        :class="{
          'xlsx-image-layer__item--selected': item.image.id === selectedImageId,
          'xlsx-image-layer__item--gesture': interaction.isActive(item.image.id),
        }"
        :style="item.style"
        @pointerdown.stop="onImagePointerDown(item, $event)"
        @contextmenu.prevent.stop="onImageContextMenu(item, $event)"
        @keydown="onImageKeydown(item.image, $event)"
      >
        <div class="xlsx-image-layer__content">
          <img
            v-if="item.image.src && !failedImageIds.has(item.image.id) && permittedImageIds.has(item.image.id)"
            :src="item.image.src"
            :alt="item.image.name ?? m.xlsx.image.alt"
            class="xlsx-image-layer__img"
            :style="imageCropStyle(item.image)"
            draggable="false"
            decoding="async"
            data-image-state="loading"
            @load="onImageLoad(item.image, $event)"
            @error="onImageError(item.image, $event)"
          />
          <div
            v-else
            class="xlsx-image-layer__placeholder"
            :data-image-state="failedImageIds.has(item.image.id) ? 'error' : 'queued'"
          >🖼️</div>
        </div>
        <div
          v-if="item.image.id === selectedImageId"
          class="xlsx-image-layer__handles"
          data-testid="xlsx-image-handles"
        >
          <div
            v-for="handle in DRAWING_RESIZE_HANDLES"
            :key="handle"
            class="xlsx-image-layer__handle"
            data-testid="xlsx-image-resize-handle"
            :data-handle="handle"
            :style="handleStyle(handle)"
            @pointerdown.stop="onHandlePointerDown(item, handle, $event)"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onUnmounted, ref, watch, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxImage, XlsxImageRect, XlsxImageResizeHandlePosition } from "@arcships/xlsx-core";
import { createGridPaneLayout, gridBodyLayerStyle, resolveAnchorRect, type GridPaneKey } from "./grid-metrics";
import { DRAWING_RESIZE_HANDLES, drawingHandleStyle, useDrawingGestures } from "./drawing-interaction";
import { useXlsxMessages } from "../i18n";

const props = defineProps<{
  controller: XlsxViewerController;
  showImages?: boolean;
  scrollLeft?: number;
  scrollTop?: number;
  viewportWidth?: number;
  viewportHeight?: number;
}>();

const emit = defineEmits<{
  drawingContextMenu: [payload: { id: string; kind: "image"; clientX: number; clientY: number }];
}>();

const layerStyle = gridBodyLayerStyle;
const m = useXlsxMessages();
const selectedImageId = computed(() => props.controller.selectedImageId);
const failedImageIds = ref(new Set<string>());
const permittedImageIds = ref(new Set<string>());
const activeImageIds = new Set<string>();
const completedImageIds = new Set<string>();
const knownSources = new Map<string, string>();

interface ImageItem {
  image: XlsxImage;
  pane: GridPaneKey;
  rect: XlsxImageRect;
  style: CSSProperties;
}

const paneLayout = computed(() => createGridPaneLayout(
  props.controller.activeSheet,
  props.controller.zoomScale,
  props.scrollLeft ?? 0,
  props.scrollTop ?? 0,
));

const interaction = useDrawingGestures({
  onSelect: (id) => props.controller.selectImage(id),
  onMove: (id, deltaX, deltaY) => props.controller.moveImageBy(id, deltaX, deltaY),
  onResize: (id, handle, deltaX, deltaY) => props.controller.resizeImageBy(id, handle, deltaX, deltaY),
});

function numericPx(value: string | number | undefined): number {
  return typeof value === "number" ? value : Number.parseFloat(value ?? "0") || 0;
}

function imageTransform(image: XlsxImage): string | undefined {
  const parts: string[] = [];
  if (image.rotationDeg) parts.push(`rotate(${image.rotationDeg}deg)`);
  if (image.flipH) parts.push("scaleX(-1)");
  if (image.flipV) parts.push("scaleY(-1)");
  return parts.length ? parts.join(" ") : undefined;
}

/**
 * Crop a picture by scaling it past its frame and offsetting the hidden edge,
 * mirroring DrawingML `a:srcRect` (fractions of the intrinsic image).
 */
function imageCropStyle(image: XlsxImage): CSSProperties | undefined {
  const crop = image.srcRect;
  if (!crop) return undefined;
  const visibleWidth = Math.max(0.01, 1 - crop.left - crop.right);
  const visibleHeight = Math.max(0.01, 1 - crop.top - crop.bottom);
  return {
    height: `${100 / visibleHeight}%`,
    left: `${-(crop.left / visibleWidth) * 100}%`,
    top: `${-(crop.top / visibleHeight) * 100}%`,
    width: `${100 / visibleWidth}%`,
  };
}

function handleStyle(handle: XlsxImageResizeHandlePosition): CSSProperties {
  return drawingHandleStyle(handle);
}

const imageItems = computed<ImageItem[]>(() => {
  if (!props.showImages) return [];
  const layout = paneLayout.value;
  const images = props.controller.images as XlsxImage[];
  return images
    .filter((img) => img.anchor)
    .flatMap((image) => {
      const rect = resolveAnchorRect(image.anchor, props.controller.activeSheet, props.controller.zoomScale);
      const isSelected = image.id === selectedImageId.value;
      const placement = layout.place(rect);
      if (!layout.isVisible(rect, placement.pane, props.viewportWidth ?? 0, props.viewportHeight ?? 0)) {
        return [];
      }
      const preview = interaction.previewFor(image.id);
      const style: CSSProperties = {
        ...placement.style,
        position: "absolute",
        pointerEvents: "auto",
        cursor: isSelected ? "move" : "pointer",
        outline: isSelected ? "2px solid #2563eb" : "1px solid transparent",
        outlineOffset: "-1px",
        transform: imageTransform(image),
        zIndex: Math.max(0, Math.round(image.zIndex ?? 0)),
      };
      if (preview) {
        style.left = `${numericPx(placement.style.left) + preview.left - rect.left}px`;
        style.top = `${numericPx(placement.style.top) + preview.top - rect.top}px`;
        style.width = `${preview.width}px`;
        style.height = `${preview.height}px`;
      }
      return [{ image, pane: placement.pane, rect, style } satisfies ImageItem];
    });
});

const panes = computed(() => paneLayout.value.panes.map((pane) => ({
  ...pane,
  items: imageItems.value.filter((item) => item.pane === pane.key),
})));

function scheduleImageDecodes(): void {
  const visibleIds = new Set(imageItems.value.map((item) => item.image.id));
  for (const id of [...activeImageIds, ...completedImageIds]) {
    if (!visibleIds.has(id)) {
      activeImageIds.delete(id);
      completedImageIds.delete(id);
      knownSources.delete(id);
    }
  }
  const nextPermitted = new Set(permittedImageIds.value);
  for (const id of nextPermitted) {
    if (!visibleIds.has(id)) nextPermitted.delete(id);
  }
  const limit = Math.max(1, Math.floor(props.controller.maxConcurrentImageDecodes ?? 4));
  for (const item of imageItems.value) {
    const { id, src } = item.image;
    const previousSource = knownSources.get(id);
    if (previousSource !== undefined && previousSource !== src) {
      activeImageIds.delete(id);
      completedImageIds.delete(id);
      nextPermitted.delete(id);
      const nextFailed = new Set(failedImageIds.value);
      nextFailed.delete(id);
      failedImageIds.value = nextFailed;
    }
    knownSources.set(id, src);
    if (
      !src ||
      failedImageIds.value.has(id) ||
      completedImageIds.has(id) ||
      activeImageIds.has(id) ||
      activeImageIds.size >= limit
    ) continue;
    activeImageIds.add(id);
    nextPermitted.add(id);
  }
  permittedImageIds.value = nextPermitted;
}

function onImageLoad(image: XlsxImage, event: Event): void {
  const target = event.currentTarget as HTMLImageElement | null;
  if (target) target.dataset.imageState = "ready";
  activeImageIds.delete(image.id);
  completedImageIds.add(image.id);
  scheduleImageDecodes();
}

function onImageError(image: XlsxImage, event: Event): void {
  const target = event.currentTarget as HTMLImageElement | null;
  if (target) target.dataset.imageState = "error";
  activeImageIds.delete(image.id);
  const nextPermitted = new Set(permittedImageIds.value);
  nextPermitted.delete(image.id);
  permittedImageIds.value = nextPermitted;
  const nextFailed = new Set(failedImageIds.value);
  nextFailed.add(image.id);
  failedImageIds.value = nextFailed;
  props.controller.reportImageDecodeError?.(image.id);
  scheduleImageDecodes();
}

watch(
  () => [props.showImages, ...(props.controller.images as XlsxImage[]).map((image) => `${image.id}:${image.src}`)],
  scheduleImageDecodes,
  { immediate: true },
);

onUnmounted(() => {
  activeImageIds.clear();
  completedImageIds.clear();
  knownSources.clear();
  failedImageIds.value.clear();
  permittedImageIds.value.clear();
});

function onImagePointerDown(item: ImageItem, event: PointerEvent) {
  interaction.beginMove(item.image.id, item.rect, event);
}

function onImageContextMenu(item: ImageItem, event: MouseEvent) {
  props.controller.selectImage(item.image.id);
  emit("drawingContextMenu", {
    id: item.image.id,
    kind: "image",
    clientX: event.clientX,
    clientY: event.clientY,
  });
}

function onHandlePointerDown(item: ImageItem, handle: XlsxImageResizeHandlePosition, event: PointerEvent) {
  interaction.beginResize(item.image.id, item.rect, handle, event);
}

function onImageKeydown(image: XlsxImage, event: KeyboardEvent) {
  if (props.controller.selectedImageId !== image.id) {
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
    event.preventDefault();
    event.stopPropagation();
    void props.controller.copySelectionToClipboard?.();
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "x") {
    event.preventDefault();
    event.stopPropagation();
    props.controller.cutSelection?.();
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
    event.preventDefault();
    event.stopPropagation();
    void props.controller.pasteFromClipboard?.();
    return;
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
    event.preventDefault();
    event.stopPropagation();
    props.controller.duplicateDrawing?.();
    return;
  }
  const step = event.shiftKey ? 10 : 1;
  const nudge: Record<string, [number, number]> = {
    ArrowLeft: [-1, 0],
    ArrowRight: [1, 0],
    ArrowUp: [0, -1],
    ArrowDown: [0, 1],
  };
  const delta = nudge[event.key];
  if (delta) {
    event.preventDefault();
    event.stopPropagation();
    props.controller.moveImageBy(image.id, delta[0] * step, delta[1] * step);
    return;
  }
  if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault();
    event.stopPropagation();
    props.controller.deleteImage(image.id);
    return;
  }
  if (event.key === "Escape") {
    event.stopPropagation();
    if (props.controller.cancelCut?.()) return;
    props.controller.clearSelectedImage();
  }
}
</script>

<style scoped>
/* Geometry (inset past the headers) comes from `gridBodyLayerStyle`; each
   freeze pane clips its own images so a scrolled-away image never paints over
   frozen panes, the headers, or the host UI around the surface. The layer is
   its own stacking context so pane/object z-index values cannot escape it. */
.xlsx-image-layer {
  pointer-events: none; position: absolute;
  isolation: isolate; z-index: 2;
}
.xlsx-image-layer__pane {
  overflow: hidden; pointer-events: none; position: absolute;
}
.xlsx-image-layer__item { background: transparent; overflow: visible; transform-origin: center; }
.xlsx-image-layer__content { inset: 0; overflow: hidden; position: absolute; }
.xlsx-image-layer__item--selected { box-shadow: 0 0 0 2px #2563eb; }
.xlsx-image-layer__item--gesture { cursor: move; }
.xlsx-image-layer__img {
  display: block; height: 100%; left: 0; object-fit: fill;
  pointer-events: none; position: absolute; top: 0; width: 100%;
}
.xlsx-image-layer__placeholder {
  align-items: center; background: rgba(128,128,128,0.1);
  display: flex; font-size: 24px; height: 100%;
  justify-content: center; width: 100%;
}
.xlsx-image-layer__handles { inset: 0; pointer-events: none; position: absolute; }
.xlsx-image-layer__handle {
  background: #ffffff; border: 1px solid #2563eb; box-sizing: border-box;
  height: 8px; pointer-events: auto; position: absolute; width: 8px;
  z-index: 1000000;
}
</style>
