<template>
  <div class="xlsx-image-layer" data-testid="xlsx-image-layer">
    <div
      v-for="item in imageItems"
      :key="item.image.id"
      class="xlsx-image-layer__item"
      data-testid="xlsx-image-item"
      :class="{
        'xlsx-image-layer__item--selected': item.image.id === selectedImageId,
      }"
      :style="item.style"
      @pointerdown.stop="onImagePointerDown(item.image, $event)"
    >
      <img
        v-if="item.image.src && !failedImageIds.has(item.image.id) && permittedImageIds.has(item.image.id)"
        :src="item.image.src"
        :alt="item.image.name ?? 'image'"
        class="xlsx-image-layer__img"
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
  </div>
</template>

<script setup lang="ts">
import { computed, onUnmounted, ref, watch, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxImage, XlsxImageAnchor, XlsxSheetData } from "@arcships/xlsx-core";
import { anchorToAbsoluteRect, emuToPixels } from "@arcships/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  showImages?: boolean;
  scrollLeft?: number;
  scrollTop?: number;
}>();

const selectedImageId = computed(() => props.controller.selectedImageId);
const failedImageIds = ref(new Set<string>());
const permittedImageIds = ref(new Set<string>());
const activeImageIds = new Set<string>();
const completedImageIds = new Set<string>();
const knownSources = new Map<string, string>();

interface ImageItem {
  image: XlsxImage;
  style: CSSProperties;
}

function frozenExtent(sheet: XlsxSheetData | null, axis: "column" | "row", zoomScale: number): number {
  if (!sheet?.freezePanes) return 0;
  const limit = axis === "column" ? sheet.freezePanes.col : sheet.freezePanes.row;
  const indices = axis === "column" ? sheet.visibleCols : sheet.visibleRows;
  const sizes = axis === "column" ? sheet.colWidths : sheet.rowHeights;
  const fallback = axis === "column" ? sheet.defaultColWidthPx : sheet.defaultRowHeightPx;
  const zoom = Math.max(0.1, zoomScale / 100);
  return indices.reduce((total, actualIndex, displayIndex) =>
    actualIndex < limit ? total + (sizes[displayIndex] || fallback) * zoom : total,
  0);
}

function resolveAnchorRect(anchor: XlsxImageAnchor, sheet: XlsxSheetData | null, zoomScale: number) {
  const rect = anchorToAbsoluteRect(anchor, sheet);
  const zoom = Math.max(0.1, zoomScale / 100);
  return {
    left: emuToPixels(rect.x) * zoom,
    top: emuToPixels(rect.y) * zoom,
    width: Math.max(1, emuToPixels(rect.cx) * zoom),
    height: Math.max(1, emuToPixels(rect.cy) * zoom),
  };
}

const imageItems = computed<ImageItem[]>(() => {
  if (!props.showImages) return [];
  const images = props.controller.images as XlsxImage[];
  return images
    .filter((img) => img.anchor)
    .map((image) => {
      const rect = resolveAnchorRect(image.anchor, props.controller.activeSheet, props.controller.zoomScale);
      const isSelected = image.id === selectedImageId.value;
      const frozenWidth = frozenExtent(props.controller.activeSheet, "column", props.controller.zoomScale);
      const frozenHeight = frozenExtent(props.controller.activeSheet, "row", props.controller.zoomScale);
      const scrollX = rect.left < frozenWidth ? 0 : props.scrollLeft ?? 0;
      const scrollY = rect.top < frozenHeight ? 0 : props.scrollTop ?? 0;
      return {
        image,
        style: {
          left: `${rect.left + 48 - scrollX}px`,
          top: `${rect.top + 24 - scrollY}px`,
          width: `${rect.width}px`,
          height: `${rect.height}px`,
          position: "absolute",
          pointerEvents: "auto",
          cursor: "pointer",
          outline: isSelected ? "2px solid #2563eb" : "1px solid transparent",
          outlineOffset: "-1px",
          zIndex: isSelected ? 5 : 2,
        } satisfies CSSProperties,
      };
    });
});

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

function onImagePointerDown(image: XlsxImage, _event: PointerEvent) {
  props.controller.selectImage(image.id);
}
</script>

<style scoped>
.xlsx-image-layer {
  height: 100%; left: 0; pointer-events: none;
  position: absolute; top: 0; width: 100%;
}
.xlsx-image-layer__item { background: transparent; overflow: hidden; }
.xlsx-image-layer__item--selected { box-shadow: 0 0 0 2px #2563eb; }
.xlsx-image-layer__img { display: block; pointer-events: none; width: 100%; height: 100%; object-fit: contain; }
.xlsx-image-layer__placeholder {
  align-items: center; background: rgba(128,128,128,0.1);
  display: flex; font-size: 24px; height: 100%;
  justify-content: center; width: 100%;
}
</style>
