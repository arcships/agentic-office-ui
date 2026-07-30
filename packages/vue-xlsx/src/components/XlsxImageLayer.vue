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
  </div>
</template>

<script setup lang="ts">
import { computed, onUnmounted, ref, watch, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxImage } from "@arcships/xlsx-core";
import { createGridPaneLayout, gridBodyLayerStyle, resolveAnchorRect, type GridPaneKey } from "./grid-metrics";

const props = defineProps<{
  controller: XlsxViewerController;
  showImages?: boolean;
  scrollLeft?: number;
  scrollTop?: number;
}>();

const layerStyle = gridBodyLayerStyle;
const selectedImageId = computed(() => props.controller.selectedImageId);
const failedImageIds = ref(new Set<string>());
const permittedImageIds = ref(new Set<string>());
const activeImageIds = new Set<string>();
const completedImageIds = new Set<string>();
const knownSources = new Map<string, string>();

interface ImageItem {
  image: XlsxImage;
  pane: GridPaneKey;
  style: CSSProperties;
}

const paneLayout = computed(() => createGridPaneLayout(
  props.controller.activeSheet,
  props.controller.zoomScale,
  props.scrollLeft ?? 0,
  props.scrollTop ?? 0,
));

const imageItems = computed<ImageItem[]>(() => {
  if (!props.showImages) return [];
  const layout = paneLayout.value;
  const images = props.controller.images as XlsxImage[];
  return images
    .filter((img) => img.anchor)
    .map((image) => {
      const rect = resolveAnchorRect(image.anchor, props.controller.activeSheet, props.controller.zoomScale);
      const isSelected = image.id === selectedImageId.value;
      const placement = layout.place(rect);
      return {
        image,
        pane: placement.pane,
        style: {
          ...placement.style,
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

function onImagePointerDown(image: XlsxImage, _event: PointerEvent) {
  props.controller.selectImage(image.id);
}
</script>

<style scoped>
/* Geometry (inset past the headers) comes from `gridBodyLayerStyle`; each
   freeze pane clips its own images so a scrolled-away image never paints over
   frozen panes, the headers, or the host UI around the surface. */
.xlsx-image-layer {
  pointer-events: none; position: absolute;
}
.xlsx-image-layer__pane {
  overflow: hidden; pointer-events: none; position: absolute;
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
