<template>
  <div class="xlsx-image-layer">
    <div
      v-for="item in imageItems"
      :key="item.image.id"
      class="xlsx-image-layer__item"
      :class="{
        'xlsx-image-layer__item--selected': item.image.id === selectedImageId,
      }"
      :style="item.style"
      @pointerdown.stop="onImagePointerDown(item.image, $event)"
    >
      <img
        v-if="item.image.src"
        :src="item.image.src"
        :alt="item.image.name ?? 'image'"
        class="xlsx-image-layer__img"
        draggable="false"
      />
      <div v-else class="xlsx-image-layer__placeholder">🖼️</div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import type { XlsxViewerController, XlsxImage, XlsxImageRect, XlsxImageAnchor } from "@extend-ai/xlsx-core";
import { emuToPixels } from "@extend-ai/xlsx-core";

const props = defineProps<{
  controller: XlsxViewerController;
  showImages?: boolean;
}>();

const selectedImageId = computed(() => props.controller.selectedImageId);

interface ImageItem {
  image: XlsxImage;
  style: CSSProperties;
}

function resolveAnchorRect(anchor: XlsxImageAnchor): { left: number; top: number; width: number; height: number } | null {
  if (anchor.kind === "absolute") {
    return {
      left: emuToPixels(anchor.positionEmu.x),
      top: emuToPixels(anchor.positionEmu.y),
      width: emuToPixels(anchor.sizeEmu.cx),
      height: emuToPixels(anchor.sizeEmu.cy),
    };
  }
  if (anchor.kind === "one-cell") {
    return {
      left: emuToPixels(anchor.from.colOffsetEmu),
      top: emuToPixels(anchor.from.rowOffsetEmu),
      width: emuToPixels(anchor.sizeEmu.cx),
      height: emuToPixels(anchor.sizeEmu.cy),
    };
  }
  return null;
}

const imageItems = computed<ImageItem[]>(() => {
  if (!props.showImages) return [];
  const images = props.controller.images as XlsxImage[];
  return images
    .filter((img) => img.anchor)
    .map((image) => {
      const rect = resolveAnchorRect(image.anchor);
      if (!rect) return { image, style: { display: "none" } };
      const isSelected = image.id === selectedImageId.value;
      return {
        image,
        style: {
          left: `${rect.left + 48}px`,
          top: `${rect.top + 24}px`,
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
