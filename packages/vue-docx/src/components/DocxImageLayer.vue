<template>
  <div v-if="floatingImages.length > 0" class="docx-image-layer">
    <img
      v-for="image of floatingImages"
      :key="image.id"
      :src="image.src"
      :alt="image.alt ?? 'Floating image'"
      :style="imageStyle(image)"
      class="docx-floating-image"
    />
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocxEditorController, LayoutPage } from "@extend-ai/docx-core"

const props = defineProps<{
  page: LayoutPage
  controller?: DocxEditorController
}>()

interface FloatingImage {
  id: string
  src?: string
  alt?: string
  widthPx: number
  heightPx: number
}

const floatingImages = computed<FloatingImage[]>(() => {
  const images: FloatingImage[] = []
  for (const block of props.page.blocks) {
    if (block.kind === "paragraph") {
      for (const run of block.runs) {
        if (run.kind === "image" && run.src) {
          images.push({
            id: run.id,
            src: run.src,
            alt: run.alt,
            widthPx: run.widthPx ?? 100,
            heightPx: run.heightPx ?? 100,
          })
        }
      }
    }
  }
  return images
})

function imageStyle(image: FloatingImage): Record<string, string> {
  return {
    maxWidth: `${image.widthPx}px`,
    maxHeight: `${image.heightPx}px`,
  }
}
</script>

<style scoped>
.docx-image-layer {
  padding: 4px 0;
}
.docx-floating-image {
  display: block;
  margin: 4px 0;
}
</style>
