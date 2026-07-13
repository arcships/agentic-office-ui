<template>
  <div
    ref="elementRef"
    class="pptx-stage"
    data-testid="pptx-stage"
    @click="onClick"
    @contextmenu.prevent="onContextMenu"
  />
</template>

<script setup lang="ts">
import { ref } from "vue"
import type {
  PptxStageContextMenu,
  PptxStageObjectClick,
  PptxStageSelection,
} from "./headless-types"

const elementRef = ref<HTMLElement | null>(null)

const emit = defineEmits<{
  contextMenu: [ctx: PptxStageContextMenu]
  selectionChange: [selection: PptxStageSelection]
  objectClick: [object: PptxStageObjectClick]
}>()

interface StageHit {
  slideIndex: number
  objectKey?: string
}

function resolveHit(event: MouseEvent): StageHit | null {
  const target = event.target as Element | null
  if (!target || typeof target.closest !== "function") return null
  const slideElement = target.closest<HTMLElement>("[data-slide-index]")
  const slideIndex = Number(slideElement?.dataset.slideIndex)
  if (!Number.isInteger(slideIndex) || slideIndex < 0) return null
  const objectElement = target.closest<HTMLElement>("[data-pptx-object-key]")
  return {
    slideIndex,
    objectKey: objectElement?.dataset.pptxObjectKey || undefined,
  }
}

function onClick(event: MouseEvent): void {
  const hit = resolveHit(event)
  if (!hit) return
  emit("selectionChange", { kind: "slide", slideIndex: hit.slideIndex })
  if (hit.objectKey) {
    emit("objectClick", {
      kind: "object",
      slideIndex: hit.slideIndex,
      objectKey: hit.objectKey,
    })
  }
}

function onContextMenu(event: MouseEvent): void {
  const hit = resolveHit(event)
  if (!hit) return
  const el = elementRef.value
  const rect = el?.getBoundingClientRect?.()
  const position = {
    slideIndex: hit.slideIndex,
    clientX: event.clientX,
    clientY: event.clientY,
    containerX: rect ? event.clientX - rect.left : 0,
    containerY: rect ? event.clientY - rect.top : 0,
  }
  emit("contextMenu", hit.objectKey
    ? { ...position, kind: "object", objectKey: hit.objectKey }
    : { ...position, kind: "slide" })
}

defineExpose({
  get element(): HTMLElement | null {
    return elementRef.value
  },
})
</script>

<style scoped>
.pptx-stage {
  background: var(--pptx-surface-bg, transparent);
}
</style>
