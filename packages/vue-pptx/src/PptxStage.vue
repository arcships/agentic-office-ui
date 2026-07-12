<template>
  <div ref="elementRef" class="pptx-stage" data-testid="pptx-stage" @contextmenu.prevent="onContextMenu" />
</template>

<script setup lang="ts">
import { ref } from "vue"

const elementRef = ref<HTMLElement | null>(null)

const emit = defineEmits<{
  contextMenu: [ctx: { clientX: number; clientY: number; containerX: number; containerY: number }]
  selectionChange: [sel: { kind: string; slideIndex?: number }]
  objectClick: [obj: { objectKey: string }]
}>()

function onContextMenu(event: MouseEvent): void {
  const el = elementRef.value
  const rect = el?.getBoundingClientRect()
  emit("contextMenu", {
    clientX: event.clientX,
    clientY: event.clientY,
    containerX: rect ? event.clientX - rect.left : 0,
    containerY: rect ? event.clientY - rect.top : 0,
  })
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
