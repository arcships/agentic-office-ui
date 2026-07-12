<template>
  <div ref="elementRef" class="pptx-stage" data-testid="pptx-stage" @contextmenu.prevent="onContextMenu" />
</template>

<script setup lang="ts">
import { ref } from "vue"

const elementRef = ref<HTMLElement | null>(null)

const emit = defineEmits<{
  contextMenu: [ctx: { clientX: number; clientY: number }]
}>()

function onContextMenu(event: MouseEvent): void {
  emit("contextMenu", { clientX: event.clientX, clientY: event.clientY })
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
