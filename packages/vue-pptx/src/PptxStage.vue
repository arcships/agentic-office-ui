<template>
  <div ref="elementRef" class="pptx-stage" data-testid="pptx-stage" />
</template>

<script setup lang="ts">
import { ref } from "vue"

const elementRef = ref<HTMLElement | null>(null)

defineExpose({
  get element(): HTMLElement | null {
    return elementRef.value
  },
})
</script>
