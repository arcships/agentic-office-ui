<template>
  <div class="docx-page-wrapper" :data-page="page.number">
    <slot />
  </div>
</template>

<script setup lang="ts">
import type { LayoutPage } from "@extend-ai/docx-core"

defineProps<{
  page: LayoutPage
}>()
</script>

<style scoped>
.docx-page-wrapper {
  box-sizing: border-box;
  background: #fff;
  border: 1px solid #d4d4d4;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
  margin-bottom: 16px;
}
</style>
