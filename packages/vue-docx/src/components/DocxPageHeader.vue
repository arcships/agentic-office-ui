<template>
  <header v-if="hasContent" class="docx-page-header">
    <DocxPageBody :blocks="headerBlocks" />
  </header>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { LayoutPage } from "@extend-ai/docx-core"
import DocxPageBody from "./DocxPageBody"

const props = defineProps<{ page: LayoutPage }>()

const hasContent = computed(() => headerBlocks.value.length > 0)
const headerBlocks = computed(() => [])
</script>

<style scoped>
.docx-page-header {
  border-bottom: 1px solid #e5e7eb;
  padding-bottom: 8px;
  margin-bottom: 8px;
}
</style>
