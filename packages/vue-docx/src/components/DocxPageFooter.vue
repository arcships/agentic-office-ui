<template>
  <footer v-if="hasContent" class="docx-page-footer">
    <DocxPageBody :blocks="footerBlocks" />
  </footer>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { LayoutPage } from "@extend-ai/docx-core"
import DocxPageBody from "./DocxPageBody"

const props = defineProps<{ page: LayoutPage }>()

const hasContent = computed(() => footerBlocks.value.length > 0)
const footerBlocks = computed(() => [])
</script>

<style scoped>
.docx-page-footer {
  border-top: 1px solid #e5e7eb;
  padding-top: 8px;
  margin-top: 8px;
}
</style>
