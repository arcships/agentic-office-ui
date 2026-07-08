<template>
  <div class="docx-viewer" :class="className">
    <div v-if="isLoading" class="docx-viewer-loading">Loading DOCX...</div>
    <div v-else-if="error" class="docx-viewer-error">Failed to parse DOCX: {{ error.message }}</div>
    <div v-else-if="!model" class="docx-viewer-empty">{{ emptyState ?? "No DOCX loaded." }}</div>
    <template v-else>
      <section
        v-for="page of pages"
        :key="page.number"
        :data-page="page.number"
        class="docx-page"
        :style="pageStyle"
      >
        <DocxPageBody :blocks="page.blocks" />
      </section>
    </template>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue"
import type { DocModel, LayoutOptions } from "@extend-ai/docx-core"
import { layoutDocument, importDocxBuffer } from "@extend-ai/docx-core"
import { DEFAULT_DOC_PAGE_WIDTH, DEFAULT_DOC_PAGE_HEIGHT } from "@extend-ai/docx-core"
import DocxPageBody from "./DocxPageBody"

const props = withDefaults(
  defineProps<{
    file?: ArrayBuffer
    model?: DocModel
    className?: string
    layoutOptions?: LayoutOptions
    emptyState?: string
  }>(),
  { emptyState: "No DOCX loaded." }
)

const isLoading = ref(false)
const error = ref<Error | undefined>(undefined)
const parsedModel = ref<DocModel | undefined>(undefined)

const resolvedModel = computed(() => props.model ?? parsedModel.value)

watch(
  () => props.file,
  async (newFile) => {
    if (!newFile) {
      parsedModel.value = undefined
      error.value = undefined
      isLoading.value = false
      return
    }
    if (props.model) {
      return
    }
    const docxFile = newFile
    isLoading.value = true
    error.value = undefined
    try {
      const { model } = await importDocxBuffer(docxFile, {
        transferBuffer: false,
      })
      parsedModel.value = model
    } catch (err) {
      error.value = err instanceof Error ? err : new Error("Unknown DOCX parse error")
    } finally {
      isLoading.value = false
    }
  },
  { immediate: true }
)

const resolvedLayoutOptions = computed<LayoutOptions | undefined>(() => {
  if (!resolvedModel.value) return props.layoutOptions
  return props.layoutOptions ?? {}
})

const pages = computed(() => {
  if (!resolvedModel.value) return []
  return layoutDocument(resolvedModel.value, resolvedLayoutOptions.value)
})

const pageWidth = computed(
  () => resolvedLayoutOptions.value?.pageWidth ?? DEFAULT_DOC_PAGE_WIDTH
)
const pageHeight = computed(
  () => resolvedLayoutOptions.value?.pageHeight ?? DEFAULT_DOC_PAGE_HEIGHT
)
const margin = computed(() => resolvedLayoutOptions.value?.margin ?? 72)

const pageStyle = computed(() => ({
  width: `${pageWidth.value}px`,
  minHeight: `${pageHeight.value}px`,
  padding: `${margin.value}px`,
  background: "#fff",
  border: "1px solid #d4d4d4",
  boxShadow: "0 8px 24px rgba(0, 0, 0, 0.08)",
}))
</script>

<style scoped>
.docx-viewer {
  display: grid;
  justify-items: center;
  gap: 16px;
}
.docx-page {
  box-sizing: border-box;
  display: grid;
  gap: 8px;
  align-content: start;
}
.docx-viewer-loading,
.docx-viewer-error,
.docx-viewer-empty {
  padding: 24px;
  color: #6b7280;
  font-size: 14px;
}
.docx-viewer-error {
  color: #dc2626;
}
</style>
