<template>
  <DocxViewerRoot
    ref="viewerRootRef"
    :model="model"
    :controller="controller"
    :editable="editable"
    :layout-options="layoutOptions"
    :theme="theme"
    :zoom-scale="zoomScale"
    :show-tracked-changes="showTrackedChanges"
    :show-comments="showComments"
    @page-count-change="emit('pageCountChange', $event)"
    @visible-page-range="emit('visiblePageRange', $event)"
  />
</template>

<script setup lang="ts">
import { ref } from "vue"
import type {
  DocModel,
  DocxDocumentTheme,
  DocxEditorController,
  LayoutOptions,
} from "@arcships/docx-core"
import DocxViewerRoot from "./DocxViewerRoot.vue"

withDefaults(
  defineProps<{
    model: DocModel
    controller?: DocxEditorController
    editable?: boolean
    layoutOptions?: LayoutOptions
    theme?: DocxDocumentTheme
    zoomScale?: number
    showTrackedChanges?: boolean
    showComments?: boolean
  }>(),
  { editable: false, zoomScale: 100 }
)

const emit = defineEmits<{
  pageCountChange: [count: number]
  visiblePageRange: [range: { startPageIndex: number; endPageIndex: number }]
}>()

const viewerRootRef = ref<InstanceType<typeof DocxViewerRoot>>()

defineExpose({
  scrollToPage(pageIndex: number): void {
    viewerRootRef.value?.scrollToPage(pageIndex)
  },
  get scrollContainer() {
    return viewerRootRef.value?.scrollContainer
  },
})
</script>
