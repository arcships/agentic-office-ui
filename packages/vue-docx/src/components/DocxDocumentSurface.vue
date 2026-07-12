<template>
  <DocxViewerRoot
    ref="viewerRootRef"
    :model="model"
    :controller="controller"
    :editable="editable"
    :layout-options="layoutOptions"
    :theme="theme"
    :zoom-scale="zoomScale"
    :fit-width="fitWidth"
    :show-tracked-changes="showTrackedChanges"
    :show-comments="showComments"
    :search-query="searchQuery"
    :active-search-node-index="activeSearchNodeIndex"
    @page-count-change="emit('pageCountChange', $event)"
    @visible-page-range="emit('visiblePageRange', $event)"
    @context-menu="emit('contextMenu', $event)"
    @selection-change="emit('selectionChange', $event)"
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
    fitWidth?: boolean
    showTrackedChanges?: boolean
    showComments?: boolean
    searchQuery?: string
    activeSearchNodeIndex?: number
  }>(),
  { editable: false, zoomScale: 100 }
)

const emit = defineEmits<{
  pageCountChange: [count: number]
  visiblePageRange: [range: { startPageIndex: number; endPageIndex: number }]
  contextMenu: [ctx: { pageIndex: number; clientX: number; clientY: number }]
  selectionChange: [sel: { kind: string; text?: string; nodeIndex?: number }]
  objectClick: [obj: { kind: "image" | "table"; nodeIndex?: number }]
}>()

const viewerRootRef = ref<InstanceType<typeof DocxViewerRoot>>()

defineExpose({
  scrollToPage(pageIndex: number): void {
    viewerRootRef.value?.scrollToPage(pageIndex)
  },
  scrollToNode(nodeIndex: number): void {
    viewerRootRef.value?.scrollToNode(nodeIndex)
  },
  get scrollContainer() {
    return viewerRootRef.value?.scrollContainer
  },
})
</script>
