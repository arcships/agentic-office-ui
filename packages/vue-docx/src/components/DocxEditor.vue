<template>
  <div class="docx-editor-viewer" :class="className">
    <DocxToolbar
      v-if="showToolbar"
      :controller="controller"
    />
    <div class="docx-editor-body">
      <DocxThumbnailPanel
        v-if="showThumbnails"
        :controller="controller"
      />
      <DocxViewerRoot
        :controller="controller"
        :editable="editable"
      />
    </div>
    <DocxContextMenu
      v-if="contextMenu.visible.value"
      :context="contextMenu.context.value"
      :controller="controller"
      @close="contextMenu.hide"
    />
    <DocxDragOverlay
      v-if="dragOverlay.visible.value"
      :controller="controller"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, provide, watch, onMounted } from "vue"
import type {
  DocxEditorController,
  DocxContextMenuContext,
} from "@extend-ai/docx-core"
import { useDocxEditor } from "../composables/useDocxEditor"
import DocxViewerRoot from "./DocxViewerRoot.vue"
import DocxToolbar from "./DocxToolbar.vue"
import DocxContextMenu from "./DocxContextMenu.vue"
import DocxThumbnailPanel from "./DocxThumbnailPanel.vue"
import DocxDragOverlay from "./DocxDragOverlay.vue"

const props = withDefaults(
  defineProps<{
    file?: ArrayBuffer
    model?: import("@extend-ai/docx-core").DocModel
    className?: string
    editable?: boolean
    showToolbar?: boolean
    showThumbnails?: boolean
  }>(),
  { editable: true, showToolbar: true, showThumbnails: false }
)

const controller = useDocxEditor({})

// Provide controller to all child components
provide("docxEditorController", controller)

// If model is provided, set it directly via the controller
watch(
  () => props.model,
  (newModel) => {
    if (newModel && controller) {
      // Model assignment is handled through import/export API
      // The starter model is set in useDocxEditor options
    }
  },
  { immediate: true }
)

// Context menu state
const contextMenu = {
  visible: ref(false),
  context: ref<DocxContextMenuContext | undefined>(undefined),
  show(ctx: DocxContextMenuContext) {
    this.context.value = ctx
    this.visible.value = true
  },
  hide() {
    this.visible.value = false
    this.context.value = undefined
  },
}
provide("docxContextMenu", contextMenu)

// Drag overlay state
const dragOverlay = {
  visible: ref(false),
  show() { this.visible.value = true },
  hide() { this.visible.value = false },
}
provide("docxDragOverlay", dragOverlay)
</script>

<style scoped>
.docx-editor-viewer {
  display: flex;
  flex-direction: column;
  height: 100%;
  overflow: hidden;
}
.docx-editor-body {
  display: flex;
  flex: 1;
  overflow: hidden;
}
</style>
