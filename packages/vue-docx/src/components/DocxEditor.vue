<template>
  <div class="docx-editor-viewer" :class="className" data-testid="docx-editor">
    <!-- Toolbar -->
    <DocxToolbar
      v-if="showToolbar"
      :controller="controller"
    />

    <!-- Editor body (thumbnails + viewer root) -->
    <div class="docx-editor-body">
      <!-- Thumbnail panel -->
      <DocxThumbnailPanel
        v-if="showThumbnails"
        :controller="controller"
        :current-page="currentPage"
        :total-pages="totalPages"
        @select-page="onSelectPage"
      />

      <!-- Viewer root (main scrollable area with page surfaces) -->
      <DocxDocumentSurface
        ref="viewerRootRef"
        :model="renderedModel"
        :controller="controller"
        :editable="editable"
        @page-count-change="onPageCountChange"
        @visible-page-range="onVisiblePageRange"
      />
    </div>

    <!-- Context menu -->
    <DocxContextMenu
      v-if="editable && contextMenu.visible.value"
      :context="contextMenu.context.value"
      :controller="controller"
      @close="contextMenu.hide"
    />

    <!-- Drag overlay -->
    <DocxDragOverlay
      v-if="editable && dragOverlay.visible.value"
      :controller="controller"
    />

    <!-- Status bar -->
    <div class="docx-editor-statusbar">
      <span class="docx-editor-statusbar-page">
        Page {{ currentPage }} of {{ totalPages }}
      </span>
      <span class="docx-editor-statusbar-status" data-testid="editor-status">
        {{ controller.status }}
      </span>
      <span v-if="controller.fileName" class="docx-editor-statusbar-file" data-testid="editor-file-name">
        {{ controller.fileName }}
      </span>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, provide, watch, onMounted, onUnmounted } from "vue"
import type {
  DocxEditorController,
  DocxContextMenuContext,
} from "@extend-ai/docx-core"
import { useDocxEditor } from "../composables/useDocxEditor"
import DocxDocumentSurface from "./DocxDocumentSurface.vue"
import DocxToolbar from "./DocxToolbar.vue"
import DocxContextMenu from "./DocxContextMenu.vue"
import DocxThumbnailPanel from "./DocxThumbnailPanel.vue"
import DocxDragOverlay from "./DocxDragOverlay.vue"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    editor?: DocxEditorController
    file?: ArrayBuffer
    model?: import("@extend-ai/docx-core").DocModel
    className?: string
    editable?: boolean
    showToolbar?: boolean
    showThumbnails?: boolean
  }>(),
  { editable: true, showToolbar: true, showThumbnails: false }
)

// ── Controller ─────────────────────────────────────────────────────
const internalController = !props.editor
  ? useDocxEditor({
      starterModel: props.model,
      initialFileName: props.file ? "(imported document)" : "(new document)",
    })
  : null
const controller = props.editor ?? internalController!
const renderedModel = computed(() =>
  !props.editor && !props.editable && props.model
    ? props.model
    : controller.model
)

// Provide controller to all child components
provide("docxEditorController", controller)

// ── Import file ────────────────────────────────────────────────────
watch(
  () => props.file,
  async (newFile) => {
    if (!controller || props.editor) return
    if (newFile) {
      await controller.importDocxFile(
        new File([newFile], "document.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })
      )
    } else {
      // Clearing the controlled file must invalidate a pending import. The
      // public newDocument() path owns that cancellation and restores a clean
      // editor state without exposing another controller method.
      controller.newDocument()
    }
  },
  { immediate: true }
)

// ── Viewer root ref ────────────────────────────────────────────────
const viewerRootRef = ref<InstanceType<typeof DocxDocumentSurface> | undefined>()

// ── Page tracking ──────────────────────────────────────────────────
const currentPage = ref(1)
const totalPages = ref(1)

function onPageCountChange(count: number): void {
  totalPages.value = count
  controller.syncPaginationInfo({ currentPage: currentPage.value, totalPages: count })
}

function onVisiblePageRange(range: { startPageIndex: number; endPageIndex: number }): void {
  const midPage = Math.floor((range.startPageIndex + range.endPageIndex) / 2) + 1
  if (midPage !== currentPage.value) {
    currentPage.value = midPage
    controller.syncPaginationInfo({ currentPage: midPage, totalPages: totalPages.value })
  }
}

function onSelectPage(pageIndex: number): void {
  viewerRootRef.value?.scrollToPage(pageIndex)
}

// ── Context menu ───────────────────────────────────────────────────
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

// ── Drag overlay ───────────────────────────────────────────────────
const dragOverlay = {
  visible: ref(false),
  show() { this.visible.value = true },
  hide() { this.visible.value = false },
}
provide("docxDragOverlay", dragOverlay)

// ── Keyboard shortcuts ─────────────────────────────────────────────
function onKeydown(event: KeyboardEvent): void {
  const mod = event.metaKey || event.ctrlKey
  const ctrl = controller

  if (mod && event.key === "b") {
    event.preventDefault()
    ctrl.toggleBold()
  } else if (mod && event.key === "i") {
    event.preventDefault()
    ctrl.toggleItalic()
  } else if (mod && event.key === "u") {
    event.preventDefault()
    ctrl.toggleUnderline()
  } else if (mod && event.key === "z" && !event.shiftKey) {
    event.preventDefault()
    ctrl.undo()
  } else if (mod && event.key === "z" && event.shiftKey) {
    event.preventDefault()
    ctrl.redo()
  } else if (mod && event.key === "s") {
    event.preventDefault()
    ctrl.exportDocx()
  }
}

let isMounted = false
let keydownAttached = false

function attachEditingListeners(): void {
  if (!isMounted || keydownAttached || !props.editable) return
  window.addEventListener("keydown", onKeydown)
  keydownAttached = true
}

function detachEditingListeners(): void {
  if (!keydownAttached) return
  window.removeEventListener("keydown", onKeydown)
  keydownAttached = false
}

watch(
  () => props.editable,
  (editable) => {
    if (editable) {
      attachEditingListeners()
      return
    }
    detachEditingListeners()
    contextMenu.hide()
    dragOverlay.hide()
    controller.clearSelectionSession()
  }
)

onMounted(() => {
  isMounted = true
  attachEditingListeners()
})

onUnmounted(() => {
  detachEditingListeners()
  isMounted = false
})
</script>

<style scoped>
.docx-editor-viewer {
  display: flex;
  flex-direction: column;
  height: 100%;
  overflow: hidden;
  background: #f3f4f6;
}
.docx-editor-body {
  display: flex;
  flex: 1;
  overflow: hidden;
}
.docx-editor-statusbar {
  display: flex;
  align-items: center;
  gap: 16px;
  padding: 4px 12px;
  background: #f9fafb;
  border-top: 1px solid #e5e7eb;
  font-size: 12px;
  color: #6b7280;
  flex-shrink: 0;
}
.docx-editor-statusbar-page {
  font-weight: 500;
}
.docx-editor-statusbar-status {
  color: #9ca3af;
}
.docx-editor-statusbar-file {
  margin-left: auto;
  color: #374151;
  font-weight: 500;
}
</style>
