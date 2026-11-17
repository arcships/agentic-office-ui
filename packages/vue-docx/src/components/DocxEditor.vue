<template>
  <div
    ref="editorRootRef"
    class="docx-editor-viewer"
    :class="className"
    data-testid="docx-editor"
    @keydown="onKeydown"
  >
    <!-- Toolbar -->
    <DocxToolbar
      v-if="showToolbar"
      :controller="controller"
      :zoom="zoomPercent"
      :read-only="interactionLocked"
      :show-thumbnails="thumbnailPanelVisible"
      :current-page="currentPage"
      :total-pages="totalPages"
      @update:zoom="zoomPercent = $event"
      @toggle-thumbnails="thumbnailPanelVisible = !thumbnailPanelVisible"
      @select-page="onSelectPage($event - 1)"
    />

    <!-- Editor body (thumbnails + viewer root) -->
    <div class="docx-editor-body">
      <!-- Thumbnail panel -->
      <DocxThumbnailPanel
        v-if="thumbnailPanelVisible && !pageLimitError"
        :controller="controller"
        :current-page="currentPage"
        :total-pages="totalPages"
        :show-toggle="false"
        @select-page="onSelectPage"
      />

      <!-- Viewer root (main scrollable area with page surfaces) -->
      <DocxDocumentSurface
        v-if="!pageLimitError"
        ref="viewerRootRef"
        :model="renderedModel"
        :controller="controller"
        :editable="editable"
        :zoom-scale="zoomPercent"
        @page-count-change="onPageCountChange"
        @visible-page-range="onVisiblePageRange"
        @context-menu="onDocumentContextMenu"
      />
      <div
        v-else
        class="docx-editor-limit-error"
        data-testid="editor-load-error"
        :data-error-code="pageLimitError.code"
        :data-error-phase="pageLimitError.phase"
        :data-error-limit="pageLimitError.limit"
        :data-error-actual="pageLimitError.actual"
        :data-error-allowed="pageLimitError.allowed"
      >
        {{ pageLimitError.message }}
      </div>
    </div>

    <!-- Context menu -->
    <DocxContextMenu
      v-if="!interactionLocked && contextMenu.visible.value"
      :context="contextMenu.context.value"
      :controller="controller"
      :x="contextMenu.x.value"
      :y="contextMenu.y.value"
      :text="contextMenu.text.value"
      :caret-offset="contextMenu.caretOffset.value"
      @close="contextMenu.hide"
    />

    <!-- Drag overlay -->
    <DocxDragOverlay
      v-if="!interactionLocked && dragOverlay.visible.value"
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
  DocxRuntime,
} from "@arcships/docx-core"
import { createDocxRuntime, DocxImportError } from "@arcships/docx-core"
import { useDocxEditor } from "../composables/useDocxEditor"
import {
  docxTableContextFromNode,
  docxTextHostFromNode,
  readDocxTextHostText,
  syncDomSelectionToEditor,
  textOffsetAtDomPoint,
} from "../composables/editor-selection"
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
    model?: import("@arcships/docx-core").DocModel
    className?: string
    editable?: boolean
    showToolbar?: boolean
    showThumbnails?: boolean
    runtime?: DocxRuntime
  }>(),
  { editable: true, showToolbar: true, showThumbnails: false }
)

const emit = defineEmits<{
  (event: "load-error", error: Error): void
}>()

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
const ownedLimitRuntime = createDocxRuntime()
const effectiveRuntime = computed(() => props.runtime ?? ownedLimitRuntime)
const pageLimitError = ref<DocxImportError | undefined>()

watch(
  () => [renderedModel.value, props.runtime] as const,
  () => { pageLimitError.value = undefined },
)

const interactionLocked = computed(
  () => !props.editable || controller.showTrackedChanges
)

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

// Root element ref keeps context-menu/selection scoping explicit.
const editorRootRef = ref<HTMLElement | undefined>()

// ── Page tracking ──────────────────────────────────────────────────
const currentPage = ref(1)
const totalPages = ref(1)
const zoomPercent = ref(100)
const thumbnailPanelVisible = ref(props.showThumbnails)

watch(() => props.showThumbnails, (visible) => { thumbnailPanelVisible.value = visible })

function onPageCountChange(count: number): void {
  const allowed = effectiveRuntime.value.limits?.maxDocxPages
  if (allowed !== undefined && count > allowed) {
    if (!pageLimitError.value) {
      const normalized = new DocxImportError(
        "LIMIT_EXCEEDED",
        `DOCX 页数 ${count} 超过允许值 ${allowed}。`,
        {
          phase: "layout",
          limit: "maxDocxPages",
          actual: count,
          allowed,
        },
      )
      pageLimitError.value = normalized
      emit("load-error", normalized)
    }
    return
  }
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
interface DocumentContextMenuEvent {
  pageIndex: number
  clientX: number
  clientY: number
  containerX: number
  containerY: number
  target?: EventTarget | null
}

const contextMenu = {
  visible: ref(false),
  context: ref<DocxContextMenuContext | undefined>(undefined),
  x: ref(0),
  y: ref(0),
  text: ref<string | undefined>(undefined),
  caretOffset: ref<number | undefined>(undefined),
  show(
    x: number,
    y: number,
    ctx: DocxContextMenuContext,
    text?: string,
    caretOffset?: number,
  ) {
    this.x.value = x
    this.y.value = y
    // The menu anchors on the context so custom renderers and standalone
    // consumers see the same pointer position as the built-in menu.
    this.context.value = { ...ctx, position: { x, y } }
    this.text.value = text
    this.caretOffset.value = caretOffset
    this.visible.value = true
  },
  hide() {
    this.visible.value = false
    this.context.value = undefined
    this.text.value = undefined
    this.caretOffset.value = undefined
  },
}

function onDocumentContextMenu(payload: DocumentContextMenuEvent): void {
  if (interactionLocked.value) return

  const root = viewerRootRef.value?.scrollContainer
  if (root) syncDomSelectionToEditor(controller, root)

  const target = typeof Node !== "undefined" && payload.target instanceof Node
    ? payload.target
    : null
  const tableContext = docxTableContextFromNode(target)
  const textHost = docxTextHostFromNode(target)
  const range = controller.activeTextRange

  const context: DocxContextMenuContext = tableContext
    ? { kind: "table", tableContext }
    : { kind: "text" }
  if (range) context.activeTextRange = range

  let text: string | undefined
  let caretOffset: number | undefined
  if (textHost) {
    text = readDocxTextHostText(textHost)
    caretOffset = textOffsetAtDomPoint(textHost, payload.clientX, payload.clientY)
    if (caretOffset === undefined && range) {
      caretOffset = range.start.offset
    }
    if (caretOffset === undefined) caretOffset = text.length
  }

  contextMenu.show(payload.clientX, payload.clientY, context, text, caretOffset)
}

// ── Drag overlay ───────────────────────────────────────────────────
const dragOverlay = {
  visible: ref(false),
  show() { this.visible.value = true },
  hide() { this.visible.value = false },
}
provide("docxDragOverlay", dragOverlay)

// ── Keyboard shortcuts ─────────────────────────────────────────────
// The listener lives on the editor root, so host-page inputs and other editor
// instances never see these shortcuts. Commands are also suppressed while the
// document is read-only or while tracked-changes review locks editing.
function onKeydown(event: KeyboardEvent): void {
  if (!props.editable || event.defaultPrevented || event.isComposing) return

  const target = event.target as HTMLElement | null
  const tagName = target?.tagName?.toLowerCase()
  if (tagName === "input" || tagName === "textarea" || tagName === "select") return
  if (target?.closest?.(".docx-toolbar-shell")) return

  const mod = event.metaKey || event.ctrlKey
  if (!mod) return

  const key = event.key.toLowerCase()
  if (key === "s") {
    event.preventDefault()
    controller.exportDocx()
    return
  }
  if (interactionLocked.value) return

  if (key === "b") {
    event.preventDefault()
    controller.toggleBold()
  } else if (key === "i") {
    event.preventDefault()
    controller.toggleItalic()
  } else if (key === "u") {
    event.preventDefault()
    controller.toggleUnderline()
  } else if (key === "z" && !event.shiftKey) {
    event.preventDefault()
    controller.undo()
  } else if (key === "z" && event.shiftKey) {
    event.preventDefault()
    controller.redo()
  }
}

// ── Unsaved changes guard ──────────────────────────────────────────
function onBeforeUnload(event: BeforeUnloadEvent): void {
  if (!controller.isDirty) return
  event.preventDefault()
  event.returnValue = ""
}

let beforeUnloadAttached = false

function syncBeforeUnload(): void {
  const shouldAttach = typeof window !== "undefined" && props.editable && controller.isDirty
  if (shouldAttach && !beforeUnloadAttached) {
    window.addEventListener("beforeunload", onBeforeUnload)
    beforeUnloadAttached = true
  } else if (!shouldAttach && beforeUnloadAttached) {
    window.removeEventListener("beforeunload", onBeforeUnload)
    beforeUnloadAttached = false
  }
}

watch(
  [() => props.editable, () => controller.isDirty],
  syncBeforeUnload,
  { immediate: true },
)

watch(
  () => interactionLocked.value,
  (locked) => {
    if (!locked) return
    contextMenu.hide()
    dragOverlay.hide()
    controller.clearSelectionSession()
  },
)

onMounted(() => {
  syncBeforeUnload()
})

onUnmounted(() => {
  if (beforeUnloadAttached) {
    window.removeEventListener("beforeunload", onBeforeUnload)
    beforeUnloadAttached = false
  }
  ownedLimitRuntime.dispose()
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
.docx-editor-limit-error {
  align-items: center;
  color: #b91c1c;
  display: flex;
  flex: 1;
  justify-content: center;
  padding: 24px;
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
