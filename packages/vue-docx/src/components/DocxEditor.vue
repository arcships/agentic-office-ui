<template>
  <div
    ref="editorRootRef"
    class="docx-editor-viewer"
    :class="[className, { 'docx-editor-viewer--dark': isDarkTheme }]"
    :style="minHeightStyle"
    data-testid="docx-editor"
    @keydown="onKeydown"
    @dblclick="onEditorDoubleClick"
    @pointerover="onEditorPointerOver"
    @pointerout="onEditorPointerOut"
    @pointermove="onEditorPointerMove"
    @pointerleave="onEditorPointerLeave"
  >
    <!-- Toolbar -->
    <DocxToolbar
      v-if="showToolbar"
      :controller="controller"
      :theme="isDarkTheme ? 'dark' : 'light'"
      :zoom="zoomPercent"
      :read-only="interactionLocked"
      :show-thumbnails="thumbnailPanelVisible"
      :show-style-preview="showStylePreview"
      :current-page="currentPage"
      :total-pages="totalPages"
      :find-open="findState.open"
      @update:zoom="zoomPercent = $event"
      @toggle-thumbnails="thumbnailPanelVisible = !thumbnailPanelVisible"
      @select-page="onSelectPage($event - 1)"
      @toggle-find="find.toggleFind()"
    >
      <template #find>
        <OfficeFindBar
          v-if="findState.open"
          ref="findBarRef"
          :query="findState.query"
          :status="findState.status"
          :active-index="findState.activeIndex"
          :result-count="findState.matches.length"
          :placeholder="m.docx.toolbar.searchPlaceholder"
          input-test-id="docx-editor-search-input"
          :case-sensitive="findState.caseSensitive"
          :whole-word="findState.wholeWord"
          :regex="findState.regex"
          :show-replace="!interactionLocked"
          :replacement="findState.replacement"
          :replace-disabled="!find.canReplace.value"
          :replace-all-disabled="!find.canReplaceAll.value"
          :replace-status="findReplaceStatus"
          @update:query="find.setQuery"
          @update:replacement="find.setReplacement"
          @update:case-sensitive="find.setOption('caseSensitive', $event)"
          @update:whole-word="find.setOption('wholeWord', $event)"
          @update:regex="find.setOption('regex', $event)"
          @previous="find.searchPrevious()"
          @next="find.searchNext()"
          @close="find.closeFind()"
          @replace="find.replaceActive()"
          @replace-all="find.replaceAll()"
        />
      </template>
    </DocxToolbar>

    <!-- Editor body (thumbnails + viewer root) -->
    <div class="docx-editor-body">
      <!-- Thumbnail panel -->
      <DocxThumbnailPanel
        v-if="thumbnailPanelVisible && !pageLimitError"
        :controller="controller"
        :is-dark="isDarkTheme"
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
        :theme="isDarkTheme ? 'dark' : 'light'"
        :zoom-scale="zoomPercent"
        :locale="locale"
        :messages="messages"
        @page-count-change="onPageCountChange"
        @visible-page-range="onVisiblePageRange"
        @context-menu="onDocumentContextMenu"
        @search-state-change="find.syncSearchState"
        @request-page-reveal="emit('requestPageReveal', $event)"
        @bookmark-missing="emit('bookmarkMissing', $event)"
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
      @close="contextMenu.hide()"
    />

    <!-- Drag overlay -->
    <DocxDragOverlay
      v-if="!interactionLocked && dragOverlay.visible.value"
      :controller="controller"
    />

    <!-- Node reorder handle and drop indicator -->
    <button
      v-if="reorderHandleVisible"
      type="button"
      class="docx-node-reorder-handle"
      data-testid="docx-node-reorder-handle"
      :data-docx-node-index="reorder.hover.value?.nodeIndex"
      :aria-label="m.docx.nodeDrag.handle"
      :title="m.docx.nodeDrag.handleTitle"
      :style="reorderHandleStyle"
      @mousedown.prevent.stop="onReorderHandlePointerDown"
      @keydown.up.prevent.stop="onReorderHandleKey(-1)"
      @keydown.down.prevent.stop="onReorderHandleKey(1)"
    >
      <span aria-hidden="true">⠿</span>
    </button>
    <div
      v-if="reorder.dragging.value && reorder.dropTop.value !== undefined"
      class="docx-node-drop-indicator"
      data-testid="docx-node-drop-indicator"
      :style="reorderDropIndicatorStyle"
    />

    <!-- Link hover card -->
    <DocxLinkHoverCard
      v-if="linkCard.visible.value && linkCard.state.value"
      :href="linkCard.state.value.href"
      :x="linkCard.state.value.x"
      :y="linkCard.state.value.y"
      :editable="linkCard.state.value.editable"
      :document-theme="controller.documentTheme"
      @open="onLinkCardOpen"
      @edit="onLinkCardEdit"
      @remove="onLinkCardRemove"
      @close="hideLinkCard()"
      @pointerenter="cancelLinkCardHide"
      @pointerleave="scheduleLinkCardHide"
    />

    <!-- Form field properties dialog -->
    <DocxFormFieldDialog
      v-if="formFieldDialogLocation && formFieldDialogField"
      :field="formFieldDialogField"
      :location="formFieldDialogLocation"
      :controller="controller"
      @close="closeFormFieldDialog"
    />

    <!-- Status bar -->
    <OfficeLiveStatus test-id="editor-live-status" :message="liveStatusMessage" />
    <div class="docx-editor-statusbar">
      <span class="docx-editor-statusbar-page">
        {{ pageIndicator }}
      </span>
      <span class="docx-editor-statusbar-status" data-testid="editor-status">
        {{ localizedStatus }}
      </span>
      <span v-if="controller.fileName" class="docx-editor-statusbar-file" data-testid="editor-file-name">
        {{ controller.fileName }}
      </span>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, provide, watch, nextTick, onMounted, onUnmounted } from "vue"
import type { OfficeLocale, OfficeMessagesOverride } from "@arcships/office-interaction"
import { formatOfficeMessage } from "@arcships/office-interaction"
import type {
  DocxEditorController,
  DocxContextMenuContext,
  DocxFormFieldLocation,
  DocxRuntime,
  DocxTextRangeLocation,
  FormFieldRunNode,
} from "@arcships/docx-core"
import { createDocxRuntime, DocxImportError } from "@arcships/docx-core"
import {
  getParagraphAtLocation,
  imageLocationToParagraphLocation,
  resolveDocxImageWrapState,
} from "@arcships/docx-core"
import { OfficeLiveStatus } from "@arcships/vue-ui"
import { OfficeFindBar } from "@arcships/vue-ui"
import { useDocxEditor } from "../composables/useDocxEditor"
import { useDocxFindReplace } from "../composables/useDocxFindReplace"
import {
  docxImageLocationFromNode,
  docxTableContextFromNode,
  docxTextHostFromNode,
  docxTextLocationForNode,
  readDocxTextHostText,
  syncDomSelectionToEditor,
  textOffsetAtDomPoint,
} from "../composables/editor-selection"
import {
  parseDocxFormFieldLocationKey,
  resolveFormFieldAt,
} from "../composables/docx-form-fields"
import { useDocxNodeReorder } from "../composables/useDocxNodeReorder"
import DocxDocumentSurface from "./DocxDocumentSurface.vue"
import DocxToolbar from "./DocxToolbar.vue"
import DocxContextMenu from "./DocxContextMenu.vue"
import DocxThumbnailPanel from "./DocxThumbnailPanel.vue"
import DocxDragOverlay from "./DocxDragOverlay.vue"
import DocxLinkHoverCard from "./DocxLinkHoverCard.vue"
import DocxFormFieldDialog from "./DocxFormFieldDialog.vue"
import { provideDocxMessages } from "../i18n"
import type { DocxBookmarkRevealResult } from "../composables/docx-bookmarks"

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
    /** Optional minimum height; the editor itself does not impose one. */
    minHeight?: string | number
    /** Dark editor shell and document theme. Follows the controller theme when omitted. */
    isDark?: boolean
    locale?: OfficeLocale
    messages?: OfficeMessagesOverride
    /** Author name stored on comments created in this editor. */
    commentAuthor?: string
    /** Comment initials; derived from `commentAuthor` when omitted. */
    commentInitials?: string
    /** Link hover card with open/edit/remove actions. Default `true`. */
    showLinkCard?: boolean
    /** Double-clicking a form field opens its properties dialog. Default `true`. */
    showFormFieldDialog?: boolean
    /** Hovering the paragraph-style picker previews each style. Default `true`. */
    showStylePreview?: boolean
    /** Drag handle that reorders body paragraphs/tables. Default `true`. */
    showNodeReorder?: boolean
  }>(),
  { editable: true, showToolbar: true, showThumbnails: false, minHeight: undefined, isDark: false, locale: "en", messages: undefined, commentAuthor: undefined, commentInitials: undefined, showLinkCard: true, showFormFieldDialog: true, showStylePreview: true, showNodeReorder: true }
)

const emit = defineEmits<{
  (event: "load-error", error: Error): void
  (event: "update:isDark", isDark: boolean): void
  /**
   * A link card action asked to open `href`. External URLs are opened by the
   * editor; internal `#bookmark` targets are host navigation, so the host can
   * scroll to the anchor it resolves.
   */
  (event: "link-open", href: string): void
  /** A bookmark reveal needs `pageIndex` mounted/visible in the host. */
  (event: "requestPageReveal", pageIndex: number): void
  /** An internal link named a bookmark this document does not contain. */
  (event: "bookmarkMissing", name: string): void
}>()

const m = provideDocxMessages(() => props.locale, () => props.messages)

// ── Controller ─────────────────────────────────────────────────────
const internalController = !props.editor
  ? useDocxEditor({
      starterModel: props.model,
      initialFileName: props.file ? m.value.docx.editor.importedDocument : m.value.docx.editor.newDocument,
      statusMessages: () => m.value.docx.editor.status,
      commentAuthor: props.commentAuthor ?? m.value.docx.editor.commentAuthor,
      commentInitials: props.commentInitials,
    })
  : null
const controller = props.editor ?? internalController!
const isDarkTheme = computed(() => props.isDark || controller.documentTheme === "dark")
const minHeightStyle = computed(() => {
  const value = props.minHeight
  if (value === undefined) return undefined
  return { minHeight: typeof value === "number" ? `${value}px` : value }
})
watch(
  () => controller.documentTheme,
  (theme) => { emit("update:isDark", theme === "dark") },
)
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

// ── Find & replace ─────────────────────────────────────────────────
// The surface owns the search session; the composable owns the bar's inputs and
// turns Replace / Replace All into the controller's one-undo-entry commands.
const findBarRef = ref<InstanceType<typeof OfficeFindBar> | undefined>()
const find = useDocxFindReplace({
  controller: () => controller,
  surface: () => viewerRootRef.value,
})
const findState = computed(() => find.state.value)
const findReplaceStatus = computed(() => {
  const feedback = findState.value.feedback
  if (!feedback) return ""
  if (feedback.kind === "unavailable") return m.value.common.search.noResults
  if (feedback.count === 0) return ""
  return formatOfficeMessage(m.value.common.search.replaced, { count: feedback.count })
})

watch(
  () => findState.value.open,
  (open) => { if (open) void nextTick(() => findBarRef.value?.focus()) },
)

// Root element ref keeps context-menu/selection scoping explicit.
const editorRootRef = ref<HTMLElement | undefined>()

// ── Page tracking ──────────────────────────────────────────────────
const currentPage = ref(1)
const totalPages = ref(1)
const pageIndicator = computed(() => formatOfficeMessage(m.value.common.page.indicator, {
  current: currentPage.value,
  total: totalPages.value,
}))
const zoomPercent = ref(100)
const thumbnailPanelVisible = ref(props.showThumbnails)

watch(() => props.showThumbnails, (visible) => { thumbnailPanelVisible.value = visible })

// The editor status bar is visible but was never announced. Mirror it in a
// live region so page turns and controller state reach assistive technology.
const localizedStatus = computed(() => String(controller.status ?? "").trim())
const liveStatusMessage = computed(() => {
  if (pageLimitError.value) return formatOfficeMessage(m.value.docx.editor.liveError, { message: pageLimitError.value.message })
  const pages = totalPages.value > 0
    ? formatOfficeMessage(m.value.common.page.indicator, {
      current: Math.min(Math.max(currentPage.value, 1), totalPages.value),
      total: totalPages.value,
    })
    : ""
  const status = localizedStatus.value
  return [pages, status].filter(Boolean).join("，") || m.value.docx.editor.liveReady
})

function onPageCountChange(count: number): void {
  const allowed = effectiveRuntime.value.limits?.maxDocxPages
  if (allowed !== undefined && count > allowed) {
    if (!pageLimitError.value) {
      const normalized = new DocxImportError(
        "LIMIT_EXCEEDED",
        formatOfficeMessage(m.value.docx.viewer.pageLimitExceeded, { count, allowed }),
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
  const imageLocation = docxImageLocationFromNode(target)
  const textHost = docxTextHostFromNode(target)
  const range = controller.activeTextRange

  let context: DocxContextMenuContext = tableContext
    ? { kind: "table", tableContext }
    : { kind: "text" }
  if (!tableContext && imageLocation) {
    const { paragraph } = getParagraphAtLocation(
      controller.model,
      imageLocationToParagraphLocation(imageLocation),
    )
    const image = paragraph?.children[imageLocation.childIndex]
    if (image?.type === "image") {
      context = {
        kind: "image",
        image: {
          location: imageLocation,
          floating: image.floating,
          wrap: resolveDocxImageWrapState(image.floating),
        },
      }
    }
  }
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

// ── Link hover card ────────────────────────────────────────────────
// Hovering a rendered link shows open/edit/remove without changing the
// selection. Editing and removing select the linked characters first, so the
// controller's `setLink()` rewrites exactly that link.
interface LinkCardState {
  href: string
  x: number
  y: number
  editable: boolean
  location?: DocxTextRangeLocation
  start?: number
  end?: number
}

const linkCard = {
  visible: ref(false),
  state: ref<LinkCardState | undefined>(undefined),
}
let linkCardHideTimer: ReturnType<typeof setTimeout> | undefined

function cancelLinkCardHide(): void {
  if (linkCardHideTimer !== undefined) clearTimeout(linkCardHideTimer)
  linkCardHideTimer = undefined
}

function hideLinkCard(): void {
  cancelLinkCardHide()
  linkCard.visible.value = false
  linkCard.state.value = undefined
}

function scheduleLinkCardHide(): void {
  cancelLinkCardHide()
  linkCardHideTimer = setTimeout(() => {
    linkCard.visible.value = false
    linkCard.state.value = undefined
    linkCardHideTimer = undefined
  }, 140)
}

function linkElementFromEvent(event: Event): Element | undefined {
  const target = event.target as (Element & { closest?: (selector: string) => Element | null }) | null
  if (!target?.closest) return undefined
  const anchor = target.closest("a[href]")
  if (!anchor) return undefined
  // A link inside the hover card itself is card chrome, not a document link.
  if (anchor.closest?.("[data-docx-link-card]")) return undefined
  return anchor
}

function onEditorPointerOver(event: PointerEvent | Event): void {
  if (props.showLinkCard === false || !props.editable || interactionLocked.value) return
  const anchor = linkElementFromEvent(event)
  if (!anchor) return
  const href = anchor.getAttribute?.("href")
  if (!href) return

  const start = Number(anchor.getAttribute?.("data-docx-link-start"))
  const end = Number(anchor.getAttribute?.("data-docx-link-end"))
  const hasRange = Number.isFinite(start) && Number.isFinite(end) && end > start
  const location = docxTextLocationForNode(anchor)
  const rect = (anchor as Element & {
    getBoundingClientRect?: () => { left: number; bottom: number }
  }).getBoundingClientRect?.()

  cancelLinkCardHide()
  linkCard.state.value = {
    href,
    x: rect ? Math.round(rect.left) : 0,
    y: rect ? Math.round(rect.bottom + 6) : 0,
    editable: hasRange && Boolean(location),
    location: hasRange ? location : undefined,
    start: hasRange ? start : undefined,
    end: hasRange ? end : undefined,
  }
  linkCard.visible.value = true
}

function onEditorPointerOut(event: PointerEvent | Event): void {
  if (!linkCard.visible.value) return
  if (!linkElementFromEvent(event)) return
  scheduleLinkCardHide()
}

function selectHoveredLink(): boolean {
  const state = linkCard.state.value
  if (!state?.editable || !state.location) return false
  if (state.start === undefined || state.end === undefined) return false
  controller.setActiveTextRange({
    start: { location: state.location, offset: state.start },
    end: { location: state.location, offset: state.end },
  })
  return true
}

function onLinkCardOpen(href: string): void {
  hideLinkCard()
  // Internal bookmarks are document navigation: reveal the target locally and
  // still expose the action to the host for its own routing.
  if (href.startsWith("#")) {
    void viewerRootRef.value?.revealBookmark(href)
    emit("link-open", href)
    return
  }
  const target = globalThis as { window?: { open?: (url: string, name?: string, features?: string) => unknown } }
  target.window?.open?.(href, "_blank", "noopener,noreferrer")
  emit("link-open", href)
}

function onLinkCardEdit(href: string): void {
  if (!selectHoveredLink()) {
    hideLinkCard()
    return
  }
  controller.setLink(href)
  hideLinkCard()
}

function onLinkCardRemove(): void {
  if (!selectHoveredLink()) {
    hideLinkCard()
    return
  }
  controller.setLink(undefined)
  hideLinkCard()
}

// ── Form field properties dialog ───────────────────────────────────
const formFieldDialogLocation = ref<DocxFormFieldLocation | undefined>()
const formFieldDialogField = computed<FormFieldRunNode | undefined>(() =>
  resolveFormFieldAt(controller.model, formFieldDialogLocation.value)
)

function onEditorDoubleClick(event: MouseEvent | Event): void {
  if (props.showFormFieldDialog === false || !props.editable || interactionLocked.value) return
  const target = event.target as (Element & { closest?: (selector: string) => Element | null }) | null
  const fieldElement = target?.closest?.("[data-docx-form-field-location]") as Element | null | undefined
  const location = parseDocxFormFieldLocationKey(
    fieldElement?.getAttribute?.("data-docx-form-field-location")
  )
  if (!location) return
  event.preventDefault?.()
  controller.selectFormField(location)
  formFieldDialogLocation.value = location
}

function closeFormFieldDialog(): void {
  const location = formFieldDialogLocation.value
  formFieldDialogLocation.value = undefined
  if (location) controller.endFormFieldEdit(location)
}

// ── Drag-to-reorder body nodes ─────────────────────────────────────
const reorder = useDocxNodeReorder(controller, {
  enabled: () => props.editable && !interactionLocked.value && props.showNodeReorder !== false,
})

const reorderHandleVisible = computed(
  () => reorder.enabled.value && !reorder.dragging.value && reorder.hover.value !== undefined
)

const reorderHandleStyle = computed(() => {
  const hover = reorder.hover.value
  if (!hover) return undefined
  return {
    position: "fixed" as const,
    left: `${Math.max(0, Math.round(hover.left - 22))}px`,
    top: `${Math.round(hover.top + 2)}px`,
    zIndex: 40,
  }
})

const reorderDropIndicatorStyle = computed(() => {
  const hover = reorder.hover.value
  return {
    position: "fixed" as const,
    left: `${Math.max(0, Math.round((hover?.left ?? 0) - 8))}px`,
    width: `${Math.max(24, Math.round((hover?.width ?? 280) + 8))}px`,
    top: `${Math.round(reorder.dropTop.value ?? 0)}px`,
    zIndex: 45,
  }
})

let reorderMouseMove: ((event: MouseEvent) => void) | undefined
let reorderMouseUp: (() => void) | undefined
let reorderKeyDown: ((event: KeyboardEvent) => void) | undefined

function detachReorderListeners(): void {
  const doc = (globalThis as { document?: Document }).document
  if (reorderMouseMove) doc?.removeEventListener("mousemove", reorderMouseMove)
  if (reorderMouseUp) doc?.removeEventListener("mouseup", reorderMouseUp)
  if (reorderKeyDown) doc?.removeEventListener("keydown", reorderKeyDown)
  reorderMouseMove = undefined
  reorderMouseUp = undefined
  reorderKeyDown = undefined
}

function finishReorderDrag(): void {
  reorder.finishDrag()
  detachReorderListeners()
}

function onReorderHandlePointerDown(event: PointerEvent | MouseEvent): void {
  const nodeIndex = reorder.hover.value?.nodeIndex
  if (nodeIndex === undefined || !reorder.enabled.value) return
  reorder.startDrag(nodeIndex)
  const doc = (globalThis as { document?: Document }).document
  if (!doc?.addEventListener) return
  detachReorderListeners()
  reorderMouseMove = (moveEvent: MouseEvent) => reorder.updateDrag(moveEvent)
  reorderMouseUp = () => finishReorderDrag()
  reorderKeyDown = (keyEvent: KeyboardEvent) => {
    if (keyEvent.key !== "Escape") return
    reorder.cancelDrag()
    detachReorderListeners()
  }
  doc.addEventListener("mousemove", reorderMouseMove)
  doc.addEventListener("mouseup", reorderMouseUp)
  doc.addEventListener("keydown", reorderKeyDown)
  event.preventDefault?.()
}

function onReorderHandleKey(delta: -1 | 1): void {
  const nodeIndex = reorder.hover.value?.nodeIndex
  if (nodeIndex === undefined) return
  if (reorder.moveBy(nodeIndex, delta)) reorder.clearHover()
}

function onEditorPointerMove(event: PointerEvent | MouseEvent): void {
  if (reorder.dragging.value) {
    reorder.updateDrag(event)
    return
  }
  const target = event.target as (Element & { closest?: (selector: string) => Element | null }) | null
  // The handle sits beside the node, outside the paragraph host: hovering it
  // must keep the handle (and therefore the node it moves) alive.
  if (target?.closest?.(".docx-node-reorder-handle")) return
  reorder.onHover(event)
}

function onEditorPointerLeave(): void {
  reorder.clearHover()
  scheduleLinkCardHide()
}

// ── Keyboard shortcuts ─────────────────────────────────────────────
// The listener lives on the editor root, so host-page inputs and other editor
// instances never see these shortcuts. Commands are also suppressed while the
// document is read-only or while tracked-changes review locks editing.
function onKeydown(event: KeyboardEvent): void {
  if (event.defaultPrevented || event.isComposing) return

  const target = event.target as HTMLElement | null
  const tagName = target?.tagName?.toLowerCase()
  const inField = tagName === "input" || tagName === "textarea" || tagName === "select"

  // Ctrl/Cmd + F opens the find bar even in a read-only editor: searching is
  // always allowed, only the replace half is disabled.
  const mod = event.metaKey || event.ctrlKey
  const key = event.key.toLowerCase()
  if (mod && key === "f" && props.showToolbar && !inField) {
    event.preventDefault()
    find.openFind()
    return
  }
  if (event.key === "Escape" && findState.value.open) {
    event.preventDefault()
    find.closeFind()
    return
  }

  if (!props.editable) return
  if (inField) return
  if (target?.closest?.(".docx-toolbar-shell")) return

  if (!mod) return

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
    hideLinkCard()
    cancelLinkCardHide()
    formFieldDialogLocation.value = undefined
    reorder.cancelDrag()
    detachReorderListeners()
    controller.clearSelectionSession()
  },
)

onMounted(() => {
  syncBeforeUnload()
  // The surface exists now, so a query typed before mount (or a model restored
  // from a cache) can resolve its results.
  void find.refresh()
})

onUnmounted(() => {
  if (beforeUnloadAttached) {
    window.removeEventListener("beforeunload", onBeforeUnload)
    beforeUnloadAttached = false
  }
  cancelLinkCardHide()
  detachReorderListeners()
  ownedLimitRuntime.dispose()
})

defineExpose({
  /** Reveals an internal `#bookmark` target in this document. */
  revealBookmark(name: string): Promise<DocxBookmarkRevealResult> {
    return viewerRootRef.value?.revealBookmark(name) ??
      Promise.resolve({ status: "missing", name: name.replace(/^#/, "") })
  },
})
</script>

<style scoped>
.docx-editor-viewer {
  --docx-editor-bg: #f3f4f6;
  --docx-editor-panel: #f9fafb;
  --docx-editor-border: #e5e7eb;
  --docx-editor-fg: #374151;
  --docx-editor-muted: #6b7280;
  container: docx-shell / inline-size;
  display: flex;
  flex-direction: column;
  height: 100%;
  overflow: hidden;
  background: var(--docx-editor-bg);
  color: var(--docx-editor-fg);
}
.docx-editor-viewer--dark {
  --docx-editor-bg: #09090b;
  --docx-editor-panel: #18181b;
  --docx-editor-border: #3f3f46;
  --docx-editor-fg: #f4f4f5;
  --docx-editor-muted: #a1a1aa;
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
  background: var(--docx-editor-panel);
  border-top: 1px solid var(--docx-editor-border);
  font-size: 12px;
  color: var(--docx-editor-muted);
  flex-shrink: 0;
}
.docx-editor-statusbar-page {
  font-weight: 500;
}
.docx-editor-statusbar-status {
  color: var(--docx-editor-muted);
}
.docx-editor-statusbar-file {
  margin-left: auto;
  color: var(--docx-editor-fg);
  font-weight: 500;
}

.docx-node-reorder-handle {
  align-items: center;
  background: var(--docx-editor-panel);
  border: 1px solid var(--docx-editor-border);
  border-radius: 4px;
  color: var(--docx-editor-muted);
  cursor: grab;
  display: inline-flex;
  font-size: 12px;
  height: 20px;
  justify-content: center;
  line-height: 1;
  padding: 0;
  width: 16px;
}
.docx-node-reorder-handle:hover,
.docx-node-reorder-handle:focus-visible {
  border-color: #3b82f6;
  color: #2563eb;
  outline: none;
}
.docx-node-drop-indicator {
  background: #3b82f6;
  border-radius: 2px;
  height: 2px;
  pointer-events: none;
}

@media print {
  .docx-editor-viewer {
    background: #fff !important;
    color: #000 !important;
    height: auto !important;
    min-height: 0 !important;
    overflow: visible !important;
  }
  .docx-toolbar-shell,
  .docx-thumbnail-panel,
  .docx-editor-statusbar,
  .docx-context-menu-backdrop,
  .docx-context-menu,
  .docx-drag-overlay,
  .docx-editor-limit-error { display: none !important; }
  .docx-editor-body { display: block !important; overflow: visible !important; }
}
</style>
