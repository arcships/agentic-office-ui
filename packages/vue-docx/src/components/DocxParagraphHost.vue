<template>
  <div
    ref="paragraphEl"
    class="docx-paragraph-host docx-viewer-paragraph"
    :class="{
      'docx-paragraph-host--editable': editable,
      'docx-paragraph-host--active': isActiveParagraph,
    }"
    :contenteditable="editable ? 'true' : undefined"
    :data-docx-paragraph-host="true"
    :data-docx-paragraph-node-index="paragraphIndex"
    data-testid="editor-paragraph"
    :data-node-index="paragraphIndex"
    :suppresscontenteditablewarning="editable ? 'true' : undefined"
    :style="paragraphStyle"
    :spellcheck="editable ? 'true' : undefined"
    @input="onInput"
    @pointerdown="onPointerDown"
    @mouseup="onMouseUp"
    @keydown="onKeydown"
    @keyup="onKeyup"
    @click="onClick"
    @focus="onFocus"
    @blur="onBlur"
    @compositionstart="onCompositionStart"
    @compositionend="onCompositionEnd"
    @paste="onPaste"
    v-html="displayHtml"
  />
</template>

<script setup lang="ts">
import { ref, computed, watch, nextTick, onUnmounted } from "vue"
import type {
  DocxEditorController,
  DocxTextRange,
  NumberingDefinitionSet,
  ParagraphNode,
  DocxDocumentTheme,
} from "@arcships/docx-core"
import {
  compareTextRangeBoundaries,
  paragraphText,
  paragraphIsList,
} from "@arcships/docx-core"
import { renderParagraphRuns, type ParagraphRunRenderOptions } from "../render/paragraph-runs"
import { renderStaticHtml } from "../render/static-html"
import {
  restoreDomSelectionFromTextRange,
  syncDomSelectionToEditor,
} from "../composables/editor-selection"

// ── Constants ──────────────────────────────────────────────────────
const HEADING_FONT_SIZES: Record<number, string> = {
  1: "2rem", 2: "1.6rem", 3: "1.35rem",
  4: "1.2rem", 5: "1.05rem", 6: "0.95rem",
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    paragraph: ParagraphNode
    paragraphIndex: number
    editable?: boolean
    documentTheme?: DocxDocumentTheme
    controller?: DocxEditorController
    numberingLabel?: string
    numberingDefinitions?: NumberingDefinitionSet
    showTrackedChanges?: boolean
    showCommentHighlights?: boolean
    pageNumber?: number
    totalPages?: number
    pageNumberFormat?: string
    withinHeaderFooter?: boolean
    headerFooterRegion?: "header" | "footer"
  }>(),
  {
    editable: false,
    documentTheme: "light" as DocxDocumentTheme,
  }
)

const emit = defineEmits<{
  textInput: [nodeIndex: number, text: string]
  focus: [nodeIndex: number]
  blur: [nodeIndex: number]
  click: [nodeIndex: number, event: MouseEvent]
}>()

// ── State ──────────────────────────────────────────────────────────
const paragraphEl = ref<HTMLElement | undefined>()
const draftText = ref<string | null>(null)
const draftDirty = ref(false)
const isComposing = ref(false)

// ── Computed ───────────────────────────────────────────────────────
const isActiveParagraph = computed(() => {
  if (!props.editable || !props.controller) return false
  const sel = props.controller.selection
  if (sel.kind === "paragraph") {
    return sel.nodeIndex === props.paragraphIndex
  }
  return false
})

const paragraphStyle = computed<Record<string, any>>(() => {
  const style = props.paragraph.style
  return {
    margin: "0",
    minHeight: "1em",
    textAlign: (style?.align as any) ?? undefined,
    fontWeight: style?.headingLevel ? 700 : undefined,
    fontSize: style?.headingLevel
      ? HEADING_FONT_SIZES[style.headingLevel]
      : undefined,
  }
})

const displayHtml = computed(() => {
  const theme = props.documentTheme
  const runOptions: ParagraphRunRenderOptions = {
    showTrackedChanges: props.showTrackedChanges,
    showCommentHighlights: props.showCommentHighlights,
    numberingDefinitions: props.numberingDefinitions,
    withinHeaderFooter: props.withinHeaderFooter,
    headerFooterRegion: props.headerFooterRegion,
    pageNumberFormat: props.pageNumberFormat,
    skipFloatingImages: props.editable,
    renderWrappedFloatingImages: props.editable,
  }

  const vnodes = renderParagraphRuns(
    props.paragraph,
    `p-${props.paragraphIndex}`,
    theme,
    props.numberingLabel ? { label: props.numberingLabel } as any : undefined,
    undefined,
    undefined,
    undefined,
    props.pageNumber,
    props.totalPages,
    runOptions
  )

  const result = Array.isArray(vnodes) ? vnodes : [vnodes]
  return result.map((v) => renderStaticHtml(v)).join("")
})

// ── Draft management ───────────────────────────────────────────────
function readCurrentDomText(): string {
  const root = paragraphEl.value
  if (!root) return ""

  function readNode(node: Node): string {
    const element = node as Element
    if (element.matches?.("[data-docx-numbering-label='true']")) return ""
    if (element.matches?.("[data-docx-image-child-index]")) return ""
    if (element.tagName?.toLowerCase() === "br") return "\n"
    if (node.nodeType === 3) return node.textContent ?? ""
    return Array.from(node.childNodes).map(readNode).join("")
  }

  return readNode(root)
}

function syncDraftFromDom(): void {
  if (!props.editable || !paragraphEl.value) return
  const domText = readCurrentDomText()
  draftText.value = domText
  draftDirty.value = domText !== paragraphText(props.paragraph)
}

function clearDraft(): void {
  draftText.value = null
  draftDirty.value = false
}

function commitDraft(): void {
  if (draftText.value === null || !draftDirty.value) {
    clearDraft()
    return
  }
  const text = draftText.value
  clearDraft()
  emit("textInput", props.paragraphIndex, text)
}

// ── Event handlers ─────────────────────────────────────────────────
function selectionRoot(): HTMLElement | undefined {
  const element = paragraphEl.value
  return (element?.closest?.(".docx-editor-viewer") as HTMLElement | undefined) ?? element
}

function syncCurrentSelection(): void {
  const root = selectionRoot()
  if (!root || !props.controller) return
  syncDomSelectionToEditor(props.controller, root)
}

function currentSelectionOffsets(textLength: number): { start: number; end: number } {
  syncCurrentSelection()
  const range = props.controller?.activeTextRange
  const startsHere = range?.start.location.kind === "paragraph" &&
    range.start.location.nodeIndex === props.paragraphIndex
  const endsHere = range?.end.location.kind === "paragraph" &&
    range.end.location.nodeIndex === props.paragraphIndex
  if (range && startsHere && endsHere) {
    return {
      start: Math.max(0, Math.min(range.start.offset, textLength)),
      end: Math.max(0, Math.min(range.end.offset, textLength)),
    }
  }
  return { start: textLength, end: textLength }
}

function restoreControllerSelection(focus = true): void {
  const controller = props.controller
  const root = selectionRoot()
  const range = controller?.activeTextRange
  if (!controller || !root || !range) return
  void nextTick(() => {
    restoreDomSelectionFromTextRange(root, range, { focus })
  })
}

function localTextRange(start: number, end: number): DocxTextRange {
  const location = { kind: "paragraph" as const, nodeIndex: props.paragraphIndex }
  return {
    start: { location, offset: start },
    end: { location, offset: end },
  }
}

function deleteExpandedRange(): boolean {
  if (!props.controller) return false
  const currentText = readCurrentDomText()
  const offsets = currentSelectionOffsets(currentText.length)
  let range: DocxTextRange | undefined
  if (offsets.start < offsets.end) {
    range = localTextRange(offsets.start, offsets.end)
  } else {
    syncCurrentSelection()
    const activeRange = props.controller.activeTextRange
    if (activeRange && compareTextRangeBoundaries(activeRange.start, activeRange.end) < 0) {
      range = activeRange
    }
  }
  if (!range) return false

  syncDraftFromDom()
  commitDraft()
  const nextRange = props.controller.deleteExpandedSelection(range)
  if (nextRange) restoreControllerSelection()
  return Boolean(nextRange)
}

function deleteAdjacentParagraph(direction: "backward" | "forward"): boolean {
  const controller = props.controller
  if (!controller) return false
  const currentText = readCurrentDomText()
  const offsets = currentSelectionOffsets(currentText.length)
  const atBoundary = direction === "backward"
    ? offsets.start === 0 && offsets.end === 0
    : offsets.start === currentText.length && offsets.end === currentText.length
  if (!atBoundary) return false

  const adjacentIndex = props.paragraphIndex + (direction === "backward" ? -1 : 1)
  const adjacentNode = controller.model.nodes[adjacentIndex]
  if (!adjacentNode || adjacentNode.type !== "paragraph") return false

  syncDraftFromDom()
  commitDraft()
  const range = direction === "backward"
    ? {
        start: {
          location: { kind: "paragraph" as const, nodeIndex: adjacentIndex },
          offset: paragraphText(adjacentNode).length,
        },
        end: {
          location: { kind: "paragraph" as const, nodeIndex: props.paragraphIndex },
          offset: 0,
        },
      }
    : {
        start: {
          location: { kind: "paragraph" as const, nodeIndex: props.paragraphIndex },
          offset: currentText.length,
        },
        end: {
          location: { kind: "paragraph" as const, nodeIndex: adjacentIndex },
          offset: 0,
        },
      }
  const nextRange = controller.deleteExpandedSelection(range)
  if (nextRange) restoreControllerSelection()
  return Boolean(nextRange)
}

function onInput(_event: Event): void {
  if (!props.editable || isComposing.value) return
  syncDraftFromDom()
  syncCurrentSelection()
}

function onPointerDown(): void {
  if (!props.editable || !props.controller) return
  props.controller.beginSelectionSession("pointer", { settleAfterMs: 250 })
}

function onMouseUp(): void {
  if (!props.editable || !props.controller) return
  syncCurrentSelection()
  props.controller.clearSelectionSession("pointer")
}

function onKeydown(event: KeyboardEvent): void {
  if (!props.editable || !props.controller) return
  if (isComposing.value || event.isComposing) return
  props.controller.beginSelectionSession("keyboard", { settleAfterMs: 120 })

  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault()
    const currentText = readCurrentDomText()
    const offsets = currentSelectionOffsets(currentText.length)
    const target = { kind: "paragraph" as const, nodeIndex: props.paragraphIndex }
    const result = paragraphIsList(props.paragraph, currentText)
      ? props.controller.insertListItemAfterSelection(
          currentText,
          offsets.start,
          offsets.end,
          target
        )
      : props.controller.splitParagraphAtSelection(
          currentText,
          offsets.start,
          offsets.end,
          target
        )
    if (result) {
      clearDraft()
      restoreControllerSelection()
    }
    return
  }

  if (event.key === "Backspace" || event.key === "Delete") {
    if (deleteExpandedRange()) {
      event.preventDefault()
      return
    }
    const direction = event.key === "Backspace" ? "backward" : "forward"
    if (deleteAdjacentParagraph(direction)) {
      event.preventDefault()
    }
  }
}

function onKeyup(_event: KeyboardEvent): void {
  if (!props.editable) return
  syncDraftFromDom()
  syncCurrentSelection()
}

function onClick(event: MouseEvent): void {
  syncCurrentSelection()
  emit("click", props.paragraphIndex, event)
}

function onFocus(): void {
  if (!props.editable) return
  clearDraft()
  syncCurrentSelection()
  emit("focus", props.paragraphIndex)
}

function onBlur(): void {
  if (!props.editable) return
  syncDraftFromDom()
  commitDraft()
  emit("blur", props.paragraphIndex)
}

function onCompositionStart(): void {
  isComposing.value = true
  props.controller?.beginSelectionSession("composition")
}

function onCompositionEnd(_event: Event): void {
  isComposing.value = false
  syncDraftFromDom()
  syncCurrentSelection()
  props.controller?.clearSelectionSession("composition")
}

function onPaste(event: ClipboardEvent): void {
  if (!props.editable) return
  event.preventDefault()
  const text = event.clipboardData?.getData("text/plain") ?? ""
  if (text) {
    // Insert plain text at cursor
    document.execCommand("insertText", false, text)
    syncDraftFromDom()
    syncCurrentSelection()
  }
}

// ── Watch for model changes (clears draft when paragraph changes externally) ─
watch(
  () => paragraphText(props.paragraph),
  () => {
    if (!draftDirty.value) clearDraft()
  }
)

onUnmounted(() => {
  if (isComposing.value) props.controller?.clearSelectionSession("composition")
})

// ── Expose ─────────────────────────────────────────────────────────
defineExpose({
  get element() { return paragraphEl.value },
  focus() { paragraphEl.value?.focus() },
  restoreDraft(text: string) {
    draftText.value = text
    draftDirty.value = text !== paragraphText(props.paragraph)
  },
})
</script>

<style scoped>
.docx-paragraph-host {
  margin: 0;
  min-height: 1em;
  outline: none;
  white-space: pre-wrap;
  word-wrap: break-word;
  line-height: 1.35;
  font-family: "Calibri", sans-serif;
  font-size: 11pt;
}
.docx-paragraph-host--editable {
  cursor: text;
}
.docx-paragraph-host--editable:empty::before {
  content: "Type here...";
  color: #9ca3af;
}
.docx-paragraph-host--active {
  background: rgba(59, 130, 246, 0.04);
}
</style>
