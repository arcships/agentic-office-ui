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
import { ref, computed, watch, nextTick, onBeforeUnmount, onMounted, onUnmounted } from "vue"
import type {
  DocxFloatingPageOriginPx,
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
  paragraphTypographyStyle,
} from "@arcships/docx-core"
import { renderParagraphRuns, type ParagraphRunRenderOptions } from "../render/paragraph-runs"
import { paragraphStyleToCss } from "../render/paragraph-style"
import { renderStaticHtml } from "../render/static-html"
import {
  restoreDomSelectionFromTextRange,
  syncDomSelectionToEditor,
  readDocxTextHostText,
} from "../composables/editor-selection"

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
    sectionPages?: number
    resolveStyleRefFieldValue?: (
      target: string,
      beforeNodeIndex?: number,
    ) => string | undefined
    resolveBookmarkFieldValue?: (target: string) => string | undefined
    resolveBookmarkPageNumber?: (target: string) => number | undefined
    fieldResolutionNodeIndex?: number
    withinHeaderFooter?: boolean
    headerFooterRegion?: "header" | "footer"
    noteMarkerLabels?: { footnote?: Map<number, string>; endnote?: Map<number, string> }
    skipImageChildIndexes?: readonly number[]
    /** Page-relative origin used to place margin/page/column anchored images. */
    floatingPageOriginPx?: DocxFloatingPageOriginPx
    /** Space to insert above the paragraph so it clears an overflowing float. */
    wrapClearanceTopPx?: number
    /**
     * Render only this wrap-line slice of the paragraph (pagination/columns).
     * The host clips overflow and offsets the content so the slice starts at
     * its own top edge. Ignored while editing.
     */
    lineRange?: {
      startLineIndex: number
      endLineIndex: number
      totalLineCount: number
      lineHeightPx: number
    }
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
  measure: [nodeIndex: number, heightPx: number]
}>()

// ── State ──────────────────────────────────────────────────────────
const paragraphEl = ref<HTMLElement | undefined>()
const draftText = ref<string | null>(null)
const draftDirty = ref(false)
const isComposing = ref(false)
let draftBaseText: string | undefined
let draftDocumentNonce: number | undefined
let unregisterDraftFlusher: (() => void) | undefined

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
  const resolved: Record<string, any> = {
    ...paragraphStyleToCss(props.paragraph, {
      numberingDefinitions: props.numberingDefinitions,
    }),
    // Keep the shared paragraph typography (alignment/direction) too.
    ...paragraphTypographyStyle(props.paragraph.style),
  }
  // A floating image that outlives this paragraph is cleared by pushing the next
  // paragraph below it; the paragraph's own spacing-before still applies.
  const wrapClearanceTopPx = Math.max(0, Math.round(props.wrapClearanceTopPx ?? 0))
  if (wrapClearanceTopPx > 0) {
    const modelMarginTopPx = Number.parseFloat(String(resolved.marginTop ?? ""))
    const baseMarginTopPx = Number.isFinite(modelMarginTopPx) ? modelMarginTopPx : 0
    resolved.marginTop = `${baseMarginTopPx + wrapClearanceTopPx}px`
  }

  const lineRange = props.lineRange
  if (lineRange && !props.editable) {
    const lineHeightPx = Math.max(1, Math.round(lineRange.lineHeightPx))
    const startLineIndex = Math.max(0, Math.round(lineRange.startLineIndex))
    const endLineIndex = Math.max(
      startLineIndex + 1,
      Math.round(lineRange.endLineIndex)
    )
    const sliceLineCount = endLineIndex - startLineIndex
    const marginTopPx = Number.parseFloat(String(resolved.marginTop ?? ""))
    const marginBottomPx = Number.parseFloat(String(resolved.marginBottom ?? ""))
    const keepsTopSpacing = startLineIndex === 0
    const keepsBottomSpacing = endLineIndex >= lineRange.totalLineCount
    const topSpacingPx = keepsTopSpacing && Number.isFinite(marginTopPx)
      ? marginTopPx
      : 0
    const bottomSpacingPx = keepsBottomSpacing && Number.isFinite(marginBottomPx)
      ? marginBottomPx
      : 0
    // Pull the paragraph up so the first line of the slice lands at the top of
    // the slot, then clip everything after the slice.
    resolved.marginTop = `${topSpacingPx - startLineIndex * lineHeightPx}px`
    resolved.marginBottom = `${bottomSpacingPx}px`
    resolved.height = `${Math.max(
      lineHeightPx,
      sliceLineCount * lineHeightPx + topSpacingPx + bottomSpacingPx
    )}px`
    resolved.overflow = "hidden"
  }

  return resolved
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
    sectionPages: props.sectionPages,
    resolveStyleRefFieldValue: props.resolveStyleRefFieldValue
      ? (target: string) =>
          props.resolveStyleRefFieldValue?.(
            target,
            props.fieldResolutionNodeIndex,
          )
      : undefined,
    resolveBookmarkFieldValue: props.resolveBookmarkFieldValue,
    resolveBookmarkPageNumber: props.resolveBookmarkPageNumber,
    skipFloatingImages: props.editable,
    renderWrappedFloatingImages: props.editable,
    noteMarkerLabels: props.noteMarkerLabels,
    skipImageChildIndexes: props.skipImageChildIndexes,
  }

  const vnodes = renderParagraphRuns(
    props.paragraph,
    `p-${props.paragraphIndex}`,
    theme,
    props.numberingLabel ? { label: props.numberingLabel } as any : undefined,
    undefined,
    props.floatingPageOriginPx,
    undefined,
    props.pageNumber,
    props.totalPages,
    runOptions
  )

  const result = Array.isArray(vnodes) ? vnodes : [vnodes]
  return result.map((v) => renderStaticHtml(v)).join("")
})

function readCurrentDomText(): string {
  const root = paragraphEl.value
  if (!root) return ""
  // Test renderers (and some host shims) expose only textContent.
  if (!root.childNodes) return root.textContent ?? ""
  // The shared reader skips numbering labels, note markers and tracked-change
  // deletion/move-from runs, which are rendered but are not paragraph text.
  return readDocxTextHostText(root)
}

function syncDraftFromDom(): void {
  if (!props.editable || !paragraphEl.value) return
  const domText = readCurrentDomText()
  if (draftText.value === null) {
    draftBaseText = paragraphText(props.paragraph)
    draftDocumentNonce = props.controller?.documentLoadNonce
  }
  draftText.value = domText
  draftDirty.value = domText !== paragraphText(props.paragraph)
}

function clearDraft(): void {
  draftText.value = null
  draftDirty.value = false
  draftBaseText = undefined
  draftDocumentNonce = undefined
}

function commitDraft(requireUnchangedBase = false): void {
  if (draftText.value === null || !draftDirty.value) {
    clearDraft()
    return
  }
  if (requireUnchangedBase) {
    const controller = props.controller
    const currentParagraph = controller?.model.nodes[props.paragraphIndex]
    if (
      !controller ||
      draftDocumentNonce !== controller.documentLoadNonce ||
      !currentParagraph ||
      currentParagraph.type !== "paragraph" ||
      paragraphText(currentParagraph) !== (draftBaseText ?? paragraphText(props.paragraph))
    ) {
      // The document changed under the draft (import/new document/external
      // edit). Dropping the stale draft is safer than writing it elsewhere.
      clearDraft()
      return
    }
  }
  const text = draftText.value
  clearDraft()
  emit("textInput", props.paragraphIndex, text)
}

function flushDraft(): void {
  if (!props.controller) return
  if (props.editable) syncDraftFromDom()
  commitDraft(true)
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

  if (event.key === "Tab" && !event.ctrlKey && !event.metaKey && !event.altKey) {
    const currentText = readCurrentDomText()
    if (!paragraphIsList(props.paragraph, currentText)) return
    event.preventDefault()
    syncDraftFromDom()
    commitDraft()
    const changed = props.controller.adjustSelectedListDepth(
      event.shiftKey ? -1 : 1
    )
    if (changed) restoreControllerSelection()
    return
  }

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
  commitDraft(true)
  // Blur closes the typing burst: the next edit is a new undo step.
  props.controller?.endTextCommitBurst?.()
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

function clipboardHtmlToText(html: string): string {
  if (!html) return ""
  if (typeof DOMParser !== "undefined") {
    const parsed = new DOMParser().parseFromString(html, "text/html")
    parsed.querySelectorAll("script,style").forEach((node) => node.remove())
    parsed.querySelectorAll("br").forEach((node) => node.replaceWith("\n"))
    parsed.querySelectorAll("p,div,li,tr,h1,h2,h3,h4,h5,h6,blockquote").forEach((node) => {
      node.append("\n")
    })
    return parsed.body?.textContent ?? ""
  }
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li|tr|h[1-6]|blockquote)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
}

function onPaste(event: ClipboardEvent): void {
  if (!props.editable || !props.controller) return
  const clipboard = event.clipboardData
  const html = clipboard?.getData("text/html") ?? ""
  const plain = clipboard?.getData("text/plain") ?? ""
  const text = html ? clipboardHtmlToText(html) || plain : plain
  if (!text) {
    // Leave file/image clipboards to the browser instead of swallowing them.
    return
  }

  event.preventDefault()
  const currentText = readCurrentDomText()
  const offsets = currentSelectionOffsets(currentText.length)
  const controller = props.controller

  // Use the model range for cross-paragraph selections; the local paragraph
  // offsets can only describe a selection contained in this host.
  const activeRange = controller.activeTextRange
  if (activeRange && compareTextRangeBoundaries(activeRange.start, activeRange.end) < 0) {
    syncDraftFromDom()
    commitDraft(true)
    const nextRange = controller.replaceExpandedSelection(text, activeRange)
    if (nextRange) restoreControllerSelection()
    return
  }

  const caretOffset = offsets.end + text.length
  const nextText = `${currentText.slice(0, offsets.end)}${text}${currentText.slice(offsets.end)}`
  if (nextText === currentText) return
  clearDraft()
  controller.commitParagraphText(props.paragraphIndex, nextText)
  void nextTick(() => {
    controller.setActiveTextRange(localTextRange(caretOffset, caretOffset))
    restoreControllerSelection()
  })
}

// ── Watch for model changes (clears draft when paragraph changes externally) ─
watch(
  () => paragraphText(props.paragraph),
  () => {
    if (!draftDirty.value) clearDraft()
  }
)

let paragraphResizeObserver: ResizeObserver | undefined

function emitMeasuredHeight(): void {
  // A clipped line-range fragment reports only its slice height; feeding that
  // back as the paragraph height would shrink the paragraph on every pass.
  if (props.lineRange) return
  const element = paragraphEl.value
  if (!element) return
  const marginTopPx = Number.parseFloat(String(paragraphStyle.value.marginTop ?? ""))
  const marginBottomPx = Number.parseFloat(
    String(paragraphStyle.value.marginBottom ?? "")
  )
  const heightPx =
    element.offsetHeight +
    (Number.isFinite(marginTopPx) ? marginTopPx : 0) +
    (Number.isFinite(marginBottomPx) ? marginBottomPx : 0)
  if (Number.isFinite(heightPx) && heightPx > 0) {
    emit("measure", props.paragraphIndex, heightPx)
  }
}

watch(
  () => displayHtml.value,
  () => { void nextTick(emitMeasuredHeight) },
  { flush: "post" }
)

onMounted(() => {
  unregisterDraftFlusher = props.controller?.registerTextDraftFlusher(flushDraft)
  void nextTick(emitMeasuredHeight)
  if (typeof ResizeObserver !== "undefined" && paragraphEl.value) {
    paragraphResizeObserver = new ResizeObserver(() => emitMeasuredHeight())
    paragraphResizeObserver.observe(paragraphEl.value)
  }
})

onBeforeUnmount(() => {
  flushDraft()
  paragraphResizeObserver?.disconnect()
  paragraphResizeObserver = undefined
})

onUnmounted(() => {
  if (isComposing.value) props.controller?.clearSelectionSession("composition")
  unregisterDraftFlusher?.()
  unregisterDraftFlusher = undefined
})

// ── Expose ─────────────────────────────────────────────────────────
defineExpose({
  get element() { return paragraphEl.value },
  focus() { paragraphEl.value?.focus() },
  restoreDraft(text: string) {
    draftText.value = text
    draftDirty.value = text !== paragraphText(props.paragraph)
    draftBaseText = paragraphText(props.paragraph)
    draftDocumentNonce = props.controller?.documentLoadNonce
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
