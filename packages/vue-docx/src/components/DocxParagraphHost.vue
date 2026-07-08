<template>
  <div
    ref="paragraphEl"
    class="docx-paragraph-host"
    :class="{
      'docx-paragraph-host--editable': editable,
      'docx-paragraph-host--active': isActiveParagraph,
    }"
    :contenteditable="editable ? 'true' : undefined"
    :data-docx-paragraph-host="true"
    :data-docx-paragraph-node-index="paragraphIndex"
    :suppresscontenteditablewarning="editable ? 'true' : undefined"
    :style="paragraphStyle"
    :spellcheck="editable ? 'true' : undefined"
    @input="onInput"
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
import { ref, computed, watch, inject, nextTick, onMounted, onUnmounted } from "vue"
import type {
  DocxEditorController,
  ParagraphNode,
  DocxDocumentTheme,
} from "@extend-ai/docx-core"
import {
  paragraphText,
  paragraphIsEffectivelyEmpty,
} from "@extend-ai/docx-core"
import { renderParagraphRuns, type ParagraphRunRenderOptions } from "../render/paragraph-runs"
import { renderStaticHtml } from "../render/static-html"

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
    controller: DocxEditorController
    numberingLabel?: string
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
const isComposing = ref(false)
const lastCommittedTextLength = ref(0)

// ── Computed ───────────────────────────────────────────────────────
const isActiveParagraph = computed(() => {
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
  const text = draftText.value ?? paragraphText(props.paragraph)
  const theme = props.documentTheme
  const runOptions: ParagraphRunRenderOptions = {
    showTrackedChanges: false,
    showCommentHighlights: false,
  }

  const vnodes = renderParagraphRuns(
    props.paragraph,
    `p-${props.paragraphIndex}`,
    theme,
    props.numberingLabel ? { label: props.numberingLabel } as any : undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    runOptions
  )

  // If there's a draft, render draft text instead of VNodes
  if (draftText.value !== null) {
    return escapeHtml(text)
  }

  const result = Array.isArray(vnodes) ? vnodes : [vnodes]
  return result.map((v) => renderStaticHtml(v)).join("")
})

// ── Draft management ───────────────────────────────────────────────
function readCurrentDomText(): string {
  if (!paragraphEl.value) return ""
  return paragraphEl.value.textContent ?? ""
}

function syncDraftFromDom(): void {
  if (!props.editable || !paragraphEl.value) return
  const domText = readCurrentDomText()
  draftText.value = domText
}

function clearDraft(): void {
  draftText.value = null
  lastCommittedTextLength.value = 0
}

function commitDraft(): void {
  if (draftText.value === null) return
  const text = draftText.value
  clearDraft()
  emit("textInput", props.paragraphIndex, text)
}

// ── Event handlers ─────────────────────────────────────────────────
function onInput(event: Event): void {
  if (!props.editable || isComposing.value) return
  const target = event.target as HTMLElement
  draftText.value = target.textContent ?? ""
}

function onKeydown(event: KeyboardEvent): void {
  if (!props.editable) return

  // Enter → commit and split
  if (event.key === "Enter" && !event.shiftKey && !isComposing.value) {
    event.preventDefault()
    syncDraftFromDom()
    commitDraft()
    props.controller.splitParagraphAtSelection("", 0)
    return
  }

  // Backspace at start of empty paragraph → merge with previous
  if (event.key === "Backspace") {
    syncDraftFromDom()
    const text = draftText.value ?? ""
    if (text.length === 0) {
      event.preventDefault()
      // Delete the paragraph (handled by controller)
    }
    return
  }
}

function onKeyup(_event: KeyboardEvent): void {
  if (!props.editable) return
  syncDraftFromDom()
}

function onClick(event: MouseEvent): void {
  emit("click", props.paragraphIndex, event)
}

function onFocus(): void {
  if (!props.editable) return
  clearDraft()
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
}

function onCompositionEnd(event: Event): void {
  isComposing.value = false
  const target = event.target as HTMLElement
  draftText.value = target.textContent ?? ""
}

function onPaste(event: ClipboardEvent): void {
  if (!props.editable) return
  event.preventDefault()
  const text = event.clipboardData?.getData("text/plain") ?? ""
  if (text) {
    // Insert plain text at cursor
    document.execCommand("insertText", false, text)
    syncDraftFromDom()
  }
}

// ── Watch for model changes (clears draft when paragraph changes externally) ─
watch(
  () => paragraphText(props.paragraph),
  (newText) => {
    if (draftText.value === null) {
      lastCommittedTextLength.value = newText.length
    }
  }
)

// ── HTML escape helper ─────────────────────────────────────────────
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;")
}

// ── Expose ─────────────────────────────────────────────────────────
defineExpose({
  get element() { return paragraphEl.value },
  focus() { paragraphEl.value?.focus() },
  restoreDraft(text: string) { draftText.value = text },
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
