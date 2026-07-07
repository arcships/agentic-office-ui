<template>
  <div :class="cn('docx-editor', props.className)" :style="props.style">
    <!-- Toolbar -->
    <div class="docx-editor-toolbar" role="toolbar" aria-label="DOCX editor toolbar">
      <div class="toolbar-cluster">
        <button class="toolbar-button" @click="editor.undo()" :disabled="!editor.canUndo" title="Undo" aria-label="Undo">↶</button>
        <button class="toolbar-button" @click="editor.redo()" :disabled="!editor.canRedo" title="Redo" aria-label="Redo">↷</button>
      </div>
      <div class="toolbar-cluster">
        <select
          class="toolbar-select"
          @change="(e) => editor.setParagraphHeading(Number((e.target as HTMLSelectElement).value) || 0)"
          title="Heading level"
          aria-label="Heading level"
        >
          <option value="0">Body</option>
          <option value="1">Heading 1</option>
          <option value="2">Heading 2</option>
          <option value="3">Heading 3</option>
        </select>
        <button class="toolbar-button text-button" @click="editor.toggleBold()" title="Bold" aria-label="Bold"><b>B</b></button>
        <button class="toolbar-button text-button" @click="editor.toggleItalic()" title="Italic" aria-label="Italic"><i>I</i></button>
        <button class="toolbar-button text-button" @click="editor.toggleUnderline()" title="Underline" aria-label="Underline"><u>U</u></button>
        <button class="toolbar-button text-button" @click="editor.toggleStrike()" title="Strikethrough" aria-label="Strikethrough"><s>S</s></button>
      </div>
      <div class="toolbar-cluster">
        <button class="toolbar-button" @click="editor.setParagraphAlignment('left')" title="Align left" aria-label="Align left">L</button>
        <button class="toolbar-button" @click="editor.setParagraphAlignment('center')" title="Align center" aria-label="Align center">C</button>
        <button class="toolbar-button" @click="editor.setParagraphAlignment('right')" title="Align right" aria-label="Align right">R</button>
        <button class="toolbar-button" @click="editor.setParagraphAlignment('justify')" title="Justify" aria-label="Justify">J</button>
      </div>
      <div class="toolbar-cluster">
        <button class="toolbar-button wide-button" @click="editor.insertParagraph()" title="Insert paragraph" aria-label="Insert paragraph">+ ¶</button>
        <button class="toolbar-button wide-button" @click="editor.duplicateParagraph()" title="Duplicate paragraph" aria-label="Duplicate paragraph">Copy ¶</button>
        <button class="toolbar-button wide-button danger-button" @click="editor.removeParagraph()" title="Remove paragraph" aria-label="Remove paragraph">Del ¶</button>
      </div>
      <div class="toolbar-cluster">
        <button class="toolbar-button" @click="editor.toggleDocumentTheme()" title="Toggle theme" aria-label="Toggle theme">
          {{ editor.documentTheme.mode === 'dark' ? '☀' : '☾' }}
        </button>
        <button class="toolbar-button wide-button" @click="downloadDocx" title="Download DOCX" aria-label="Download DOCX">⬇ DOCX</button>
      </div>
      <span class="toolbar-spacer" />
      <span class="page-count">Page {{ currentPage }} / {{ totalPages }}</span>
    </div>

    <!-- Document content (uses same layout as DocxViewer) -->
    <div class="docx-editor-content">
      <div
        v-for="page in pages"
        :key="page.number"
        class="docx-page"
        :style="{
          width: `${pageWidth}px`,
          minHeight: `${pageHeight}px`,
          padding: `${marginTop}px ${marginRight}px ${marginBottom}px ${marginLeft}px`,
          margin: '0 auto 24px',
          background: editor.documentTheme.mode === 'dark' ? '#1a1a2e' : 'white',
          color: editor.documentTheme.mode === 'dark' ? '#e0e0e0' : 'inherit',
        }"
      >
        <div
          v-for="block in page.blocks"
          :key="blockKey(block)"
          class="docx-block"
          :style="{
            width: '100%',
            minHeight: `${block.height}px`,
            marginBottom: '6px',
            cursor: 'text',
          }"
          @click="selectBlock(block)"
        >
          <p
            v-if="block.kind === 'paragraph'"
            :key="`${blockKey(block)}:${paragraphText(block)}`"
            :contenteditable="true"
            @focus="selectBlock(block)"
            @input="onParagraphInput(block, $event)"
            :style="{
              fontSize: block.headingLevel ? `${headingScale(block.headingLevel) * 100}%` : '100%',
              fontWeight: block.headingLevel ? 'bold' : 'normal',
              textAlign: block.align ?? 'left',
            }"
          >
            <template v-for="run in block.children" :key="run.id">
              <span v-if="run.kind === 'text'" :style="runTextStyle(run)">{{ run.text }}</span>
              <img v-else-if="run.kind === 'image' && run.src" :src="run.src" />
            </template>
          </p>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick } from "vue"
import type { LayoutBlock, LayoutRun } from "@extend-ai/docx-core"
import { layoutDocument, resolveDocumentLayout, headingScale as headingScaleFn } from "@extend-ai/docx-core"
import { useDocxEditor } from "./composables"
import type { UseDocxEditorOptions } from "./composables"

export interface DocxEditorViewerProps {
  editor: ReturnType<typeof useDocxEditor>
  className?: string
  style?: Record<string, string | number>
}

const props = defineProps<DocxEditorViewerProps>()
const editor = props.editor

const layoutMetrics = computed(() =>
  resolveDocumentLayout(editor.model)
)
const pageWidth = computed(() => layoutMetrics.value.pageWidth)
const pageHeight = computed(() => layoutMetrics.value.pageHeight)
const marginTop = computed(() => layoutMetrics.value.marginTop)
const marginBottom = computed(() => layoutMetrics.value.marginBottom)
const marginLeft = computed(() => layoutMetrics.value.marginLeft)
const marginRight = computed(() => layoutMetrics.value.marginRight)

const pages = computed(() => layoutDocument(editor.model))
const totalPages = computed(() => Math.max(1, pages.value.length))
const currentPage = computed(() => Math.min(editor.currentPage || 1, totalPages.value))

async function downloadDocx() {
  const blob = await editor.exportDocx()
  const url = URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = editor.fileName || "Document.docx"
  a.rel = "noopener"
  document.body.append(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 0)
}

function runTextStyle(run: LayoutRun) {
  const s = run.style
  if (!s) return {}
  return {
    fontWeight: s.bold ? "bold" : "normal",
    fontStyle: s.italic ? "italic" : "normal",
    textDecoration: s.underline ? "underline" : s.strike ? "line-through" : "none",
    color: s.color ?? "inherit",
    fontFamily: s.fontFamily ?? "inherit",
    fontSize: s.fontSizePt ? `${s.fontSizePt}pt` : "inherit",
  }
}

function blockKey(block: LayoutBlock) {
  return block.id
}

function paragraphText(block: LayoutBlock) {
  return block.kind === "paragraph"
    ? (block.children ?? []).map((run) => run.kind === "text" ? run.text : "").join("")
    : ""
}

function topLevelParagraphIndex(block: LayoutBlock) {
  const match = /^paragraph-(\d+)$/.exec(block.id)
  return match ? Number(match[1]) : -1
}

function selectBlock(block: LayoutBlock) {
  const nodeIndex = topLevelParagraphIndex(block)
  if (nodeIndex >= 0) editor.selectParagraph?.(nodeIndex, 0)
}

function caretOffsetWithin(element: HTMLElement): number {
  const selection = window.getSelection()
  if (!selection || selection.rangeCount === 0) return element.innerText.length
  const range = selection.getRangeAt(0)
  const before = range.cloneRange()
  before.selectNodeContents(element)
  before.setEnd(range.endContainer, range.endOffset)
  return before.toString().length
}

function restoreCaret(element: HTMLElement, offset: number) {
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT)
  let remaining = offset
  let node = walker.nextNode()
  while (node) {
    const length = node.textContent?.length ?? 0
    if (remaining <= length) {
      const range = document.createRange()
      const selection = window.getSelection()
      range.setStart(node, remaining)
      range.collapse(true)
      selection?.removeAllRanges()
      selection?.addRange(range)
      return
    }
    remaining -= length
    node = walker.nextNode()
  }
  const range = document.createRange()
  const selection = window.getSelection()
  range.selectNodeContents(element)
  range.collapse(false)
  selection?.removeAllRanges()
  selection?.addRange(range)
}

function onParagraphInput(block: LayoutBlock, event: Event) {
  const target = event.target as HTMLElement
  const text = target.innerText
  const caret = caretOffsetWithin(target)
  const nodeIndex = topLevelParagraphIndex(block)
  if (nodeIndex >= 0) {
    editor.updateParagraphText?.(nodeIndex, text)
    nextTick(() => restoreCaret(target, Math.min(caret, target.innerText.length)))
  }
}

function cn(...classes: (string | undefined | null | false)[]): string {
  return classes.filter(Boolean).join(" ")
}

const headingScale = headingScaleFn
</script>

<style scoped>
.docx-editor {
  display: flex;
  min-height: 300px;
  min-width: 0;
  flex-direction: column;
  overflow: hidden;
  background: var(--background, #fff);
}
.docx-editor-toolbar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 6px;
  border-bottom: 1px solid var(--border, #e5e7eb);
  background: color-mix(in oklch, var(--background, #fff) 82%, var(--muted, #f5f5f5));
  padding: 6px;
  flex-shrink: 0;
}
.toolbar-cluster {
  display: inline-flex;
  min-width: 0;
  align-items: center;
  gap: 4px;
  border: 1px solid var(--border, #e5e7eb);
  border-radius: var(--radius, 10px);
  background: color-mix(in oklch, var(--background, #fff) 90%, transparent);
  padding: 3px;
  box-shadow: var(--shadow-sm, 0 1px 2px rgb(0 0 0 / 0.05));
}
.toolbar-button,
.toolbar-select {
  height: 28px;
  border: 1px solid transparent;
  border-radius: calc(var(--radius, 10px) - 4px);
  background: transparent;
  color: var(--foreground, #111827);
  font-size: 12px;
  line-height: 1;
}
.toolbar-button {
  display: inline-flex;
  min-width: 28px;
  align-items: center;
  justify-content: center;
  padding: 0 8px;
  cursor: pointer;
}
.wide-button { min-width: 40px; }
.danger-button:hover:not(:disabled) { border-color: color-mix(in oklch, var(--destructive, #dc2626) 45%, var(--border, #e5e7eb)); color: var(--destructive, #dc2626); }
.toolbar-button:hover:not(:disabled),
.toolbar-select:hover:not(:disabled) {
  background: var(--accent, #f5f5f5);
  border-color: var(--border, #e5e7eb);
}
.toolbar-button:active:not(:disabled) { background: var(--secondary, #f4f4f5); }
.toolbar-button:disabled { cursor: not-allowed; opacity: 0.45; }
.text-button { font-size: 13px; font-family: Georgia, 'Times New Roman', serif; }
.toolbar-select { min-width: 104px; padding: 0 26px 0 8px; cursor: pointer; }
.toolbar-spacer { flex: 1 1 auto; min-width: 12px; }
.page-count { color: var(--muted-foreground, #737373); font-size: 12px; padding: 0 6px; white-space: nowrap; }
.docx-editor-content {
  flex: 1;
  overflow: auto;
  min-width: 0;
  background: color-mix(in oklch, var(--muted, #f5f5f5) 58%, var(--background, #fff));
  padding: 24px;
}
.docx-page {
  border: 1px solid color-mix(in oklch, var(--border, #e5e7eb) 75%, transparent);
  box-shadow: 0 12px 28px rgb(15 23 42 / 0.10), 0 2px 6px rgb(15 23 42 / 0.08);
}
.docx-block p:focus-visible {
  border-radius: 4px;
}
@media (max-width: 640px) {
  .docx-editor-content { padding: 12px; }
  .toolbar-cluster { max-width: 100%; overflow-x: auto; }
}
</style>
