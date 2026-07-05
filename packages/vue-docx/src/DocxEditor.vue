<template>
  <div :class="cn('docx-editor', props.className)" :style="props.style">
    <!-- Toolbar -->
    <div class="docx-editor-toolbar flex items-center gap-1 border-b p-1">
      <button @click="editor.undo()" :disabled="!editor.canUndo" title="Undo">↩</button>
      <button @click="editor.redo()" :disabled="!editor.canRedo" title="Redo">↪</button>
      <span class="h-4 w-px bg-border mx-1" />
      <button @click="editor.toggleBold()" title="Bold"><b>B</b></button>
      <button @click="editor.toggleItalic()" title="Italic"><i>I</i></button>
      <button @click="editor.toggleUnderline()" title="Underline"><u>U</u></button>
      <span class="h-4 w-px bg-border mx-1" />
      <select
        @change="(e) => editor.setParagraphHeading(Number((e.target as HTMLSelectElement).value) || 0)"
        title="Heading level"
      >
        <option value="0">Normal</option>
        <option value="1">Heading 1</option>
        <option value="2">Heading 2</option>
        <option value="3">Heading 3</option>
      </select>
      <span class="h-4 w-px bg-border mx-1" />
      <button @click="editor.toggleDocumentTheme()" title="Toggle dark mode">
        {{ editor.documentTheme.mode === 'dark' ? '☀' : '🌙' }}
      </button>
      <span class="flex-1" />
      <span class="text-xs text-muted-foreground">{{ editor.currentPage }} / {{ editor.totalPages }}</span>
    </div>

    <!-- Document content (uses same layout as DocxViewer) -->
    <div class="docx-editor-content overflow-auto flex-1">
      <div
        v-for="page in pages"
        :key="page.number"
        class="docx-page"
        :style="{
          width: `${pageWidth}px`,
          minHeight: `${pageHeight}px`,
          padding: `${marginTop}px ${marginRight}px ${marginBottom}px ${marginLeft}px`,
          margin: '0 auto 16px',
          background: editor.documentTheme.mode === 'dark' ? '#1a1a2e' : 'white',
          color: editor.documentTheme.mode === 'dark' ? '#e0e0e0' : 'inherit',
        }"
      >
        <div
          v-for="block in page.blocks"
          :key="blockKey(block)"
          class="docx-block"
          :style="{
            left: `${block.x}px`,
            top: `${block.y - (page.number - 1) * pageHeight - marginTop}px`,
            width: `${block.width}px`,
            minHeight: `${block.height}px`,
            position: 'relative',
            cursor: 'text',
          }"
          @click="selectBlock(block)"
        >
          <p
            v-if="block.kind === 'paragraph'"
            :contenteditable="true"
            @focus="selectBlock(block)"
            @input="onParagraphInput(block, $event)"
            :style="{
              fontSize: block.headingLevel ? `${headingScale(block.headingLevel) * 100}%` : '100%',
              fontWeight: block.headingLevel ? 'bold' : 'normal',
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
import { computed } from "vue"
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
  const text = block.kind === "paragraph" && block.children ? block.children.map((run) => run.kind === "text" ? run.text : "").join("|") : block.id
  return `${block.id}-${text}`
}

function topLevelParagraphIndex(block: LayoutBlock) {
  const match = /^paragraph-(\d+)$/.exec(block.id)
  return match ? Number(match[1]) : -1
}

function selectBlock(block: LayoutBlock) {
  const nodeIndex = topLevelParagraphIndex(block)
  if (nodeIndex >= 0) editor.selectParagraph?.(nodeIndex, 0)
}

function onParagraphInput(block: LayoutBlock, event: Event) {
  const target = event.target as HTMLElement
  const text = target.innerText
  const nodeIndex = topLevelParagraphIndex(block)
  if (nodeIndex >= 0) editor.updateParagraphText?.(nodeIndex, text)
}

function cn(...classes: (string | undefined | null | false)[]): string {
  return classes.filter(Boolean).join(" ")
}

const headingScale = headingScaleFn
</script>
