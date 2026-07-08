<template>
  <div
    ref="paragraphEl"
    class="docx-paragraph-host"
    :class="{ 'docx-paragraph-editable': editable }"
    :contenteditable="editable ? 'true' : undefined"
    :data-paragraph-index="paragraphIndex"
    @input="onInput"
    @keydown="onKeydown"
    @click="onClick"
    @focus="onFocus"
    @blur="onBlur"
    v-html="renderedHtml"
  />
</template>

<script setup lang="ts">
import { ref, computed } from "vue"
import type { ParagraphNode, DocxDocumentTheme } from "@extend-ai/docx-core"
import { renderParagraphRuns } from "../render/paragraph-runs"
import { renderStaticHtml } from "../render/static-html"

const HEADING_FONT_SIZES: Record<number, string> = {
  1: "2rem", 2: "1.6rem", 3: "1.35rem", 4: "1.2rem", 5: "1.05rem", 6: "0.95rem",
}

const props = withDefaults(
  defineProps<{
    paragraph: ParagraphNode
    paragraphIndex: number
    editable?: boolean
    documentTheme?: DocxDocumentTheme
    numberingLabel?: import("@extend-ai/docx-core").ParagraphNumberingLabel
    floatingPageOriginPx?: {
      left: number; top: number
      marginLeft?: number; marginTop?: number
      columnLeft?: number; columnTop?: number
      pageWidth?: number
    }
  }>(),
  { editable: false, documentTheme: "light" as DocxDocumentTheme }
)

const emit = defineEmits<{
  textInput: [index: number, draft: string]
  focus: [index: number]
  blur: [index: number]
  click: [index: number, event: MouseEvent]
}>()

const paragraphEl = ref<HTMLElement | undefined>()

const renderedHtml = computed(() => {
  const vnodes = renderParagraphRuns(
    props.paragraph,
    `p-${props.paragraphIndex}`,
    props.documentTheme,
    props.numberingLabel,
    undefined,
    props.floatingPageOriginPx
  )
  const result = Array.isArray(vnodes) ? vnodes : [vnodes]
  return result.map((v) => renderStaticHtml(v)).join("")
})

function onInput(event: Event) {
  const target = event.target as HTMLElement
  emit("textInput", props.paragraphIndex, target.textContent ?? "")
}

function onKeydown(_event: KeyboardEvent) {}

function onClick(event: MouseEvent) {
  emit("click", props.paragraphIndex, event)
}

function onFocus() {
  emit("focus", props.paragraphIndex)
}

function onBlur() {
  emit("blur", props.paragraphIndex)
}
</script>

<style scoped>
.docx-paragraph-host {
  margin: 0;
  min-height: 1em;
  outline: none;
  white-space: pre-wrap;
  word-wrap: break-word;
}
.docx-paragraph-editable {
  cursor: text;
}
.docx-paragraph-editable:empty::before {
  content: attr(data-placeholder);
  color: #9ca3af;
}
</style>
