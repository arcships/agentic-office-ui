<template>
  <DocxBlockHighlights
    ref="highlights"
    :root="root"
    :blocks="targets"
    :page-index="pageIndex"
    :color="color"
    :on-geometry="onGeometry"
  />
</template>

<script setup lang="ts">
import { computed, ref } from "vue"
import type {
  DocxWordRange,
  DocumentBlock,
  DocumentWord,
} from "@arcships/docx-core"
import type { DocxWordHighlightsProps } from "./block-highlight-types"
import DocxBlockHighlights from "./DocxBlockHighlights.vue"

/**
 * Word-level highlight overlay.
 *
 * Accepts extracted `DocumentWord` objects, a single emphasized `word`, or
 * `wordRanges` (`{ blockId, start, end }`) that reference UTF-16 offsets in a
 * paragraph block. Boxes are produced by resolving the ranges through the
 * rendered text nodes, so a word spanning styles or lines gets one box per
 * visible fragment.
 */
const props = withDefaults(
  defineProps<DocxWordHighlightsProps>(),
  {
    root: null,
    word: undefined,
    words: undefined,
    wordRanges: undefined,
    pageIndex: undefined,
    color: "#f59e0b",
    onGeometry: undefined,
  },
)

const highlights = ref<InstanceType<typeof DocxBlockHighlights>>()

function wordBlock(word: DocumentWord): DocumentBlock {
  return {
    object: "block",
    id: word.id,
    type: "text",
    content: word.content,
    details: { kind: "word", blockId: word.blockId },
    sourceRefs: word.sourceRefs,
    children: [],
  }
}

function rangeBlock(range: DocxWordRange): DocumentBlock {
  const id = range.id ?? `${range.blockId}/word/${range.start}-${range.end}`
  return {
    object: "block",
    id,
    type: "text",
    content: range.content ?? "",
    details: { kind: "word", blockId: range.blockId },
    sourceRefs: [
      {
        documentId: "",
        nodeId: range.blockId,
        partName: "",
        story: "body",
        start: range.start,
        end: range.end,
        offsetEncoding: "utf-16",
      },
    ],
    children: [],
  }
}

const targets = computed<readonly DocumentBlock[]>(() => {
  const unique = new Map<string, DocumentBlock>()
  for (const word of props.words ?? []) unique.set(word.id, wordBlock(word))
  for (const range of props.wordRanges ?? []) {
    const block = rangeBlock(range)
    unique.set(block.id, block)
  }
  if (props.word) unique.set(props.word.id, wordBlock(props.word))
  return [...unique.values()]
})

defineExpose({
  refresh: () => highlights.value?.refresh?.(),
})
</script>
