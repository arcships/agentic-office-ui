<template>
  <section
    v-if="notes.length"
    class="docx-notes-section"
    :class="`docx-notes-section--${kind}`"
    :data-testid="kind === 'footnote' ? 'docx-footnotes' : 'docx-endnotes'"
    :aria-label="kind === 'footnote' ? '脚注' : '尾注'"
  >
    <div class="docx-notes-section__rule" />
    <article
      v-for="(note, noteIndex) in notes"
      :key="`${kind}-${note.id}`"
      class="docx-notes-section__note"
      :data-note-id="note.id"
    >
      <sup class="docx-notes-section__marker">{{ noteIndex + startIndex }}</sup>
      <div class="docx-notes-section__content">
        <template v-if="note.nodes?.length">
          <DocxParagraphHost
            v-for="(node, nodeIndex) in note.nodes.filter((entry): entry is ParagraphNode => entry.type === 'paragraph')"
            :key="`${note.id}-${nodeIndex}`"
            :paragraph="node"
            :paragraph-index="-(note.id * 1000 + nodeIndex + 1)"
            :editable="false"
            :document-theme="theme"
            :numbering-definitions="numberingDefinitions"
          />
        </template>
        <span v-else>{{ note.text }}</span>
      </div>
    </article>
  </section>
</template>

<script setup lang="ts">
import type { DocumentNoteDefinition, NumberingDefinitionSet, ParagraphNode } from "@arcships/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"

withDefaults(defineProps<{
  kind: "footnote" | "endnote"
  notes: DocumentNoteDefinition[]
  numberingDefinitions?: NumberingDefinitionSet
  startIndex?: number
  theme?: "light" | "dark"
}>(), {
  startIndex: 1,
  theme: "light",
})
</script>

<style scoped>
.docx-notes-section { color: inherit; font-size: 10px; line-height: 1.35; margin-top: auto; padding-top: 10px; }
.docx-notes-section--endnote { margin-top: 18px; }
.docx-notes-section__rule { border-top: 1px solid currentColor; margin-bottom: 7px; opacity: .45; width: 34%; }
.docx-notes-section__note { display: grid; gap: 6px; grid-template-columns: 18px minmax(0, 1fr); margin: 3px 0; }
.docx-notes-section__marker { font-size: 9px; line-height: 1.4; text-align: right; }
.docx-notes-section__content { min-width: 0; }
.docx-notes-section__content :deep(.docx-paragraph-host) { font-size: inherit; line-height: inherit; margin: 0 !important; min-height: 0; }
</style>
