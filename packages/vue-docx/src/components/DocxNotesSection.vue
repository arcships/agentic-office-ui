<template>
  <section
    v-if="notes.length"
    class="docx-notes-section"
    :class="`docx-notes-section--${kind}`"
    :data-testid="kind === 'footnote' ? 'docx-footnotes' : 'docx-endnotes'"
    :aria-label="kind === 'footnote' ? m.docx.notes.footnote : m.docx.notes.endnote"
  >
    <div class="docx-notes-section__rule" />
    <article
      v-for="(note, noteIndex) in notes"
      :key="`${kind}-${note.id}`"
      class="docx-notes-section__note"
      :data-note-id="note.id"
    >
      <sup class="docx-notes-section__marker" data-docx-note-body-marker="true">{{ noteLabel(note, noteIndex) }}</sup>
      <div class="docx-notes-section__content">
        <template v-if="note.nodes?.length">
          <template v-for="(node, nodeIndex) in note.nodes" :key="`${note.id}-${nodeIndex}`">
            <DocxParagraphHost
              v-if="node.type === 'paragraph'"
              :paragraph="node"
              :paragraph-index="-(note.id * 1000 + nodeIndex + 1)"
              :editable="false"
              :document-theme="theme"
              :numbering-definitions="numberingDefinitions"
              :default-tab-stop-px="defaultTabStopPx"
            />
            <DocxTableHost
              v-else-if="node.type === 'table'"
              :table="node"
              :table-index="-(note.id * 1000 + nodeIndex + 1)"
              :editable="false"
              :document-theme="theme"
              :numbering-definitions="numberingDefinitions"
              :default-tab-stop-px="defaultTabStopPx"
            />
          </template>
        </template>
        <span v-else>{{ note.text }}</span>
      </div>
    </article>
  </section>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocModel, DocumentNoteDefinition, NumberingDefinitionSet } from "@arcships/docx-core"
import {
  resolveDefaultTabStopPx,
  resolveDocumentLayoutDefaults,
} from "@arcships/docx-core"
import DocxParagraphHost from "./DocxParagraphHost.vue"
import DocxTableHost from "./DocxTableHost.vue"
import { useDocxMessages } from "../i18n"

const props = withDefaults(defineProps<{
  kind: "footnote" | "endnote"
  notes: DocumentNoteDefinition[]
  model?: DocModel
  numberingDefinitions?: NumberingDefinitionSet
  /**
   * Note id → display label computed from the document's note numbering
   * (numFmt/numRestart/numStart). Keeps the note body label identical to the
   * in-text marker; without it the index inside this section is used.
   */
  labels?: Map<number, string>
  startIndex?: number
  theme?: "light" | "dark"
}>(), {
  startIndex: 1,
  theme: "light",
})

// Note bodies render zero-margin paragraphs (see the scoped style below), so
// only the document tab grid reaches them.
function noteLayoutDefaults() {
  return resolveDocumentLayoutDefaults(props.model?.metadata)
}
const defaultTabStopPx = computed(() =>
  resolveDefaultTabStopPx(noteLayoutDefaults().defaultTabStopTwips)
)

function noteLabel(note: DocumentNoteDefinition, noteIndex: number): string {
  const label = props.labels?.get(note.id)
  if (label) return label
  return String(noteIndex + props.startIndex)
}

const m = useDocxMessages()
</script>

<style scoped>
.docx-notes-section { color: inherit; font-size: 10px; line-height: 1.35; margin-top: auto; padding-top: 10px; }
.docx-notes-section--endnote { margin-top: 18px; }
.docx-notes-section__rule { border-top: 1px solid currentColor; margin-bottom: 7px; opacity: .45; width: 34%; }
.docx-notes-section__note { display: grid; gap: 6px; grid-template-columns: 18px minmax(0, 1fr); margin: 3px 0; }
.docx-notes-section__marker { font-size: 9px; line-height: 1.4; text-align: right; }
.docx-notes-section__content { min-width: 0; }
.docx-notes-section__content :deep(.docx-paragraph-host) { font-size: inherit !important; line-height: inherit !important; margin: 0 !important; min-height: 0; }
.docx-notes-section__content :deep(.docx-table-host) { font-size: inherit; line-height: inherit; }
.docx-notes-section__content :deep(.docx-table-host-table) { width: 100%; border-collapse: collapse; }
</style>
