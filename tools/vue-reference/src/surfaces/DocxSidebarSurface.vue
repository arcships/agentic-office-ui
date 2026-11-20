<script setup lang="ts">
/**
 * Issue #159 — DOCX thumbnail sidebar subject.
 *
 * Mirrors the reference surface: the production `DocxThumbnailPanel` rendered
 * inside the same 160×292 frame, the same single-page fixture, the same page
 * aspect ratio (200:300) and the upstream inline/overlay decision
 * (`window.innerWidth >= 768`), so the rail, card, label and skeleton are
 * compared one to one.
 */
import { computed } from "vue";
import { fixtures, type ParityTheme } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";
import DocxThumbnailPanel from "../../../../packages/vue-docx/src/components/DocxThumbnailPanel.vue";
import { createBlankDocumentModel } from "../../../../packages/docx-core/src/editor/helpers/default-model";

const props = defineProps<{ cellWidth: number; theme: ParityTheme }>();
const sidebar = fixtures.documentSidebar;

const base = createBlankDocumentModel();
const model = {
  ...base,
  nodes: Array.from({ length: 12 }, (_, index) => ({
    type: "paragraph" as const,
    children: [{ type: "text" as const, text: `Legal contract clause ${index + 1}` }],
  })),
};
const layoutOptions = {
  pageWidth: 200,
  pageHeight: 300,
  margin: 20,
  minLineHeight: 20,
  paragraphSpacing: 6,
};

const inline = computed(() => window.innerWidth >= 768);
const mode = computed(() => (inline.value ? "inline" : "overlay"));

const states = [
  { state: "open", matrixState: "default", collapsed: false, loading: false },
  { state: "closed", matrixState: "default", collapsed: true, loading: false },
  { state: "skeleton", matrixState: "loading", collapsed: false, loading: true },
];
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <div
      :style="{
        width: `${sidebar.widthPx}px`,
        height: entry.state === 'skeleton' ? '240px' : '292px',
        overflow: 'hidden',
        position: 'relative',
        background: 'var(--office-sidebar)',
      }"
    >
      <DocxThumbnailPanel
        :model="entry.loading ? undefined : model"
        :layout-options="layoutOptions"
        :default-collapsed="entry.collapsed"
        :mode="mode"
        :show-toggle="false"
        :current-page="1"
        :total-pages="entry.loading ? 0 : 1"
        :is-dark="props.theme === 'dark'"
      />
    </div>
  </ParityCell>
</template>
