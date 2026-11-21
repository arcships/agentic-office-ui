<script setup lang="ts">
/**
 * Issue #159 — DOCX toolbar subject.
 *
 * The production `DocxViewerToolbar` with its default (upstream) control set:
 * thumbnails toggle, Page + page number + of N, zoom out/level/in, separator
 * and the closed actions trigger. Every other command the product ships is in
 * the actions menu or behind an opt-in prop, so this cell measures the same row
 * as the pinned Extend reference.
 */
import { fixtures, type ParityTheme } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";
import DocxViewerToolbar from "../../../../packages/vue-docx/src/components/DocxViewerToolbar.vue";

const props = defineProps<{ cellWidth: number; theme: ParityTheme }>();
const toolbar = fixtures.docxToolbar;

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "disabled", matrixState: "empty", disabled: true },
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
    <DocxViewerToolbar
      :can-download="!entry.disabled"
      :disabled="entry.disabled"
      file-name="legal-contract.docx"
      :is-dark="props.theme === 'dark'"
      :sidebar-open="toolbar.sidebarOpen"
      :show-tracked-changes="toolbar.showTrackedChanges"
      revision-display="balloons"
      :show-comments="toolbar.showComments"
      :tracked-change-count="2"
      :comment-count="3"
      :current-page="toolbar.activePage"
      :total-pages="entry.disabled ? 0 : toolbar.pageCount"
      :zoom="toolbar.zoomPercent"
      view-mode="continuous"
      fit-mode="none"
      search-query=""
      :search-result-count="0"
      :search-result-index="-1"
      :search-case-sensitive="false"
      :search-whole-word="false"
      :search-regex="false"
      :show-upload="toolbar.showUpload"
      :show-download="toolbar.showDownload"
    />
  </ParityCell>
</template>
