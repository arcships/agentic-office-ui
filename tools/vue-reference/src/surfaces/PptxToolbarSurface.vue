<script setup lang="ts">
/**
 * Issue #161 — PPTX toolbar parity surface.
 *
 * Subject: the production `PptxViewerToolbar` (used by `PptxViewer`) rendered
 * with the shared fixture. The reference (../extend-reference) is the vendored
 * pinned Extend `PptxToolbar` row.
 */
import PptxViewerToolbar from "../../../../packages/vue-pptx/src/PptxViewerToolbar.vue";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

defineProps<{ cellWidth: number; theme: string }>();

const toolbar = fixtures.pptxToolbar;

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "active", matrixState: "active", disabled: false },
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
    <PptxViewerToolbar
      :active-slide-index="toolbar.activeSlideIndex"
      :slide-count="entry.disabled ? 0 : toolbar.slideCount"
      :disabled="entry.disabled"
      :sidebar-open="toolbar.sidebarOpen"
      :show-sidebar-toggle="true"
      :zoom="toolbar.zoomPercent"
      :show-view-mode="true"
      :view-mode="'single'"
      :show-present-toggle="true"
      :show-notes-toggle="true"
      :notes-open="false"
      :show-fullscreen-control="true"
      :fullscreen="false"
      :show-search="true"
      :search-open="false"
      label="Presentation toolbar"
    />
  </ParityCell>
</template>
