<script setup lang="ts">
import PdfViewerToolbar from "../../../../packages/vue-pdf/src/components/PdfViewerToolbar.vue";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

defineProps<{ cellWidth: number; theme: string }>();

const toolbar = fixtures.pdfToolbar;

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
    <PdfViewerToolbar
      :active-page="toolbar.activePage"
      :total-pages="entry.disabled ? 0 : toolbar.numPages"
      :disabled="entry.disabled"
      :sidebar-open="toolbar.sidebarOpen"
      :zoom="toolbar.zoomPercent / 100"
      page-layout="continuous"
      file-name="sample.pdf"
      :show-download="toolbar.showDownload"
      :show-print="true"
      :show-rotate-controls="toolbar.showRotateControls"
      :can-download="true"
    />
  </ParityCell>
</template>
