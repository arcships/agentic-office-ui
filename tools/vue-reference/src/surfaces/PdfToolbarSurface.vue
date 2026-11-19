<script setup lang="ts">
import { computed } from "vue";
import PdfViewer from "../../../../packages/vue-pdf/src/components/PdfViewer.vue";
import type { OfficeUrlPolicy } from "../../../../packages/office-runtime/src/url-policy";
import { fixtures, type ParityTheme } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const props = defineProps<{ cellWidth: number; theme: ParityTheme }>();

const toolbar = fixtures.pdfToolbar;
const sample = fixtures.documents.find((entry) => entry.id === "pdf");

const urlPolicy: OfficeUrlPolicy = {
  baseUrl: window.location.href,
  allowRelativeUrl: true,
  allowedProtocols: ["http:", "https:"],
  allowedOrigins: [window.location.origin],
  allowHttpOnLocalhost: true,
};

const source = computed(() => ({
  kind: "url" as const,
  url: "/samples/sample.pdf",
  name: sample?.name ?? "sample.pdf",
}));

const states = [
  { state: "default", matrixState: "default", withSource: true },
  { state: "hover", matrixState: "hover", withSource: true },
  { state: "focus", matrixState: "focus", withSource: true },
  { state: "active", matrixState: "active", withSource: true },
  { state: "disabled", matrixState: "empty", withSource: false },
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
    <PdfViewer
      :source="entry.withSource ? source : undefined"
      :url-policy="urlPolicy"
      :show-toolbar="true"
      :show-rotate-controls="toolbar.showRotateControls"
      :show-download="toolbar.showDownload"
      :is-dark="props.theme === 'dark'"
      :default-zoom="toolbar.zoomPercent / 100"
      :file-name="sample?.name"
      style="height: 280px; width: 100%"
    />
  </ParityCell>
</template>
