<script setup lang="ts">
import ParityCell from "../parity/ParityCell.vue";
import XlsxSheetTabs from "../../../../packages/vue-xlsx/src/components/XlsxSheetTabs.vue";
import { xlsxChromeController } from "./xlsx-controller";

defineProps<{ cellWidth: number }>();

const states = [
  { state: "default", matrixState: "default" },
  { state: "hover", matrixState: "hover" },
  { state: "focus", matrixState: "focus" },
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
    <XlsxSheetTabs :controller="xlsxChromeController()" panel-id="xlsx-grid-panel" />
  </ParityCell>
</template>
