<script setup lang="ts">
import ParityCell from "../parity/ParityCell.vue";
import XlsxStatusBar from "../../../../packages/vue-xlsx/src/components/XlsxStatusBar.vue";
import { xlsxChromeController } from "./xlsx-controller";

defineProps<{ cellWidth: number }>();

const states = [{ state: "default", matrixState: "default" }];
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <XlsxStatusBar :controller="xlsxChromeController()" />
  </ParityCell>
</template>
