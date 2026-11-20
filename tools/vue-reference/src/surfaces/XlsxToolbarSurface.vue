<script setup lang="ts">
import ParityCell from "../parity/ParityCell.vue";
import XlsxToolbar from "../../../../packages/vue-xlsx/src/components/XlsxToolbar.vue";
import { xlsxChromeController } from "./xlsx-controller";

defineProps<{ cellWidth: number }>();

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
    <XlsxToolbar
      :controller="xlsxChromeController({ disabled: entry.disabled })"
      :show-upload="true"
    />
  </ParityCell>
</template>
