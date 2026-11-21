<script setup lang="ts">
import FileUpload from "../../../../packages/vue-ui/src/components/FileUpload.vue";
import ParityCell from "../parity/ParityCell.vue";

const states = [
  { state: "default", matrixState: "empty" },
  { state: "hover", matrixState: "hover" },
  { state: "focus", matrixState: "focus" },
  { state: "dragging", matrixState: "active" },
  { state: "rejected", matrixState: "error" },
  { state: "file-selected", matrixState: "default" },
];

defineProps<{ cellWidth: number }>();
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <!-- No props: the cell measures the component's shipped defaults. -->
    <FileUpload />
  </ParityCell>
</template>
