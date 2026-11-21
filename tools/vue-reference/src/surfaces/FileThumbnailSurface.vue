<script setup lang="ts">
import FileThumbnail from "../../../../packages/vue-ui/src/components/FileThumbnail.vue";
import { fixtureDocument } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const pdf = fixtureDocument("pdf");
const png = fixtureDocument("png");

const states = [
  { state: "default", matrixState: "empty", url: undefined },
  { state: "loading", matrixState: "loading", url: undefined },
  { state: "error", matrixState: "error", url: undefined },
  { state: "image", matrixState: "default", url: png.asset },
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
    <FileThumbnail
      :file="{ name: pdf.name, type: pdf.type }"
      :size="80"
      :is-loading="entry.state === 'loading'"
      :has-error="entry.state === 'error'"
      :preview-image-url="entry.state === 'image' ? png.asset : null"
    />
  </ParityCell>
</template>
