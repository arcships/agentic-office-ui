<script setup lang="ts">
import { ref } from "vue";
import { OfficeSelect } from "@arcships/vue-ui";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const copy = fixtures.copy.select;
const value = ref("word");
const empty = ref("");

const options = [
  { value: "word", label: copy.options[0] },
  { value: "pdf", label: copy.options[1] },
  { value: "slides", label: copy.options[2] },
 ];

defineProps<{ cellWidth: number }>();
</script>

<template>
  <ParityCell state="default" matrix-state="default" :width="cellWidth">
    <OfficeSelect
      v-model="value"
      :options="options"
      :aria-label="copy.label"
    />
  </ParityCell>
  <ParityCell state="focus" matrix-state="focus" :width="cellWidth">
    <OfficeSelect
      v-model="value"
      :options="options"
      :aria-label="copy.label"
    />
  </ParityCell>
  <ParityCell state="disabled" matrix-state="disabled" :width="cellWidth">
    <OfficeSelect
      v-model="value"
      :options="options"
      :aria-label="copy.label"
      disabled
    />
  </ParityCell>
  <ParityCell state="placeholder" matrix-state="default" :width="cellWidth">
    <OfficeSelect
      v-model="empty"
      :options="options"
      :aria-label="copy.label"
      :placeholder="copy.placeholder"
    />
  </ParityCell>
</template>
