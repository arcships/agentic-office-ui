<script setup lang="ts">
import { ref } from "vue";
import { OfficeInput } from "@arcships/vue-ui";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const copy = fixtures.copy.input;
const value = ref(copy.value);
const invalidValue = ref(copy.invalid);
const disabledValue = ref(copy.disabled);

defineProps<{ cellWidth: number }>();
</script>

<template>
  <ParityCell state="default" matrix-state="default" :width="cellWidth">
    <OfficeInput v-model="value" size="default" :aria-label="copy.label" />
  </ParityCell>
  <ParityCell state="focus" matrix-state="focus" :width="cellWidth">
    <OfficeInput v-model="value" size="default" :aria-label="copy.label" />
  </ParityCell>
  <ParityCell state="disabled" matrix-state="disabled" :width="cellWidth">
    <OfficeInput v-model="disabledValue" :aria-label="copy.label" disabled />
  </ParityCell>
  <ParityCell state="invalid" matrix-state="error" :width="cellWidth">
    <OfficeInput v-model="invalidValue" :aria-label="copy.label" invalid />
  </ParityCell>
  <ParityCell state="size-sm" matrix-state="default" :width="cellWidth">
    <OfficeInput v-model="value" size="sm" :aria-label="copy.label" />
  </ParityCell>
</template>
