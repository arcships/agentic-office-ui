<script setup lang="ts">
import fixtures from "../../../parity-fixtures/fixtures.json";
import ESignatureFields from "../../../../packages/vue-ui/src/components/ESignatureFields.vue";
import ParityCell from "../parity/ParityCell.vue";

const signature = fixtures.blocks.signature;
const unsignedField = {
  id: "signature-1",
  label: signature.label,
  page: signature.page,
  bbox: signature.bbox,
};
const signedField = { ...unsignedField, imageDataUrl: signature.imageDataUrl };

const states = [
  { state: "unsigned", matrixState: "default" },
  { state: "signed", matrixState: "default" },
  { state: "exporting", matrixState: "loading" },
  { state: "preview", matrixState: "default" },
  { state: "loading", matrixState: "loading" },
  { state: "error", matrixState: "error" },
];

defineProps<{ cellWidth: number }>();

function isSigned(state: string) {
  return state === "signed" || state === "exporting" || state === "preview";
}
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <ESignatureFields
      :fields="isSigned(entry.state) ? [signedField] : [unsignedField]"
      :can-export="isSigned(entry.state)"
      :exporting="entry.state === 'exporting'"
      :show-preview="entry.state === 'preview'"
      :loading="entry.state === 'loading'"
      :error="
        entry.state === 'error'
          ? 'Signature fields could not be loaded.'
          : undefined
      "
    />
  </ParityCell>
</template>
