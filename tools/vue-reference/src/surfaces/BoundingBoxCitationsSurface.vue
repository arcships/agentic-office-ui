<script setup lang="ts">
import fixtures from "../../../parity-fixtures/fixtures.json";
import BoundingBoxCitations from "../../../../packages/vue-ui/src/components/BoundingBoxCitations.vue";
import ParityCell from "../parity/ParityCell.vue";

const review = fixtures.blocks.boundingBoxCitations;
const fields = review.fields;

const states = [
  { state: "default", matrixState: "default" },
  { state: "hover", matrixState: "hover" },
  { state: "focus", matrixState: "focus" },
  { state: "selected", matrixState: "default" },
  { state: "modified", matrixState: "default" },
  { state: "empty", matrixState: "empty" },
  { state: "loading", matrixState: "loading" },
  { state: "error", matrixState: "error" },
];

defineProps<{ cellWidth: number }>();

function stateFields(state: string) {
  return state === "empty" ? [] : fields;
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
    <BoundingBoxCitations
      view="panel"
      :review-fields="stateFields(entry.state)"
      :active-field-key="entry.state === 'selected' ? 'invoice_total' : undefined"
      :expected-values="
        entry.state === 'modified'
          ? { [review.modifiedFieldKey]: review.modifiedExpected }
          : undefined
      "
      :loading="entry.state === 'loading'"
      :error="
        entry.state === 'error'
          ? 'Review fields could not be loaded.'
          : undefined
      "
    />
  </ParityCell>
</template>
