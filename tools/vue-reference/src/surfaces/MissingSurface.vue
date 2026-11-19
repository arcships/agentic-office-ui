<script setup lang="ts">
import ParityCell from "../parity/ParityCell.vue";

const props = defineProps<{
  states: Array<{ state: string; matrixState: string }>;
  cellWidth: number;
  pendingIssue: string;
  note: string;
}>();
</script>

<template>
  <ParityCell
    v-for="entry in props.states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="props.cellWidth"
  >
    <p class="parity-missing" :data-parity-missing="props.pendingIssue">
      No Vue subject yet ({{ props.pendingIssue }}): {{ props.note }}
    </p>
  </ParityCell>
</template>

<style scoped>
.parity-missing {
  max-width: 100%;
  color: var(--muted-foreground);
  font-size: 13px;
  line-height: 1.4;
}
</style>
