<script setup lang="ts">
import fixtures from "../../../parity-fixtures/fixtures.json";
import LayoutBlocks from "../../../../packages/vue-ui/src/components/LayoutBlocks.vue";
import ParityCell from "../parity/ParityCell.vue";

const blocks = fixtures.blocks.layoutBlocks;

const states = [
  { state: "default", matrixState: "default", activeBlockId: undefined },
  { state: "hover", matrixState: "hover", activeBlockId: undefined },
  { state: "focus", matrixState: "focus", activeBlockId: undefined },
  { state: "active", matrixState: "default", activeBlockId: "block-2" },
  { state: "empty", matrixState: "empty", activeBlockId: undefined },
  { state: "loading", matrixState: "loading", activeBlockId: undefined },
  { state: "error", matrixState: "error", activeBlockId: undefined },
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
    <LayoutBlocks
      view="panel"
      :blocks="entry.state === 'empty' ? [] : blocks"
      :active-block-id="entry.activeBlockId"
      :loading="entry.state === 'loading'"
      :error="
        entry.state === 'error'
          ? 'Layout blocks could not be loaded.'
          : undefined
      "
    />
  </ParityCell>
</template>
