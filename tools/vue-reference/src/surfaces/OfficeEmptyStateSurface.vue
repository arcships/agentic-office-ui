<script setup lang="ts">
/**
 * Issue #158 — `OfficeEmptyState` subject.
 *
 * Mirrors `tools/extend-reference/src/surfaces/OfficeEmptyStateSurface.tsx`: the
 * pinned `PdfEditorEmptyState` layout (gap-2 px-4 py-10, `text-sm font-medium`
 * title, `max-w-56 text-xs` description, `mt-2` action). `default` renders the
 * catalog copy with no props, `action` fills the `actions` slot with the shared
 * outline/sm button.
 */
import OfficeEmptyState from "../../../../packages/vue-ui/src/components/chrome/OfficeEmptyState.vue";
import OfficeButton from "../../../../packages/vue-ui/src/components/ui/OfficeButton.vue";
import ParityCell from "../parity/ParityCell.vue";

defineProps<{ cellWidth: number }>();

const states = [
  { state: "default", matrixState: "default", action: false },
  { state: "action", matrixState: "default", action: true },
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
    <OfficeEmptyState>
      <template v-if="entry.action" #actions>
        <OfficeButton variant="outline" size="sm" type="button">Open</OfficeButton>
      </template>
    </OfficeEmptyState>
  </ParityCell>
</template>
