<script setup lang="ts">
/**
 * Issue #158 — `OfficeErrorState` subject.
 *
 * Mirrors `tools/extend-reference/src/surfaces/OfficeErrorStateSurface.tsx`: the
 * pinned viewer error card (`max-w-md rounded-lg border bg-background p-4
 * text-sm`, `font-medium` title, `mt-1` muted message, `mt-4` outline action)
 * inside the upstream centered host (`grid place-items-center bg-muted/30 p-4`).
 * The card is the parity target, so its padding, border, radius and type are
 * compared directly.
 */
import OfficeErrorState from "../../../../packages/vue-ui/src/components/chrome/OfficeErrorState.vue";
import ParityCell from "../parity/ParityCell.vue";

defineProps<{ cellWidth: number }>();

const states = [
  { state: "default", matrixState: "error", showRetry: true, long: false },
  { state: "message", matrixState: "error", showRetry: false, long: false },
  { state: "wrapped", matrixState: "error", showRetry: true, long: true },
];

const longDescription =
  "The worker returned an unexpected result while rendering page 5 of the document. Check the file, then retry or download a fresh copy.";
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <div class="office-error-frame">
      <OfficeErrorState
        :show-retry="entry.showRetry"
        :description="entry.long ? longDescription : undefined"
      />
    </div>
  </ParityCell>
</template>

<style scoped>
/* Upstream host: `grid place-items-center bg-muted/30 p-4`. */
.office-error-frame {
  display: grid;
  place-items: center;
  width: 100%;
  padding: 16px;
  background: color-mix(in oklab, var(--office-muted) 30%, transparent);
}
</style>
