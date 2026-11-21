<script setup lang="ts">
/**
 * Issue #158 — `OfficeLoadingState` subject.
 *
 * Mirrors `tools/extend-reference/src/surfaces/OfficeLoadingStateSurface.tsx`:
 * the pinned `ViewerLoadingSurface` (`grid h-full min-h-52 place-items-center`)
 * with the 16px spinner. `default` passes no text, which is the unsized state
 * the upstream surface renders; `hidden` covers the upstream
 * `showSpinner={false}` branch.
 */
import OfficeLoadingState from "../../../../packages/vue-ui/src/components/chrome/OfficeLoadingState.vue";
import ParityCell from "../parity/ParityCell.vue";

defineProps<{ cellWidth: number }>();

const frameHeight = 208;

const states = [
  { state: "default", matrixState: "loading", showSpinner: true },
  { state: "hidden", matrixState: "empty", showSpinner: false },
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
    <div :style="{ height: `${frameHeight}px` }">
      <OfficeLoadingState
        title=""
        description=""
        :show-spinner="entry.showSpinner"
      />
    </div>
  </ParityCell>
</template>
