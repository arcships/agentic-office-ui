<script setup lang="ts">
/**
 * Issue #158 — `OfficeStatusBar` subject.
 *
 * Mirrors `tools/extend-reference/src/surfaces/OfficeStatusBarSurface.tsx`: the
 * pinned footer contract (`border-t bg-muted/40 px-3 py-2 text-xs
 * text-muted-foreground`) with the page/status/file content. The footer root is
 * the parity target, so the 33px row, the 12/16 type and the left/right split
 * are compared.
 */
import OfficeStatusBar from "../../../../packages/vue-ui/src/components/chrome/OfficeStatusBar.vue";
import ParityCell from "../parity/ParityCell.vue";

defineProps<{ cellWidth: number }>();

const states = [{ state: "default", matrixState: "default" }];
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <OfficeStatusBar
      :page="2"
      :total-pages="9"
      status="Saved"
      file-name="report.pdf"
    />
  </ParityCell>
</template>
