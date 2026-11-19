<script setup lang="ts">
import { OfficeThemeProvider, OfficeTooltip } from "@arcships/vue-ui";
import { fixtures, type ParityTheme } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const copy = fixtures.copy.tooltip;

defineProps<{ cellWidth: number; theme: ParityTheme }>();
</script>

<template>
  <ParityCell state="open" matrix-state="hover" :width="cellWidth">
    <OfficeThemeProvider :theme="theme">
      <div class="tooltip-fixture">
        <OfficeTooltip
          :content="copy.content"
          :delay-duration="0"
          side="right"
          align="center"
        >
          <button
            class="tooltip-trigger"
            data-parity-tooltip-trigger
            type="button"
          >
            {{ copy.trigger }}
          </button>
        </OfficeTooltip>
      </div>
    </OfficeThemeProvider>
  </ParityCell>
</template>

<style scoped>
.tooltip-fixture {
  position: relative;
  display: inline-flex;
}

.tooltip-trigger {
  font-size: 14px;
  line-height: var(--office-leading-sm);
}
</style>
