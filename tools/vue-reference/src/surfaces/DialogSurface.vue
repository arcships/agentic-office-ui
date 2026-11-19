<script setup lang="ts">
import { nextTick, onMounted } from "vue";
import { OfficeButton, OfficeDialog, OfficeThemeProvider } from "@arcships/vue-ui";
import { fixtures, type ParityTheme } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

const copy = fixtures.copy.dialog;

defineProps<{ cellWidth: number; theme: ParityTheme }>();

// The parity cell captures the default visual state. Reka moves focus into the
// dialog on open; the reference fixture is static, so blur to compare surfaces
// rather than an incidental focus ring.
onMounted(async () => {
  await nextTick();
  window.setTimeout(() => {
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
  }, 120);
});
</script>

<template>
  <ParityCell state="open" matrix-state="default" :width="cellWidth">
    <OfficeThemeProvider :theme="theme">
      <OfficeDialog
        :open="true"
        :title="copy.title"
        :description="copy.description"
        :close-label="copy.close"
      >
        <template #footer>
          <OfficeButton variant="ghost">{{ copy.cancel }}</OfficeButton>
          <OfficeButton>{{ copy.save }}</OfficeButton>
        </template>
      </OfficeDialog>
    </OfficeThemeProvider>
  </ParityCell>
</template>
