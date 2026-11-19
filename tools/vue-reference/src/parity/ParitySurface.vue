<script setup lang="ts">
import type { ParityTheme } from "../fixtures";

defineProps<{ surface: string; theme: ParityTheme }>();
</script>

<template>
  <div
    :class="[theme === 'dark' ? 'dark' : '', 'parity-surface']"
    data-parity-app="vue-reference"
    :data-parity-surface="surface"
    :data-parity-theme="theme"
  >
    <slot />
  </div>
</template>

<style scoped>
.parity-surface {
  min-height: 100vh;
  box-sizing: border-box;
  padding: 24px;
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 16px;
  background: var(--background);
  color: var(--foreground);
}
</style>
