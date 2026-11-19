<script setup lang="ts">
defineProps<{ state: string; matrixState: string; width: number }>();
</script>

<template>
  <div
    class="parity-cell"
    :data-parity-cell="state"
    :data-parity-matrix-state="matrixState"
    :style="{ width: `${width}px` }"
  >
    <div data-parity-host>
      <slot />
    </div>
  </div>
</template>

<style scoped>
.parity-cell {
  box-sizing: border-box;
  max-width: 100%;
  padding: 12px;
}
</style>
