<template>
  <Teleport to="body">
    <div
      v-if="message"
      class="xlsx-clipboard-notice"
      data-testid="xlsx-clipboard-notice"
      role="status"
      aria-live="polite"
    >
      {{ message }}
    </div>
  </Teleport>
</template>

<script setup lang="ts">
defineProps<{ message: string }>();
</script>

<style scoped>
.xlsx-clipboard-notice {
  background: rgba(24, 24, 27, 0.92);
  border-radius: 6px;
  bottom: 24px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.24);
  color: #fafafa;
  font-size: 12px;
  left: 50%;
  max-width: 320px;
  padding: 8px 12px;
  pointer-events: none;
  position: fixed;
  transform: translateX(-50%);
  z-index: 10001;
}
</style>
