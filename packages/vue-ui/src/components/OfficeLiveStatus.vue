<template>
  <div
    class="office-live-status"
    :class="className"
    :role="politeness === 'assertive' ? 'alert' : 'status'"
    :aria-live="politeness"
    aria-atomic="true"
    :aria-busy="busy ? 'true' : undefined"
    :data-live-status="politeness"
    :data-testid="testId"
  >
    <span class="office-live-status__text">{{ message }}</span>
  </div>
</template>

<script setup lang="ts">
withDefaults(
  defineProps<{
    /** The text announced to assistive technology. Empty text announces nothing. */
    message?: string
    /** `polite` waits for idle, `assertive` interrupts. Defaults to `polite`. */
    politeness?: "polite" | "assertive"
    /** Marks the announced operation as still running. */
    busy?: boolean
    className?: string
    testId?: string
  }>(),
  {
    message: "",
    politeness: "polite",
    busy: false,
    className: undefined,
    testId: undefined,
  },
)
</script>

<style scoped>
.office-live-status {
  block-size: 1px;
  clip-path: inset(50%);
  inline-size: 1px;
  overflow: hidden;
  position: absolute;
  white-space: nowrap;
}
</style>
