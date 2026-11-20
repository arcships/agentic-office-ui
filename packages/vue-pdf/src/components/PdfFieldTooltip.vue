<script setup lang="ts">
/**
 * Shared `OfficeTooltip` around an AcroForm control.
 *
 * The tooltip is a browser-only enhancement: the component test renderer and
 * SSR have no DOM, so `enabled` is false there and the control renders as-is.
 * In a browser the control itself becomes the Reka trigger (`as-child`), so no
 * wrapper box affects the absolutely positioned field geometry.
 */
import { OfficeTooltip } from "@arcships/vue-ui"

const props = withDefaults(defineProps<{
  enabled?: boolean
  label?: string
}>(), {
  enabled: false,
  label: "",
})
</script>

<template>
  <OfficeTooltip
    v-if="props.enabled && props.label"
    :content="props.label"
    side="top"
  >
    <slot />
  </OfficeTooltip>
  <slot v-else />
</template>
