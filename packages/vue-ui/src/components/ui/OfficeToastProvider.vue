<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import { ToastProvider } from "reka-ui"
import OfficeToastViewport from "./OfficeToastViewport.vue"
import { useOfficeUiLabels } from "./labels"

export interface OfficeToastProviderProps {
  /** Default visible duration in milliseconds. */
  duration?: number
  /** Maximum number of visible toasts; older ones stay inert behind the limit. */
  limit?: number
  /** Screen-reader label for each toast. Defaults to a localized "Notifications". */
  label?: string
  /** Accessible label for every close button. */
  closeLabel?: string
  /** Screen corner the stack is anchored to. */
  position?: "top-left" | "top-right" | "bottom-left" | "bottom-right"
  /** Disable swipe-to-dismiss. */
  disableSwipe?: boolean
}

const props = withDefaults(defineProps<OfficeToastProviderProps>(), {
  duration: 5000,
  limit: undefined,
  label: undefined,
  closeLabel: undefined,
  position: "bottom-right",
  disableSwipe: false,
})

const labels = useOfficeUiLabels()
const resolvedLabel = computed(() => props.label ?? labels.value.notifications)
const resolvedCloseLabel = computed(() => props.closeLabel ?? labels.value.close)
</script>

<template>
  <ToastProvider
    :duration="props.duration"
    :limit="props.limit"
    :label="resolvedLabel"
    :disable-swipe="props.disableSwipe"
  >
    <slot />
    <OfficeToastViewport
      :label="resolvedLabel"
      :close-label="resolvedCloseLabel"
      :position="props.position"
    />
  </ToastProvider>
</template>
