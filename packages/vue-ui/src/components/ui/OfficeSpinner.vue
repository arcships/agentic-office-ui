<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import OfficeIcon from "../OfficeIcon.vue"
import { officeClass, type OfficeControlSize } from "./types"
import { useOfficeUiLabels } from "./labels"

export interface OfficeSpinnerProps {
  /** Spinner diameter preset. Defaults to the current icon size. */
  size?: OfficeControlSize
  /** Accessible name. Pass `null` to render a decorative spinner. */
  label?: string | null
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeSpinnerProps>(), {
  size: "default",
  label: undefined,
  class: "",
})

const labels = useOfficeUiLabels()
const accessibleLabel = computed(() =>
  props.label === null ? undefined : props.label || labels.value.loading,
)
</script>

<template>
  <span
    :class="officeClass('office-ui', 'office-spinner', `office-spinner--${props.size}`, props.class)"
    data-slot="spinner"
    :role="accessibleLabel ? 'status' : undefined"
    :aria-label="accessibleLabel"
    :aria-hidden="accessibleLabel ? undefined : 'true'"
  >
    <OfficeIcon name="loader-circle" />
  </span>
</template>
