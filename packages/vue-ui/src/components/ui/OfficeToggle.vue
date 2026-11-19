<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import { Toggle } from "reka-ui"
import {
  officeClass,
  type OfficeControlSize,
} from "./types"

export interface OfficeToggleProps {
  /** Controlled pressed state (`v-model`). */
  modelValue?: boolean
  /** Visual style. `outline` adds the bordered surface. */
  variant?: "default" | "outline"
  size?: OfficeControlSize
  disabled?: boolean
  /** Accessible name when the slot has no text label. */
  label?: string
  asChild?: boolean
  class?: string
}

const props = withDefaults(defineProps<OfficeToggleProps>(), {
  modelValue: false,
  variant: "default",
  size: "default",
  disabled: false,
  label: undefined,
  asChild: false,
  class: "",
})

const emit = defineEmits<{
  "update:modelValue": [value: boolean]
}>()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-toggle",
    props.variant === "outline" && "office-toggle--outline",
    props.size !== "default" && `office-toggle--${props.size}`,
    props.class,
  ),
)
</script>

<template>
  <Toggle
    :model-value="props.modelValue"
    :disabled="props.disabled"
    :as-child="props.asChild"
    :class="classes"
    :aria-label="props.label"
    data-slot="toggle"
    @update:model-value="emit('update:modelValue', $event)"
  >
    <slot />
  </Toggle>
</template>
