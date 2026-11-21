<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import { Primitive } from "reka-ui"
import OfficeSpinner from "./OfficeSpinner.vue"
import {
  officeClass,
  type OfficeButtonSize,
  type OfficeVariant,
} from "./types"

export interface OfficeButtonProps {
  /** Visual style. Defaults to the solid neutral button. */
  variant?: OfficeVariant
  /** Control height or square icon size. */
  size?: OfficeButtonSize
  /** Disables pointer and keyboard activation. */
  disabled?: boolean
  /** Shows a centered spinner, keeps the label width and sets `aria-busy`. */
  loading?: boolean
  /** Render the slot's single root element instead of a `<button>`. */
  asChild?: boolean
  /** Native button type; ignored when `asChild` renders a non-button. */
  type?: "button" | "submit" | "reset"
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeButtonProps>(), {
  variant: "default",
  size: "default",
  disabled: false,
  loading: false,
  asChild: false,
  type: "button",
  class: "",
})

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-button",
    `office-button--${props.variant}`,
    `office-button--size-${props.size}`,
    props.class,
  ),
)

const spinnerSize = "default" as const
</script>

<template>
  <Primitive
    v-if="props.asChild"
    as-child
    :class="classes"
    :data-disabled="props.disabled || props.loading ? 'true' : undefined"
    :data-loading="props.loading ? 'true' : undefined"
    :aria-busy="props.loading || undefined"
    :aria-disabled="props.disabled || props.loading ? 'true' : undefined"
    data-slot="button"
  >
    <slot />
  </Primitive>
  <Primitive
    v-else
    as="button"
    :class="classes"
    :type="props.type"
    :disabled="props.disabled || props.loading"
    :data-disabled="props.disabled || props.loading ? 'true' : undefined"
    :data-loading="props.loading ? 'true' : undefined"
    :aria-busy="props.loading || undefined"
    :aria-disabled="props.disabled || props.loading ? 'true' : undefined"
    data-slot="button"
  >
    <span v-if="props.loading" class="office-button__spinner">
      <OfficeSpinner :size="spinnerSize" :label="null" />
    </span>
    <span class="office-button__content">
      <slot />
    </span>
  </Primitive>
</template>
