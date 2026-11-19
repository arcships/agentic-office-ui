<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import OfficeButton from "./OfficeButton.vue"
import OfficeIcon from "../OfficeIcon.vue"
import type { OfficeIconName } from "../../icons"
import {
  officeClass,
  type OfficeButtonSize,
  type OfficeVariant,
} from "./types"

export interface OfficeIconButtonProps {
  /** Accessible name. Required: the button has no visible text. */
  label: string
  /** Registry icon rendered inside the button. */
  icon?: OfficeIconName
  /** Native tooltip text; defaults to `label`. */
  title?: string
  /** Visual style. Toolbar icon buttons use `ghost`. */
  variant?: OfficeVariant
  /** Square control size. Defaults to 32px (`icon`). */
  size?: OfficeButtonSize
  /** Pressed/active toolbar state. */
  active?: boolean
  disabled?: boolean
  loading?: boolean
  asChild?: boolean
  class?: string
}

const props = withDefaults(defineProps<OfficeIconButtonProps>(), {
  icon: undefined,
  title: undefined,
  variant: "ghost",
  size: "icon",
  active: false,
  disabled: false,
  loading: false,
  asChild: false,
  class: "",
})

const classes = computed(() =>
  officeClass("office-icon-button", props.class),
)
</script>

<template>
  <OfficeButton
    :class="classes"
    :variant="props.variant"
    :size="props.size"
    :disabled="props.disabled"
    :loading="props.loading"
    :as-child="props.asChild"
    :data-active="props.active ? 'true' : undefined"
    :aria-label="props.label"
    :aria-pressed="props.active ? 'true' : undefined"
    :title="props.title ?? props.label"
    data-slot="icon-button"
  >
    <slot>
      <OfficeIcon v-if="props.icon" :name="props.icon" />
    </slot>
  </OfficeButton>
</template>
