<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import { Primitive } from "reka-ui"
import {
  officeClass,
  type OfficeControlSize,
  type OfficeVariant,
} from "./types"

export interface OfficeBadgeProps {
  variant?: OfficeVariant
  size?: OfficeControlSize
  /** Render the slot's single root element instead of a `<span>`. */
  asChild?: boolean
  class?: string
}

const props = withDefaults(defineProps<OfficeBadgeProps>(), {
  variant: "default",
  size: "default",
  asChild: false,
  class: "",
})

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-badge",
    `office-badge--${props.variant}`,
    props.size !== "default" && `office-badge--${props.size}`,
    props.class,
  ),
)
</script>

<template>
  <Primitive
    as="span"
    :as-child="props.asChild"
    :class="classes"
    data-slot="badge"
  >
    <slot />
  </Primitive>
</template>
