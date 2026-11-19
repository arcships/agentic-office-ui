<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import { Primitive } from "reka-ui"
import { officeClass } from "./types"

export interface OfficeSkeletonProps {
  /** Render the slot's single root element instead of a `<div>`. */
  asChild?: boolean
  class?: string
}

const props = withDefaults(defineProps<OfficeSkeletonProps>(), {
  asChild: false,
  class: "",
})

const classes = computed(() =>
  officeClass("office-ui", "office-skeleton", props.class),
)
</script>

<template>
  <Primitive
    as="div"
    :as-child="props.asChild"
    :class="classes"
    aria-hidden="true"
    data-slot="skeleton"
  >
    <slot />
  </Primitive>
</template>
