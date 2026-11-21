<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import { Separator } from "reka-ui"
import { officeClass } from "./types"

export interface OfficeSeparatorProps {
  orientation?: "horizontal" | "vertical"
  /** Decorative separators are hidden from the accessibility tree. */
  decorative?: boolean
  asChild?: boolean
  class?: string
}

const props = withDefaults(defineProps<OfficeSeparatorProps>(), {
  orientation: "horizontal",
  decorative: true,
  asChild: false,
  class: "",
})

const classes = computed(() =>
  officeClass("office-ui", "office-separator", props.class),
)
</script>

<template>
  <Separator
    :orientation="props.orientation"
    :decorative="props.decorative"
    :as-child="props.asChild"
    :class="classes"
    data-slot="separator"
  />
</template>
