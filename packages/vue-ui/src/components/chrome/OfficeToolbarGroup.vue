<script setup lang="ts">
import { computed } from "vue"
import { officeClass } from "../ui/types"
import type { OfficeToolbarAlign, OfficeToolbarGap } from "./types"

export interface OfficeToolbarGroupProps {
  /** Accessible name when the controls form a named group. */
  label?: string
  /** Gap between controls. `sm` is 4px, `md` is 8px. */
  gap?: OfficeToolbarGap
  /** Cross-axis alignment of the controls. */
  align?: OfficeToolbarAlign
  /** Let the group absorb the remaining toolbar width. */
  grow?: boolean
  /** Allow the controls inside the group to wrap. */
  wrap?: boolean
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeToolbarGroupProps>(), {
  label: undefined,
  gap: "sm",
  align: "center",
  grow: false,
  wrap: true,
  class: "",
})

const classes = computed(() =>
  officeClass(
    "office-toolbar__group",
    `office-toolbar__group--gap-${props.gap}`,
    `office-toolbar__group--align-${props.align}`,
    props.grow && "office-toolbar__group--grow",
    !props.wrap && "office-toolbar__group--nowrap",
    props.class,
  ),
)
</script>

<template>
  <div
    :class="classes"
    role="group"
    :aria-label="props.label"
    data-slot="toolbar-group"
    data-testid="office-toolbar-group"
  >
    <slot />
  </div>
</template>
