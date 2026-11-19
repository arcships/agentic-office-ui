<script setup lang="ts">
import OfficeSpinner from "./ui/OfficeSpinner.vue"
import type { OfficeControlSize } from "./ui/types"

export interface SpinnerProps {
  /** Legacy class hook; merged on the spinner root. */
  className?: string
  /** `md` is kept as an alias of the primitive default size. */
  size?: OfficeControlSize | "md"
  /** Accessible name; pass `null` for a decorative spinner. */
  label?: string | null
}

const props = withDefaults(defineProps<SpinnerProps>(), {
  className: "",
  size: undefined,
  label: undefined,
})

const sizeMap: Record<string, OfficeControlSize> = {
  xs: "xs",
  sm: "sm",
  md: "default",
  default: "default",
  lg: "lg",
}
</script>

<template>
  <OfficeSpinner
    :size="props.size ? sizeMap[props.size] : 'default'"
    :label="props.label"
    :class="props.className"
  />
</template>
