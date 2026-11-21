<script setup lang="ts">
import OfficeTooltip from "./ui/OfficeTooltip.vue"
import type { OfficeAlign, OfficeSide } from "./ui/types"

/**
 * Backward-compatible Tooltip.
 *
 * The public props are unchanged (`content`, `side`, `align`, `delayMs`,
 * `className`); the implementation now delegates to `OfficeTooltip`, so
 * focus/hover behaviour, collision handling and ARIA come from Reka UI.
 */
export interface TooltipProps {
  content?: string
  side?: OfficeSide
  align?: OfficeAlign
  delayMs?: number
  className?: string
}

const props = withDefaults(defineProps<TooltipProps>(), {
  content: "",
  side: "top",
  align: "center",
  delayMs: 300,
  className: "",
})
</script>

<template>
  <OfficeTooltip
    :content="props.content"
    :side="props.side"
    :align="props.align"
    :delay-duration="props.delayMs"
    :content-class="props.className"
  >
    <span class="tooltip-trigger">
      <slot />
    </span>
  </OfficeTooltip>
</template>

<style scoped>
.tooltip-trigger {
  display: inline-flex;
}
</style>
