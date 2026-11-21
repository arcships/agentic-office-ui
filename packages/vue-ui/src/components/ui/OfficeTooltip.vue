<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  TooltipContent,
  TooltipPortal,
  TooltipProvider,
  TooltipRoot,
  TooltipTrigger,
} from "reka-ui"
import { useOfficePortalTarget } from "../../composables/office-portal"
import {
  officeClass,
  type OfficeAlign,
  type OfficeSide,
} from "./types"

export interface OfficeTooltipProps {
  /** Tooltip text. A default slot overrides it. */
  content?: string
  side?: OfficeSide
  align?: OfficeAlign
  /** Gap between trigger and bubble in pixels. */
  sideOffset?: number
  /** Hover/focus delay in milliseconds. */
  delayDuration?: number
  disabled?: boolean
  /** Extra classes merged on the trigger element. */
  class?: string
  /** Extra classes merged on the portalled bubble. */
  contentClass?: string
}

const props = withDefaults(defineProps<OfficeTooltipProps>(), {
  content: "",
  side: "top",
  align: "center",
  sideOffset: 4,
  delayDuration: 300,
  disabled: false,
  class: "",
  contentClass: "",
})

const portalTarget = useOfficePortalTarget()

const contentClasses = computed(() =>
  officeClass("office-ui", "office-tooltip__content", props.contentClass),
)
</script>

<template>
  <TooltipProvider :delay-duration="props.delayDuration">
    <TooltipRoot :disabled="props.disabled">
      <TooltipTrigger as-child :class="props.class" data-slot="tooltip-trigger">
        <slot />
      </TooltipTrigger>
      <TooltipPortal :to="portalTarget ?? undefined">
        <TooltipContent
          :class="contentClasses"
          :side="props.side"
          :align="props.align"
          :side-offset="props.sideOffset"
          data-slot="tooltip-content"
        >
          <span class="office-tooltip__viewport">
            <slot name="content">{{ props.content }}</slot>
          </span>
        </TooltipContent>
      </TooltipPortal>
    </TooltipRoot>
  </TooltipProvider>
</template>
