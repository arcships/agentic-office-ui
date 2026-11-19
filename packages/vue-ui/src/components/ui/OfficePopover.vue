<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  PopoverContent,
  PopoverPortal,
  PopoverRoot,
  PopoverTrigger,
} from "reka-ui"
import { useOfficePortalTarget } from "../../composables/office-portal"
import OfficeButton from "./OfficeButton.vue"
import { useOfficeUiLabels } from "./labels"
import {
  officeClass,
  type OfficeAlign,
  type OfficeButtonSize,
  type OfficeSide,
  type OfficeVariant,
} from "./types"

export interface OfficePopoverProps {
  /** Controlled open state (`v-model:open`). */
  open?: boolean
  /** Text for the generated trigger button when no `trigger` slot is given. */
  label?: string
  triggerVariant?: OfficeVariant
  triggerSize?: OfficeButtonSize
  side?: OfficeSide
  align?: OfficeAlign
  sideOffset?: number
  alignOffset?: number
  modal?: boolean
  /** Extra classes merged on the trigger. */
  class?: string
  /** Extra classes merged on the portalled surface. */
  contentClass?: string
}

const props = withDefaults(defineProps<OfficePopoverProps>(), {
  open: undefined,
  label: undefined,
  triggerVariant: "outline",
  triggerSize: "default",
  side: "bottom",
  align: "center",
  sideOffset: 4,
  alignOffset: 0,
  modal: false,
  class: "",
  contentClass: "",
})

const emit = defineEmits<{
  "update:open": [value: boolean]
}>()

const portalTarget = useOfficePortalTarget()

const labels = useOfficeUiLabels()
const triggerText = computed(() => props.label ?? labels.value.more)
const contentClasses = computed(() =>
  officeClass("office-ui", "office-popover__content", props.contentClass),
)
</script>

<template>
  <PopoverRoot :open="props.open" @update:open="emit('update:open', $event)">
    <PopoverTrigger
      v-if="$slots.trigger"
      as-child
      :class="props.class"
      data-slot="popover-trigger"
    >
      <slot name="trigger" />
    </PopoverTrigger>
    <PopoverTrigger
      v-else
      as-child
      :class="props.class"
      data-slot="popover-trigger"
    >
      <OfficeButton :variant="props.triggerVariant" :size="props.triggerSize">
        {{ triggerText }}
      </OfficeButton>
    </PopoverTrigger>
    <PopoverPortal :to="portalTarget ?? undefined">
      <PopoverContent
        :class="contentClasses"
        :side="props.side"
        :align="props.align"
        :side-offset="props.sideOffset"
        :align-offset="props.alignOffset"
        :modal="props.modal"
        data-slot="popover-content"
      >
        <slot />
      </PopoverContent>
    </PopoverPortal>
  </PopoverRoot>
</template>
