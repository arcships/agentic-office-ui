<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, nextTick, useAttrs } from "vue"
import {
  PopoverAnchor,
  PopoverContent,
  PopoverPortal,
  PopoverRoot,
  PopoverTrigger,
  type ReferenceElement,
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
  /** `data-testid` for the portalled surface. */
  contentTestId?: string
  /** Extra attributes for the portalled surface (surface contracts). */
  contentData?: Record<string, string | number | boolean | undefined>
  /**
   * Virtual reference used instead of a trigger element. Pass a zero-size
   * `getBoundingClientRect()` (for example viewport link coordinates) and the
   * surface anchors there without rendering a trigger button.
   */
  virtualRef?: ReferenceElement
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
  contentTestId: undefined,
  contentData: undefined,
  virtualRef: undefined,
})

// Unknown attributes and native listeners belong to the portalled surface:
// the root is a fragment (PopoverRoot), so fallthrough would otherwise drop
// `data-*`, `aria-*`, `role` and pointer handlers. Existing consumers that only
// pass the declared props are unaffected.
defineOptions({ inheritAttrs: false })

const emit = defineEmits<{
  "update:open": [value: boolean]
}>()

const portalTarget = useOfficePortalTarget()
const attrs = useAttrs()

const labels = useOfficeUiLabels()
const triggerText = computed(() => props.label ?? labels.value.more)
const contentClasses = computed(() =>
  officeClass("office-ui", "office-popover__content", props.contentClass),
)
// One merged binding: the template compiler rejects two `v-bind="..."`
// attributes on the same element. Surface data wins over the base record so a
// host can override a contract attribute.
const contentAttributes = computed(() => ({
  ...props.contentData,
  ...attrs,
}))

/**
 * With a virtual anchor there is no trigger element, so Reka's generated
 * `aria-labelledby` resolves to the empty string and would mask the caller's
 * `aria-label`. Drop the empty reference once the surface is in the DOM; the
 * second pass covers Teleport insertion landing after the vnode hook.
 */
function clearEmptyLabelledby(): void {
  if (typeof document === "undefined") return
  const root = portalTarget?.value ?? document.body
  for (const surface of root.querySelectorAll(
    '[data-slot="popover-content"][aria-labelledby=""]',
  )) {
    if (surface.hasAttribute("aria-label")) surface.removeAttribute("aria-labelledby")
  }
}

function onContentMounted(): void {
  clearEmptyLabelledby()
  void nextTick(clearEmptyLabelledby)
}
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
      v-else-if="!props.virtualRef && !$slots.anchor"
      as-child
      :class="props.class"
      data-slot="popover-trigger"
    >
      <OfficeButton :variant="props.triggerVariant" :size="props.triggerSize">
        {{ triggerText }}
      </OfficeButton>
    </PopoverTrigger>
    <PopoverAnchor
      v-if="props.virtualRef"
      :reference="props.virtualRef"
      class="office-popover__anchor"
      data-slot="popover-anchor"
      aria-hidden="true"
    />
    <PopoverAnchor
      v-else-if="$slots.anchor"
      as-child
      data-slot="popover-anchor"
    >
      <slot name="anchor" />
    </PopoverAnchor>
    <PopoverPortal :to="portalTarget ?? undefined">
      <PopoverContent
        :class="contentClasses"
        :data-testid="props.contentTestId"
        v-bind="contentAttributes"
        :side="props.side"
        :align="props.align"
        :side-offset="props.sideOffset"
        :align-offset="props.alignOffset"
        :modal="props.modal"
        data-slot="popover-content"
        @vue:mounted="onContentMounted"
      >
        <slot />
      </PopoverContent>
    </PopoverPortal>
  </PopoverRoot>
</template>
