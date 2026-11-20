<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  DropdownMenuContent,
  DropdownMenuPortal,
  DropdownMenuRoot,
  DropdownMenuTrigger,
} from "reka-ui"
import { useOfficePortalTarget } from "../../composables/office-portal"
import OfficeButton from "./OfficeButton.vue"
import { OfficeMenuList } from "./OfficeMenuItems"
import { useOfficeUiLabels } from "./labels"
import {
  officeClass,
  type OfficeAlign,
  type OfficeButtonSize,
  type OfficeMenuEntry,
  type OfficeSide,
  type OfficeVariant,
} from "./types"

export interface OfficeDropdownMenuProps {
  /** Menu entries, including separators, groups and submenus. */
  items: OfficeMenuEntry[]
  /** Controlled open state (`v-model:open`). */
  open?: boolean
  /** Text for the default trigger button when no `trigger` slot is given. */
  label?: string
  /** Variant/size of the generated trigger button. */
  triggerVariant?: OfficeVariant
  triggerSize?: OfficeButtonSize
  side?: OfficeSide
  align?: OfficeAlign
  sideOffset?: number
  alignOffset?: number
  disabled?: boolean
  /**
   * Modal menus trap focus and block outside pointer events. Pass `false` for
   * context-style menus that must let the underlying surface react.
   */
  modal?: boolean
  /** `data-testid` for the portalled menu content. */
  contentTestId?: string
  /** Extra `data-*` attributes for the menu content (surface contracts). */
  contentData?: Record<string, string | number | boolean | undefined>
  /** Accessible name for the menu; defaults to the trigger label. */
  contentAriaLabel?: string
  /**
   * Element that receives focus after the menu closes. Returning `null`
   * leaves Reka's default (the trigger) in place.
   */
  restoreFocus?: () => HTMLElement | null | undefined
  /** Extra classes merged on the trigger. */
  class?: string
  /** Extra classes merged on the portalled content. */
  contentClass?: string
}

const props = withDefaults(defineProps<OfficeDropdownMenuProps>(), {
  open: undefined,
  label: undefined,
  triggerVariant: "outline",
  triggerSize: "default",
  side: "bottom",
  align: "start",
  sideOffset: 4,
  alignOffset: 0,
  disabled: false,
  modal: true,
  contentTestId: undefined,
  contentData: undefined,
  contentAriaLabel: undefined,
  restoreFocus: undefined,
  class: "",
  contentClass: "",
})

const emit = defineEmits<{
  select: [entry: OfficeMenuEntry, event: Event]
  "update:open": [value: boolean]
}>()

const portalTarget = useOfficePortalTarget()

const labels = useOfficeUiLabels()

const contentClasses = computed(() =>
  officeClass("office-ui", "office-menu__content", props.contentClass),
)

function handleSelect(entry: OfficeMenuEntry, event: Event) {
  const callback = (entry as { onSelect?: (item: OfficeMenuEntry) => void }).onSelect
  callback?.(entry)
  emit("select", entry, event)
}

/**
 * Reka returns focus to the trigger on close. Hosts that open a menu from a
 * surface (grid, canvas) pass `restoreFocus` so the keyboard goes back there.
 */
function onCloseAutoFocus(event: Event) {
  const target = props.restoreFocus?.() ?? null
  if (!target) return
  event.preventDefault()
  target.focus()
}
</script>

<template>
  <DropdownMenuRoot
    :open="props.open"
    :modal="props.modal"
    :disabled="props.disabled"
    @update:open="emit('update:open', $event)"
  >
    <DropdownMenuTrigger
      v-if="$slots.trigger"
      as-child
      :class="props.class"
      data-slot="dropdown-menu-trigger"
    >
      <slot name="trigger" />
    </DropdownMenuTrigger>
    <DropdownMenuTrigger
      v-else
      as-child
      :class="props.class"
      data-slot="dropdown-menu-trigger"
    >
      <OfficeButton
        :variant="props.triggerVariant"
        :size="props.triggerSize"
        :disabled="props.disabled"
      >
        {{ props.label ?? labels.more }}
      </OfficeButton>
    </DropdownMenuTrigger>
    <DropdownMenuPortal :to="portalTarget ?? undefined">
      <DropdownMenuContent
        :class="contentClasses"
        :side="props.side"
        :align="props.align"
        :side-offset="props.sideOffset"
        :align-offset="props.alignOffset"
        :data-testid="props.contentTestId"
        :aria-label="props.contentAriaLabel"
        v-bind="props.contentData"
        data-slot="dropdown-menu-content"
        @close-auto-focus="onCloseAutoFocus"
      >
        <OfficeMenuList
          kind="dropdown"
          :items="props.items"
          :on-select="handleSelect"
        />
      </DropdownMenuContent>
    </DropdownMenuPortal>
  </DropdownMenuRoot>
</template>
