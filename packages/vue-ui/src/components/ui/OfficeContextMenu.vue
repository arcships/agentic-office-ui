<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  ContextMenuContent,
  ContextMenuPortal,
  ContextMenuRoot,
  ContextMenuTrigger,
} from "reka-ui"
import { useOfficePortalTarget } from "../../composables/office-portal"
import { OfficeMenuList } from "./OfficeMenuItems"
import {
  officeClass,
  type OfficeAlign,
  type OfficeMenuEntry,
  type OfficeSide,
} from "./types"

export interface OfficeContextMenuProps {
  /** Menu entries, including separators, groups and submenus. */
  items: OfficeMenuEntry[]
  /** Controlled open state (`v-model:open`). */
  open?: boolean
  side?: OfficeSide
  align?: OfficeAlign
  sideOffset?: number
  alignOffset?: number
  disabled?: boolean
  /** Extra classes merged on the trigger area. */
  class?: string
  /** Extra classes merged on the portalled content. */
  contentClass?: string
}

const props = withDefaults(defineProps<OfficeContextMenuProps>(), {
  open: undefined,
  side: "bottom",
  align: "start",
  sideOffset: 4,
  alignOffset: 0,
  disabled: false,
  class: "",
  contentClass: "",
})

const emit = defineEmits<{
  select: [entry: OfficeMenuEntry, event: Event]
  "update:open": [value: boolean]
}>()

const portalTarget = useOfficePortalTarget()

const contentClasses = computed(() =>
  officeClass("office-ui", "office-menu__content", props.contentClass),
)

function handleSelect(entry: OfficeMenuEntry, event: Event) {
  const callback = (entry as { onSelect?: (item: OfficeMenuEntry) => void }).onSelect
  callback?.(entry)
  emit("select", entry, event)
}
</script>

<template>
  <ContextMenuRoot
    :open="props.open"
    :disabled="props.disabled"
    @update:open="emit('update:open', $event)"
  >
    <ContextMenuTrigger
      as-child
      :class="props.class"
      data-slot="context-menu-trigger"
    >
      <slot />
    </ContextMenuTrigger>
    <ContextMenuPortal :to="portalTarget ?? undefined">
      <ContextMenuContent
        :class="contentClasses"
        :side="props.side"
        :align="props.align"
        :side-offset="props.sideOffset"
        :align-offset="props.alignOffset"
        data-slot="context-menu-content"
      >
        <OfficeMenuList
          kind="context"
          :items="props.items"
          :on-select="handleSelect"
        />
      </ContextMenuContent>
    </ContextMenuPortal>
  </ContextMenuRoot>
</template>
