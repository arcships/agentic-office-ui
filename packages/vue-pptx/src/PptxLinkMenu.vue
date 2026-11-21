<template>
  <OfficeDropdownMenu
    :open="open"
    :items="items"
    :content-test-id="'pptx-context-menu'"
    :menu-label="label"
    :restore-focus="restoreFocus"
    :side-offset="4"
    content-class="pptx-viewer__link-menu"
    @update:open="emit('update:open', $event)"
    @select="onSelect"
  >
    <template #trigger>
      <!--
        A zero-size anchor at the pointer. Reka positions the dropdown against
        it, so the menu keeps the right-click coordinates while focus, typeahead,
        arrow-key navigation and Escape stay owned by the shared menu primitive.
      -->
      <span
        class="pptx-viewer__link-menu-anchor"
        :style="anchorStyle"
        data-testid="pptx-context-menu-anchor"
        aria-hidden="true"
        tabindex="-1"
      />
    </template>
  </OfficeDropdownMenu>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue"
import { OfficeDropdownMenu } from "@arcships/vue-ui"
import type { OfficeMenuEntry } from "@arcships/vue-ui"

const props = withDefaults(defineProps<{
  /** Controlled menu visibility. */
  open: boolean
  /** Viewport coordinates of the right-click. */
  x: number
  y: number
  /** Entries built by the viewer (open link / copy link). */
  items: OfficeMenuEntry[]
  /** Accessible name of the menu surface. */
  label: string
  /** Element that receives focus when the menu closes (the viewer root). */
  restoreFocus?: (() => HTMLElement | null | undefined) | undefined
}>(), {
  x: 0,
  y: 0,
  restoreFocus: undefined,
})

const emit = defineEmits<{
  "update:open": [open: boolean]
  /** The stable `id` of the chosen entry. */
  select: [id: string]
}>()

const anchorStyle = computed<CSSProperties>(() => ({
  insetBlockStart: `${props.y}px`,
  insetInlineStart: `${props.x}px`,
}))

function onSelect(entry: OfficeMenuEntry): void {
  const id = (entry as { id?: string }).id
  if (id) emit("select", id)
}
</script>
