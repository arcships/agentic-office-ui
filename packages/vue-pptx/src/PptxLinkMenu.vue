<template>
  <OfficeDropdownMenu
    :open="open"
    :items="items"
    side="bottom"
    align="start"
    :side-offset="4"
    :content-label="label"
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
        :style="{ left: `${x}px`, top: `${y}px` }"
        data-testid="pptx-context-menu-anchor"
        aria-hidden="true"
      />
    </template>
  </OfficeDropdownMenu>
</template>

<script setup lang="ts">
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
}>(), {
  x: 0,
  y: 0,
})

const emit = defineEmits<{
  "update:open": [open: boolean]
  /** The stable `id` of the chosen entry. */
  select: [id: string]
}>()

function onSelect(entry: OfficeMenuEntry): void {
  const id = (entry as { id?: string }).id
  if (id) emit("select", id)
}
</script>
