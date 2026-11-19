<script setup lang="ts">
import { computed } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import { officeClass } from "../ui/types"

export interface OfficeToolbarProps {
  /**
   * Accessible name of the toolbar. Defaults to the shared catalog's
   * `common.chrome.toolbar` ("Document toolbar").
   */
  label?: string
  /** Allow clusters to wrap onto a new row when the viewport is narrow. */
  wrap?: boolean
  /** Keep the toolbar visible while the viewer content scrolls. */
  sticky?: boolean
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeToolbarProps>(), {
  label: undefined,
  wrap: true,
  sticky: false,
  class: "",
})

const m = useOfficeUiMessages()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-toolbar",
    !props.wrap && "office-toolbar--nowrap",
    props.sticky && "office-toolbar--sticky",
    props.class,
  ),
)
</script>

<template>
  <div
    :class="classes"
    role="toolbar"
    :aria-label="props.label ?? m.common.chrome.toolbar"
    data-slot="toolbar"
    data-testid="office-toolbar"
  >
    <div class="office-toolbar__leading">
      <slot name="leading">
        <slot />
      </slot>
    </div>
    <div v-if="$slots.center" class="office-toolbar__center">
      <slot name="center" />
    </div>
    <div v-if="$slots.trailing" class="office-toolbar__trailing">
      <slot name="trailing" />
    </div>
  </div>
</template>
