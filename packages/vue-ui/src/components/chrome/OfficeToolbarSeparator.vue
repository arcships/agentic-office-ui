<script setup lang="ts">
import { computed } from "vue"
import OfficeSeparator from "../ui/OfficeSeparator.vue"
import { officeClass } from "../ui/types"

export interface OfficeToolbarSeparatorProps {
  /** Extra classes merged on the separator. */
  class?: string
}

const props = withDefaults(defineProps<OfficeToolbarSeparatorProps>(), {
  class: "",
})

const classes = computed(() =>
  officeClass("office-toolbar__separator", props.class),
)
</script>

<template>
  <OfficeSeparator
    orientation="vertical"
    :class="classes"
    data-slot="toolbar-separator"
    data-testid="office-toolbar-separator"
  />
</template>
