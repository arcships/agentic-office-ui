<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  TabsContent,
  TabsIndicator,
  TabsList,
  TabsRoot,
  TabsTrigger,
} from "reka-ui"
import OfficeIcon from "../OfficeIcon.vue"
import {
  officeClass,
  type OfficeOrientation,
  type OfficeTab,
} from "./types"

export interface OfficeTabsProps {
  /** Selected tab value (`v-model`). */
  modelValue?: string
  /** Tab descriptors. Content is provided in the `panel-<value>` slot. */
  tabs: OfficeTab[]
  /** `default` is the segmented control; `underline` is a text tab bar. */
  variant?: "default" | "underline"
  orientation?: OfficeOrientation
  /** Extra classes merged on the root. */
  class?: string
  /** Extra classes merged on every panel. */
  contentClass?: string
}

const props = withDefaults(defineProps<OfficeTabsProps>(), {
  modelValue: undefined,
  variant: "default",
  orientation: "horizontal",
  class: "",
  contentClass: "",
})

const emit = defineEmits<{
  "update:modelValue": [value: string]
}>()

const listClasses = computed(() =>
  officeClass(
    "office-tabs__list",
    props.variant === "underline" && "office-tabs__list--underline",
  ),
)
</script>

<template>
  <TabsRoot
    :model-value="props.modelValue"
    :orientation="props.orientation"
    :class="officeClass('office-ui', 'office-tabs', props.class)"
    data-slot="tabs"
    @update:model-value="emit('update:modelValue', $event)"
  >
    <TabsList :class="listClasses" data-slot="tabs-list">
      <TabsTrigger
        v-for="tab in props.tabs"
        :key="tab.value"
        :value="tab.value"
        :disabled="tab.disabled"
        class="office-tabs__trigger"
        data-slot="tabs-trigger"
      >
        <OfficeIcon v-if="tab.icon" :name="tab.icon" />
        {{ tab.label }}
      </TabsTrigger>
      <TabsIndicator class="office-tabs__indicator" data-slot="tabs-indicator" />
    </TabsList>
    <TabsContent
      v-for="tab in props.tabs"
      :key="tab.value"
      :value="tab.value"
      :class="officeClass('office-tabs__content', props.contentClass)"
      data-slot="tabs-content"
    >
      <slot :name="`panel-${tab.value}`" />
    </TabsContent>
  </TabsRoot>
</template>
