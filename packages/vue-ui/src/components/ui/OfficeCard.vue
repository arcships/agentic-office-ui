<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import { officeClass } from "./types"

export interface OfficeCardProps {
  /** Optional heading rendered in the card header. */
  title?: string
  /** Optional supporting copy under the heading. */
  description?: string
  /** Heading element level for the generated title. */
  headingLevel?: 2 | 3 | 4 | 5 | 6
  class?: string
}

const props = withDefaults(defineProps<OfficeCardProps>(), {
  title: undefined,
  description: undefined,
  headingLevel: 3,
  class: "",
})

const classes = computed(() =>
  officeClass("office-ui", "office-card", props.class),
)
const headingTag = computed(() => `h${props.headingLevel}`)
</script>

<template>
  <div :class="classes" data-slot="card">
    <div
      v-if="$slots.header || props.title || props.description"
      class="office-card__header"
      data-slot="card-header"
    >
      <slot name="header">
        <component :is="headingTag" v-if="props.title" class="office-card__title">
          {{ props.title }}
        </component>
        <p v-if="props.description" class="office-card__description">
          {{ props.description }}
        </p>
      </slot>
    </div>
    <div class="office-card__content" data-slot="card-content">
      <slot />
    </div>
    <div v-if="$slots.footer" class="office-card__footer" data-slot="card-footer">
      <slot name="footer" />
    </div>
  </div>
</template>
