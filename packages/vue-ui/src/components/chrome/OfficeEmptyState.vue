<script setup lang="ts">
import { computed, useId } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIcon from "../OfficeIcon.vue"
import { officeClass } from "../ui/types"
import type { OfficeIconName } from "../../icons"

export interface OfficeEmptyStateProps {
  title?: string
  description?: string
  icon?: OfficeIconName
  /** Tighter padding and no minimum height. */
  compact?: boolean
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeEmptyStateProps>(), {
  title: undefined,
  description: undefined,
  icon: "folder-open",
  compact: false,
  class: "",
})

const m = useOfficeUiMessages()
const titleId = useId()
const descriptionId = useId()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-state",
    "office-state--empty",
    props.compact && "office-state--compact",
    props.class,
  ),
)
const titleText = computed(() => props.title ?? m.value.common.chrome.emptyTitle)
const descriptionText = computed(
  () => props.description ?? m.value.common.chrome.emptyDescription,
)
</script>

<template>
  <div
    :class="classes"
    role="group"
    :aria-labelledby="titleId"
    :aria-describedby="descriptionId"
    data-slot="empty-state"
    data-testid="office-empty-state"
  >
    <div class="office-state__icon" data-slot="state-icon">
      <slot name="icon">
        <OfficeIcon :name="props.icon" />
      </slot>
    </div>
    <p :id="titleId" class="office-state__title" data-slot="state-title">
      {{ titleText }}
    </p>
    <div :id="descriptionId" class="office-state__description" data-slot="state-description">
      <slot>{{ descriptionText }}</slot>
    </div>
    <div v-if="$slots.actions" class="office-state__actions" data-slot="state-actions">
      <slot name="actions" />
    </div>
  </div>
</template>
