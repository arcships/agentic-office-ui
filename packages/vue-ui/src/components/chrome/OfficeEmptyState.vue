<script setup lang="ts">
import { computed, useId } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIcon from "../OfficeIcon.vue"
import { officeClass } from "../ui/types"
import type { OfficeIconName } from "../../icons"

export interface OfficeEmptyStateProps {
  title?: string
  description?: string
  /**
   * Optional glyph. The upstream `PdfEditorEmptyState` renders no glyph unless
   * the host passes one, so this is icon-free by default; pass an
   * `OfficeIconName` (or the `icon` slot) to opt in.
   */
  icon?: OfficeIconName
  /** Tighter padding. */
  compact?: boolean
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeEmptyStateProps>(), {
  title: undefined,
  description: undefined,
  icon: undefined,
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
    :aria-labelledby="titleText ? titleId : undefined"
    :aria-describedby="descriptionText || $slots.default ? descriptionId : undefined"
    data-slot="empty-state"
    data-testid="office-empty-state"
  >
    <div
      v-if="props.icon || $slots.icon"
      class="office-state__icon"
      data-slot="state-icon"
    >
      <slot name="icon">
        <OfficeIcon :name="props.icon" />
      </slot>
    </div>
    <p v-if="titleText" :id="titleId" class="office-state__title" data-slot="state-title">
      {{ titleText }}
    </p>
    <div
      v-if="descriptionText || $slots.default"
      :id="descriptionId"
      class="office-state__description"
      data-slot="state-description"
    >
      <slot>{{ descriptionText }}</slot>
    </div>
    <div v-if="$slots.actions" class="office-state__actions" data-slot="state-actions">
      <slot name="actions" />
    </div>
  </div>
</template>
