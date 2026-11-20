<script setup lang="ts">
import { computed, useId } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIcon from "../OfficeIcon.vue"
import OfficeButton from "../ui/OfficeButton.vue"
import { officeClass } from "../ui/types"
import type { OfficeIconName } from "../../icons"

export interface OfficeErrorStateProps {
  title?: string
  /** Explanation shown below the title. Falls back to `error`. */
  description?: string
  /** Raw error message from the host; used as the description fallback. */
  error?: string
  /**
   * Optional glyph above the title. The upstream error card is icon-free, so
   * this is off by default; pass an `OfficeIconName` (or the `icon` slot) to
   * opt in.
   */
  icon?: OfficeIconName
  /** Render the built-in retry button. */
  showRetry?: boolean
  retryLabel?: string
  compact?: boolean
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeErrorStateProps>(), {
  title: undefined,
  description: undefined,
  error: undefined,
  icon: undefined,
  showRetry: true,
  retryLabel: undefined,
  compact: false,
  class: "",
})

const emit = defineEmits<{
  retry: []
}>()

const m = useOfficeUiMessages()
const titleId = useId()
const descriptionId = useId()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-state",
    "office-state--error",
    props.compact && "office-state--compact",
    props.class,
  ),
)
const titleText = computed(() => props.title ?? m.value.common.chrome.errorTitle)
const descriptionText = computed(
  () =>
    props.description
    ?? props.error
    ?? m.value.common.chrome.errorDescription,
)
</script>

<template>
  <div
    :class="classes"
    role="alert"
    :aria-labelledby="titleText ? titleId : undefined"
    :aria-describedby="descriptionText ? descriptionId : undefined"
    data-slot="error-state"
    data-testid="office-error-state"
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
    <div
      v-if="$slots.actions || props.showRetry"
      class="office-state__actions"
      data-slot="state-actions"
    >
      <slot name="actions">
        <OfficeButton
          v-if="props.showRetry"
          variant="outline"
          size="sm"
          data-testid="office-error-retry"
          @click="emit('retry')"
        >{{ props.retryLabel ?? m.common.chrome.retry }}</OfficeButton>
      </slot>
    </div>
  </div>
</template>
