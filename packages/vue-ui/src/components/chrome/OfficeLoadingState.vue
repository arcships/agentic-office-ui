<script setup lang="ts">
import { computed, useId } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import OfficeSpinner from "../ui/OfficeSpinner.vue"
import { officeClass } from "../ui/types"

export interface OfficeLoadingStateProps {
  title?: string
  description?: string
  /** Deterministic progress, 0-100. Omit for an indeterminate spinner. */
  progress?: number
  /** Hide the spinner when a progress bar is shown. */
  showSpinner?: boolean
  compact?: boolean
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeLoadingStateProps>(), {
  title: undefined,
  description: undefined,
  progress: undefined,
  showSpinner: true,
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
    "office-state--loading",
    props.compact && "office-state--compact",
    props.class,
  ),
)
const titleText = computed(() => props.title ?? m.value.common.chrome.loadingTitle)
const descriptionText = computed(
  () => props.description ?? m.value.common.chrome.loadingDescription,
)
const hasProgress = computed(
  () =>
    typeof props.progress === "number"
    && Number.isFinite(props.progress),
)
const progressValue = computed(() =>
  hasProgress.value
    ? Math.min(Math.max(props.progress as number, 0), 100)
    : 0,
)
const progressStyle = computed(() => ({
  inlineSize: `${progressValue.value}%`,
}))
</script>

<template>
  <div
    :class="classes"
    role="status"
    aria-live="polite"
    aria-busy="true"
    :aria-labelledby="titleId"
    :aria-describedby="descriptionId"
    data-slot="loading-state"
    data-testid="office-loading-state"
  >
    <div class="office-state__icon" data-slot="state-icon">
      <slot name="icon">
        <OfficeSpinner v-if="props.showSpinner" size="lg" :label="null" />
      </slot>
    </div>
    <p :id="titleId" class="office-state__title" data-slot="state-title">
      {{ titleText }}
    </p>
    <div :id="descriptionId" class="office-state__description" data-slot="state-description">
      <slot>{{ descriptionText }}</slot>
    </div>
    <div
      v-if="hasProgress"
      class="office-state__progress"
      role="progressbar"
      :aria-label="titleText"
      :aria-valuemin="0"
      :aria-valuemax="100"
      :aria-valuenow="progressValue"
      data-slot="state-progress"
      data-testid="office-loading-progress"
    >
      <span class="office-state__progress-bar" :style="progressStyle" />
    </div>
    <div v-if="$slots.actions" class="office-state__actions" data-slot="state-actions">
      <slot name="actions" />
    </div>
  </div>
</template>
