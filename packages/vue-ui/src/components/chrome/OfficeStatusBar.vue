<script setup lang="ts">
import { computed } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIcon from "../OfficeIcon.vue"
import OfficeSpinner from "../ui/OfficeSpinner.vue"
import { officeClass } from "../ui/types"

export interface OfficeStatusBarProps {
  /** Current page, 1-based. Rendered with `totalPages` when both are given. */
  page?: number
  totalPages?: number
  /** Short status text, for example "Saved" or "Rendering page 3". */
  status?: string
  fileName?: string
  /** Error text; rendered in the destructive colour with `role="alert"`. */
  error?: string
  /** Warning text; rendered in the warning colour as a status message. */
  warning?: string
  /** Show the loading spinner and default the status to `common.loading`. */
  busy?: boolean
  /**
   * Accessible name of the bar. Defaults to `common.chrome.statusBar`
   * ("Document status").
   */
  label?: string
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeStatusBarProps>(), {
  page: undefined,
  totalPages: undefined,
  status: undefined,
  fileName: undefined,
  error: undefined,
  warning: undefined,
  busy: false,
  label: undefined,
  class: "",
})

const m = useOfficeUiMessages()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-status-bar",
    (props.error || props.warning) && "office-status-bar--attention",
    props.class,
  ),
)

const pageLabel = computed(() => {
  if (typeof props.page !== "number" || typeof props.totalPages !== "number") {
    return ""
  }
  return formatOfficeMessage(m.value.common.page.indicator, {
    current: Math.min(Math.max(props.page, 1), Math.max(props.totalPages, 1)),
    total: props.totalPages,
  })
})

const statusText = computed(() => {
  if (props.status) return props.status
  if (props.busy) return m.value.common.loading
  return ""
})
const warningText = computed(
  () => `${m.value.common.chrome.warning}: ${props.warning ?? ""}`,
)
const errorText = computed(
  () => `${m.value.common.chrome.error}: ${props.error ?? ""}`,
)
</script>

<template>
  <footer
    :class="classes"
    :aria-label="props.label ?? m.common.chrome.statusBar"
    data-slot="status-bar"
    data-testid="office-status-bar"
  >
    <div class="office-status-bar__group office-status-bar__start">
      <slot name="start">
        <span
          v-if="pageLabel"
          class="office-status-bar__page"
          data-slot="status-page"
          data-testid="office-status-bar-page"
        >{{ pageLabel }}</span>
        <span
          v-if="statusText"
          class="office-status-bar__status"
          role="status"
          aria-live="polite"
          data-slot="status-text"
          data-testid="office-status-bar-status"
        >
          <OfficeSpinner v-if="props.busy" size="xs" :label="null" />
          <span>{{ statusText }}</span>
        </span>
        <span
          v-if="props.warning"
          class="office-status-bar__message office-status-bar__message--warning"
          role="status"
          :aria-label="warningText"
          data-slot="status-warning"
          data-testid="office-status-bar-warning"
        >
          <OfficeIcon name="triangle-alert" />
          <span class="office-status-bar__message-label">{{ m.common.chrome.warning }}</span>
          <span>{{ props.warning }}</span>
        </span>
        <span
          v-if="props.error"
          class="office-status-bar__message office-status-bar__message--error"
          role="alert"
          :aria-label="errorText"
          data-slot="status-error"
          data-testid="office-status-bar-error"
        >
          <OfficeIcon name="circle-x" />
          <span class="office-status-bar__message-label">{{ m.common.chrome.error }}</span>
          <span>{{ props.error }}</span>
        </span>
      </slot>
    </div>
    <div class="office-status-bar__group office-status-bar__end">
      <slot name="end">
        <span
          v-if="props.fileName"
          class="office-status-bar__file"
          :title="props.fileName"
          data-slot="status-file"
          data-testid="office-status-bar-file"
        >
          <span class="office-status-bar__file-label">{{ m.common.chrome.fileName }}</span>
          <span class="office-status-bar__file-name">{{ props.fileName }}</span>
        </span>
      </slot>
    </div>
  </footer>
</template>
