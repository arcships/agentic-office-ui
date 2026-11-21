<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  ToastAction,
  ToastClose,
  ToastDescription,
  ToastPortal,
  ToastRoot,
  ToastTitle,
  ToastViewport,
  useToastManager,
  type ToastObject,
} from "reka-ui"
import { useOfficePortalTarget } from "../../composables/office-portal"
import OfficeIcon from "../OfficeIcon.vue"
import OfficeSpinner from "./OfficeSpinner.vue"
import type { OfficeIconName } from "../../icons"

export interface OfficeToastViewportProps {
  label: string
  closeLabel: string
  position?: "top-left" | "top-right" | "bottom-left" | "bottom-right"
}

const props = withDefaults(defineProps<OfficeToastViewportProps>(), {
  position: "bottom-right",
})

const { toasts } = useToastManager()

const portalTarget = useOfficePortalTarget()

const viewportClasses = computed(
  () =>
    `office-ui office-toast__viewport office-toast__viewport--${props.position}`,
)

const STATUS_ICONS: Record<string, OfficeIconName> = {
  success: "circle-check",
  error: "circle-x",
  warning: "triangle-alert",
  info: "info",
}

function statusIcon(toast: ToastObject): OfficeIconName {
  return STATUS_ICONS[String(toast.status)] ?? "info"
}
</script>

<template>
  <ToastPortal :to="portalTarget ?? undefined">
    <ToastViewport
      :class="viewportClasses"
      :label="props.label"
      data-slot="toast-viewport"
    >
      <ToastRoot
        v-for="toast in toasts"
        :key="toast.id"
        :toast="toast"
        class="office-toast"
        data-slot="toast"
      >
        <span class="office-toast__icon">
          <OfficeSpinner
            v-if="toast.status === 'loading'"
            size="sm"
            :label="null"
          />
          <OfficeIcon v-else :name="statusIcon(toast)" />
        </span>
        <div class="office-toast__body">
          <ToastTitle
            v-if="toast.title"
            class="office-toast__title"
            data-slot="toast-title"
          />
          <ToastDescription
            v-if="toast.description"
            class="office-toast__description"
            data-slot="toast-description"
          />
        </div>
        <ToastAction
          v-if="toast.actionProps"
          class="office-ui office-button office-button--outline office-button--size-sm office-toast__action"
          data-slot="toast-action"
        />
        <ToastClose
          class="office-toast__close"
          :aria-label="props.closeLabel"
          data-slot="toast-close"
        >
          <OfficeIcon name="x" />
        </ToastClose>
      </ToastRoot>
    </ToastViewport>
  </ToastPortal>
</template>
