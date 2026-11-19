<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  DialogClose,
  DialogContent,
  DialogDescription,
  DialogOverlay,
  DialogPortal,
  DialogRoot,
  DialogTitle,
  DialogTrigger,
  VisuallyHidden,
} from "reka-ui"
import { useOfficePortalTarget } from "../../composables/office-portal"
import OfficeIcon from "../OfficeIcon.vue"
import { useOfficeUiLabels } from "./labels"
import { officeClass, type OfficeOverlaySize } from "./types"

export interface OfficeDialogProps {
  /** Controlled open state (`v-model:open`). */
  open?: boolean
  /** Dialog heading. Falls back to a localized "Dialog" label. */
  title?: string
  /** Supporting copy under the heading. */
  description?: string
  size?: OfficeOverlaySize
  /** Hide the heading visually while keeping it for screen readers. */
  titleHidden?: boolean
  showClose?: boolean
  /** Localized override for the close button. */
  closeLabel?: string
  modal?: boolean
  closeOnEscape?: boolean
  closeOnBackdrop?: boolean
  /** Extra classes merged on the dialog surface. */
  class?: string
  /** Extra classes merged on the trigger slot wrapper. */
  triggerClass?: string
}

const props = withDefaults(defineProps<OfficeDialogProps>(), {
  open: undefined,
  title: undefined,
  description: undefined,
  size: "default",
  titleHidden: false,
  showClose: true,
  closeLabel: undefined,
  modal: true,
  closeOnEscape: true,
  closeOnBackdrop: true,
  class: "",
  triggerClass: "",
})

const emit = defineEmits<{
  "update:open": [value: boolean]
}>()

const portalTarget = useOfficePortalTarget()

const labels = useOfficeUiLabels()
const resolvedCloseLabel = computed(() => props.closeLabel ?? labels.value.close)
const resolvedTitle = computed(() => props.title ?? labels.value.dialog)

const contentClasses = computed(() =>
  officeClass(
    "office-ui",
    "office-dialog__content",
    props.size !== "default" && `office-dialog__content--${props.size}`,
    props.class,
  ),
)

function onEscapeKeyDown(event: KeyboardEvent) {
  if (!props.closeOnEscape) event.preventDefault()
}

function onPointerDownOutside(event: Event) {
  if (!props.closeOnBackdrop) event.preventDefault()
}
</script>

<template>
  <DialogRoot
    :open="props.open"
    :modal="props.modal"
    @update:open="emit('update:open', $event)"
  >
    <DialogTrigger
      v-if="$slots.trigger"
      as-child
      :class="props.triggerClass"
      data-slot="dialog-trigger"
    >
      <slot name="trigger" />
    </DialogTrigger>
    <DialogPortal :to="portalTarget ?? undefined">
      <DialogOverlay
        class="office-ui office-overlay__backdrop"
        data-slot="dialog-overlay"
      />
      <DialogContent
        :class="contentClasses"
        data-slot="dialog-content"
        @escape-key-down="onEscapeKeyDown"
        @pointer-down-outside="onPointerDownOutside"
      >
        <VisuallyHidden v-if="props.titleHidden">
          <DialogTitle>{{ resolvedTitle }}</DialogTitle>
          <DialogDescription v-if="props.description || $slots.description">
            <slot name="description">{{ props.description }}</slot>
          </DialogDescription>
        </VisuallyHidden>
        <div v-else class="office-dialog__header" data-slot="dialog-header">
          <DialogTitle class="office-dialog__title" data-slot="dialog-title">
            <slot name="title">{{ resolvedTitle }}</slot>
          </DialogTitle>
          <DialogDescription
            v-if="props.description || $slots.description"
            class="office-dialog__description"
            data-slot="dialog-description"
          >
            <slot name="description">{{ props.description }}</slot>
          </DialogDescription>
          <slot name="header" />
        </div>
        <VisuallyHidden v-if="!props.description && !$slots.description">
          <DialogDescription>{{ resolvedTitle }}</DialogDescription>
        </VisuallyHidden>
        <div class="office-dialog__body" data-slot="dialog-body">
          <slot />
        </div>
        <div v-if="$slots.footer" class="office-dialog__footer" data-slot="dialog-footer">
          <slot name="footer" />
        </div>
        <DialogClose v-if="props.showClose" as-child :aria-label="resolvedCloseLabel">
          <button
            type="button"
            class="office-ui office-button office-button--ghost office-button--size-icon office-dialog__close"
            data-slot="dialog-close"
            :aria-label="resolvedCloseLabel"
          >
            <span class="office-button__content">
              <OfficeIcon name="x" />
            </span>
          </button>
        </DialogClose>
      </DialogContent>
    </DialogPortal>
  </DialogRoot>
</template>
