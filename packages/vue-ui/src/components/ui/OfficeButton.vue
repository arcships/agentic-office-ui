<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, useSlots, type VNode } from "vue"
import { Primitive } from "reka-ui"
import OfficeSpinner from "./OfficeSpinner.vue"
import { useOfficeUiLabels } from "./labels"
import {
  officeClass,
  type OfficeButtonSize,
  type OfficeVariant,
} from "./types"

export interface OfficeButtonProps {
  /** Visual style. Defaults to the solid neutral button. */
  variant?: OfficeVariant
  /** Control height or square icon size. */
  size?: OfficeButtonSize
  /** Disables pointer and keyboard activation. */
  disabled?: boolean
  /**
   * Shows a centered spinner, keeps the label width, sets `aria-busy` and
   * keeps the button focusable with `aria-disabled` while activation is
   * suppressed.
   */
  loading?: boolean
  /**
   * Accessible name announced while `loading` is true. Defaults to the visible
   * slot text, then the localized "Loading" label. `aria-busy` hides busy
   * descendants from assistive technology, so the name must be provided on the
   * control itself.
   */
  loadingLabel?: string
  /** Render the slot's single root element instead of a `<button>`. */
  asChild?: boolean
  /** Native button type; ignored when `asChild` renders a non-button. */
  type?: "button" | "submit" | "reset"
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficeButtonProps>(), {
  variant: "default",
  size: "default",
  disabled: false,
  loading: false,
  loadingLabel: undefined,
  asChild: false,
  type: "button",
  class: "",
})

const emit = defineEmits<{
  click: [event: MouseEvent]
}>()

const slots = useSlots()
const labels = useOfficeUiLabels()

const classes = computed(() =>
  officeClass(
    "office-ui",
    "office-button",
    `office-button--${props.variant}`,
    `office-button--size-${props.size}`,
    props.class,
  ),
)

/** Flatten the default slot to text so a loading button keeps its visible name. */
function nodeText(child: unknown): string {
  if (typeof child === "string" || typeof child === "number") {
    return String(child)
  }
  if (!child || typeof child !== "object") return ""
  const node = child as VNode
  if (typeof node.children === "string") return node.children
  if (Array.isArray(node.children)) {
    return node.children.map(nodeText).join(" ")
  }
  return ""
}

const loadingName = computed(() => {
  const explicit = props.loadingLabel?.trim()
  if (explicit) return explicit
  const visible = (slots.default?.() ?? [])
    .map(nodeText)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim()
  return visible || labels.value.loading
})

const accessibleLabel = computed(() =>
  props.loading ? loadingName.value : undefined,
)

/**
 * `aria-disabled` instead of the native `disabled` attribute keeps the
 * loading button in the tab order, so keyboard users can still discover it.
 * Suppress activation (including a submit or `asChild` link) explicitly.
 */
function onActivation(event: MouseEvent): void {
  if (props.disabled || props.loading) {
    event.preventDefault()
    event.stopPropagation()
    return
  }
  emit("click", event)
}

const spinnerSize = "default" as const
</script>

<template>
  <Primitive
    v-if="props.asChild"
    as-child
    :class="classes"
    :data-disabled="props.disabled || props.loading ? 'true' : undefined"
    :data-loading="props.loading ? 'true' : undefined"
    :aria-busy="props.loading || undefined"
    :aria-disabled="props.disabled || props.loading ? 'true' : undefined"
    :aria-label="accessibleLabel"
    data-slot="button"
    @click="onActivation"
  >
    <slot />
  </Primitive>
  <Primitive
    v-else
    as="button"
    :class="classes"
    :type="props.type"
    :disabled="props.disabled"
    :data-disabled="props.disabled || props.loading ? 'true' : undefined"
    :data-loading="props.loading ? 'true' : undefined"
    :aria-busy="props.loading || undefined"
    :aria-disabled="props.disabled || props.loading ? 'true' : undefined"
    :aria-label="accessibleLabel"
    data-slot="button"
    @click="onActivation"
  >
    <!-- Decorative: the button owns the accessible name and aria-busy state. -->
    <span v-if="props.loading" class="office-button__spinner">
      <OfficeSpinner :size="spinnerSize" :label="null" />
    </span>
    <span class="office-button__content">
      <slot />
    </span>
  </Primitive>
</template>
