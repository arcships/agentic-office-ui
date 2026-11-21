<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIconButton from "../ui/OfficeIconButton.vue"
import OfficeInput from "../ui/OfficeInput.vue"
import OfficePopover from "../ui/OfficePopover.vue"
import { officeClass } from "../ui/types"
import type { OfficeSearchStatus } from "./types"

export interface OfficeSearchPopoverProps {
  /** Search query (`v-model`). Leave undefined for uncontrolled use. */
  modelValue?: string
  /** Initial query for uncontrolled use. */
  defaultQuery?: string
  /** Number of hits in the active search session. */
  resultCount?: number
  /** Zero-based index of the active hit. */
  activeIndex?: number
  status?: OfficeSearchStatus
  caseSensitive?: boolean
  wholeWord?: boolean
  regex?: boolean
  /** Show the match-case / whole-word / regex toggles. */
  showOptions?: boolean
  /** Controlled open state (`v-model:open`). */
  open?: boolean
  /** Initial open state for uncontrolled use. */
  defaultOpen?: boolean
  disabled?: boolean
  placeholder?: string
  /** Accessible name of the trigger when no `trigger` slot is given. */
  triggerLabel?: string
  /**
   * Render the controls in place instead of a trigger + portalled popover.
   * Hosts that keep a permanent search field in their toolbar use this; the
   * input, counts, options and keyboard contract are identical.
   */
  inline?: boolean
  /** Native input type. Defaults to `text`; viewers can set `search`. */
  inputType?: string
  /** Override the localized status line (for example `1/4`). */
  statusText?: string
  /** Prefix for stable `data-testid` hooks (`<prefix>-input`, `-status`, …). */
  testIdPrefix?: string
  /** Extra classes merged on the trigger. */
  class?: string
  /** Extra classes merged on the portalled content. */
  contentClass?: string
}

const props = withDefaults(defineProps<OfficeSearchPopoverProps>(), {
  modelValue: undefined,
  defaultQuery: "",
  resultCount: 0,
  activeIndex: 0,
  status: "idle",
  caseSensitive: false,
  wholeWord: false,
  regex: false,
  showOptions: true,
  open: undefined,
  defaultOpen: false,
  disabled: false,
  placeholder: undefined,
  triggerLabel: undefined,
  inline: false,
  inputType: "text",
  statusText: undefined,
  testIdPrefix: "office-search",
  class: "",
  contentClass: "",
})

const emit = defineEmits<{
  "update:modelValue": [query: string]
  "update:open": [open: boolean]
  "update:caseSensitive": [value: boolean]
  "update:wholeWord": [value: boolean]
  "update:regex": [value: boolean]
  next: []
  previous: []
  submit: [query: string]
  close: []
}>()

const m = useOfficeUiMessages()

const internalQuery = ref(props.defaultQuery)
const internalOpen = ref(props.defaultOpen)
const contentRoot = ref<HTMLElement | null>(null)
const inputRef = ref<InstanceType<typeof OfficeInput> | null>(null)
let composing = false

const query = computed(() => props.modelValue ?? internalQuery.value)
const isOpen = computed(() => props.open ?? internalOpen.value)
const placeholderText = computed(
  () => props.placeholder ?? m.value.common.search.placeholder,
)
const triggerLabelText = computed(
  () => props.triggerLabel ?? m.value.common.search.placeholder,
)
const hasQuery = computed(() => query.value.trim().length > 0)
const navigationDisabled = computed(
  () => props.disabled || props.status === "searching" || props.resultCount < 1,
)
const statusText = computed(() => {
  if (props.statusText !== undefined) return props.statusText
  if (props.status === "searching") return m.value.common.search.searching
  if (props.status === "error") return m.value.common.search.error
  if (!hasQuery.value) return ""
  if (props.resultCount < 1) return m.value.common.search.noResults
  return formatOfficeMessage(m.value.common.search.resultOf, {
    current: Math.min(Math.max(props.activeIndex, 0) + 1, props.resultCount),
    total: props.resultCount,
  })
})

const testId = (suffix: string) => `${props.testIdPrefix}-${suffix}`

const optionEntries = computed(() => [
  {
    key: "caseSensitive" as const,
    suffix: "case-sensitive",
    label: m.value.common.search.caseSensitive,
    text: "Aa",
    value: props.caseSensitive,
  },
  {
    key: "wholeWord" as const,
    suffix: "whole-word",
    label: m.value.common.search.wholeWord,
    text: "ab|",
    value: props.wholeWord,
  },
  {
    key: "regex" as const,
    suffix: "regex",
    label: m.value.common.search.regex,
    text: ".*",
    value: props.regex,
  },
])

const contentClasses = computed(() =>
  officeClass("office-search-popover__content", props.contentClass),
)

watch(
  isOpen,
  async (open) => {
    if (props.inline || !open) return
    await nextTick()
    const input = contentRoot.value?.querySelector("input")
    input?.focus()
    input?.select()
  },
  { immediate: true },
)

function setOpen(value: boolean): void {
  if (props.open === undefined) internalOpen.value = value
  emit("update:open", value)
  if (!value) emit("close")
}

function requestClose(): void {
  if (props.inline) {
    emit("close")
    return
  }
  if (!isOpen.value) return
  setOpen(false)
}

function onQueryInput(value: string): void {
  if (props.modelValue === undefined) internalQuery.value = value
  emit("update:modelValue", value)
}

function onInput(value: string): void {
  if (composing) return
  onQueryInput(value)
}

function onCompositionStart(): void {
  composing = true
}

function onCompositionEnd(event: CompositionEvent): void {
  composing = false
  onQueryInput((event.target as HTMLInputElement).value)
}

function navigate(direction: -1 | 1): void {
  if (navigationDisabled.value) return
  if (direction < 0) emit("previous")
  else emit("next")
}

function onKeydown(event: KeyboardEvent): void {
  if (event.key === "Escape") {
    event.preventDefault()
    requestClose()
    return
  }
  if (event.key !== "Enter" || event.isComposing) return
  event.preventDefault()
  if (navigationDisabled.value) {
    if (hasQuery.value) emit("submit", query.value)
    return
  }
  navigate(event.shiftKey ? -1 : 1)
}

function toggle(key: "caseSensitive" | "wholeWord" | "regex"): void {
  if (props.disabled) return
  const next = !props[key]
  if (key === "caseSensitive") emit("update:caseSensitive", next)
  else if (key === "wholeWord") emit("update:wholeWord", next)
  else emit("update:regex", next)
}

defineExpose({
  focus: () => inputRef.value?.focus(),
  open: () => setOpen(true),
  close: () => setOpen(false),
})
</script>

<template>
  <component
    :is="props.inline ? 'div' : OfficePopover"
    :open="props.inline ? undefined : isOpen"
    :class="
      props.inline
        ? officeClass('office-search-popover', 'office-search-popover--inline', props.class)
        : props.class
    "
    :content-class="props.inline ? undefined : contentClasses"
    :side="props.inline ? undefined : 'bottom'"
    :align="props.inline ? undefined : 'end'"
    :side-offset="props.inline ? undefined : 6"
    :role="props.inline ? 'search' : undefined"
    :data-slot="props.inline ? 'search-popover' : undefined"
    :data-testid="props.inline ? testId('popover') : undefined"
    @update:open="setOpen"
  >
    <template v-if="!props.inline" #trigger>
      <slot name="trigger">
        <OfficeIconButton
          :label="triggerLabelText"
          icon="search"
          size="icon-sm"
          :active="isOpen"
          :disabled="props.disabled"
          :data-testid="testId('trigger')"
        />
      </slot>
    </template>
    <div
      ref="contentRoot"
      :class="props.inline ? 'office-search-popover__fields' : 'office-search-popover'"
      :role="props.inline ? undefined : 'search'"
      :data-slot="props.inline ? undefined : 'search-popover'"
      :data-testid="props.inline ? undefined : testId('popover')"
      @keydown="onKeydown"
    >
      <OfficeInput
        ref="inputRef"
        :model-value="query"
        :type="props.inputType"
        :placeholder="placeholderText"
        :aria-label="placeholderText"
        :disabled="props.disabled"
        autocomplete="off"
        spellcheck="false"
        data-slot="search-input"
        :data-testid="testId('input')"
        @update:model-value="onInput"
        @compositionstart="onCompositionStart"
        @compositionend="onCompositionEnd"
      />
      <div class="office-search-popover__row">
        <span
          class="office-search-popover__status"
          role="status"
          aria-live="polite"
          aria-atomic="true"
          data-slot="search-status"
          :data-testid="testId('status')"
        >{{ statusText }}</span>
        <div class="office-search-popover__navigation">
          <OfficeIconButton
            :label="m.common.search.previous"
            icon="chevron-left"
            variant="outline"
            size="icon-sm"
            :disabled="navigationDisabled"
            :data-testid="testId('previous')"
            @click="navigate(-1)"
          />
          <OfficeIconButton
            :label="m.common.search.next"
            icon="chevron-right"
            variant="outline"
            size="icon-sm"
            :disabled="navigationDisabled"
            :data-testid="testId('next')"
            @click="navigate(1)"
          />
        </div>
      </div>
      <div
        v-if="props.showOptions"
        class="office-search-popover__options"
        role="group"
        :aria-label="m.common.search.placeholder"
      >
        <button
          v-for="entry of optionEntries"
          :key="entry.key"
          type="button"
          class="office-ui office-toggle office-toggle--outline office-toggle--xs office-search-popover__option"
          :data-state="entry.value ? 'on' : 'off'"
          :aria-pressed="entry.value"
          :aria-label="entry.label"
          :title="entry.label"
          :disabled="props.disabled"
          :data-testid="testId(entry.suffix)"
          @click="toggle(entry.key)"
        >{{ entry.text }}</button>
        <slot name="options" />
      </div>
    </div>
  </component>
</template>
