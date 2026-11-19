<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIconButton from "../ui/OfficeIconButton.vue"
import OfficeInput from "../ui/OfficeInput.vue"
import OfficePopover from "../ui/OfficePopover.vue"
import OfficeToggle from "../ui/OfficeToggle.vue"
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
const composing = ref(false)

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
  if (props.status === "searching") return m.value.common.search.searching
  if (props.status === "error") return m.value.common.search.error
  if (!hasQuery.value) return ""
  if (props.resultCount < 1) return m.value.common.search.noResults
  return formatOfficeMessage(m.value.common.search.resultOf, {
    current: Math.min(Math.max(props.activeIndex, 0) + 1, props.resultCount),
    total: props.resultCount,
  })
})

const contentClasses = computed(() =>
  officeClass("office-search-popover__content", props.contentClass),
)

watch(
  isOpen,
  async (open) => {
    if (!open) return
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
  if (!isOpen.value) return
  setOpen(false)
}

function onQueryInput(value: string): void {
  if (props.modelValue === undefined) internalQuery.value = value
  emit("update:modelValue", value)
}

function onCompositionStart(): void {
  composing.value = true
}

function onCompositionEnd(event: CompositionEvent): void {
  composing.value = false
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
  if (event.key !== "Enter" || composing.value) return
  event.preventDefault()
  if (navigationDisabled.value) {
    if (hasQuery.value) emit("submit", query.value)
    return
  }
  navigate(event.shiftKey ? -1 : 1)
}

function toggle(
  key: "caseSensitive" | "wholeWord" | "regex",
): void {
  if (props.disabled) return
  const next = !props[key]
  if (key === "caseSensitive") emit("update:caseSensitive", next)
  else if (key === "wholeWord") emit("update:wholeWord", next)
  else emit("update:regex", next)
}

defineExpose({
  focus: () => contentRoot.value?.querySelector("input")?.focus(),
  open: () => setOpen(true),
  close: () => setOpen(false),
})
</script>

<template>
  <OfficePopover
    :open="isOpen"
    :class="props.class"
    :content-class="contentClasses"
    side="bottom"
    align="end"
    :side-offset="6"
    @update:open="setOpen"
  >
    <template #trigger>
      <slot name="trigger">
        <OfficeIconButton
          :label="triggerLabelText"
          icon="search"
          size="icon-sm"
          :active="isOpen"
          :disabled="props.disabled"
          data-testid="office-search-trigger"
        />
      </slot>
    </template>
    <div
      ref="contentRoot"
      class="office-search-popover"
      data-slot="search-popover"
      data-testid="office-search-popover"
      @keydown="onKeydown"
    >
      <OfficeInput
        :model-value="query"
        :placeholder="placeholderText"
        :aria-label="placeholderText"
        :disabled="props.disabled"
        autocomplete="off"
        spellcheck="false"
        data-slot="search-input"
        data-testid="office-search-input"
        @update:model-value="onQueryInput"
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
          data-testid="office-search-status"
        >{{ statusText }}</span>
        <div class="office-search-popover__navigation">
          <OfficeIconButton
            :label="m.common.search.previous"
            icon="chevron-left"
            variant="outline"
            size="icon-sm"
            :disabled="navigationDisabled"
            data-testid="office-search-previous"
            @click="navigate(-1)"
          />
          <OfficeIconButton
            :label="m.common.search.next"
            icon="chevron-right"
            variant="outline"
            size="icon-sm"
            :disabled="navigationDisabled"
            data-testid="office-search-next"
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
        <OfficeToggle
          :model-value="props.caseSensitive"
          size="xs"
          variant="outline"
          :label="m.common.search.caseSensitive"
          :title="m.common.search.caseSensitive"
          :disabled="props.disabled"
          class="office-search-popover__option"
          data-testid="office-search-case-sensitive"
          @update:model-value="toggle('caseSensitive')"
        >Aa</OfficeToggle>
        <OfficeToggle
          :model-value="props.wholeWord"
          size="xs"
          variant="outline"
          :label="m.common.search.wholeWord"
          :title="m.common.search.wholeWord"
          :disabled="props.disabled"
          class="office-search-popover__option"
          data-testid="office-search-whole-word"
          @update:model-value="toggle('wholeWord')"
        >ab|</OfficeToggle>
        <OfficeToggle
          :model-value="props.regex"
          size="xs"
          variant="outline"
          :label="m.common.search.regex"
          :title="m.common.search.regex"
          :disabled="props.disabled"
          class="office-search-popover__option"
          data-testid="office-search-regex"
          @update:model-value="toggle('regex')"
        >.*</OfficeToggle>
        <slot name="options" />
      </div>
    </div>
  </OfficePopover>
</template>
