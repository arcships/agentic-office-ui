<!-- Shared search controls used by OfficeSearchPopover (popover and inline). -->
<script setup lang="ts">
import { computed, ref } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIconButton from "../ui/OfficeIconButton.vue"
import OfficeInput from "../ui/OfficeInput.vue"

export interface OfficeSearchFieldsProps {
  query: string
  resultCount: number
  activeIndex: number
  statusText: string
  caseSensitive: boolean
  wholeWord: boolean
  regex: boolean
  showOptions: boolean
  disabled: boolean
  placeholder: string
  inputType: string
  testIdPrefix: string
  navigationDisabled: boolean
}

const props = defineProps<OfficeSearchFieldsProps>()

const emit = defineEmits<{
  "update:query": [query: string]
  "update:caseSensitive": [value: boolean]
  "update:wholeWord": [value: boolean]
  "update:regex": [value: boolean]
  previous: []
  next: []
}>()

const m = useOfficeUiMessages()
const testId = (suffix: string) => `${props.testIdPrefix}-${suffix}`
const inputRef = ref<InstanceType<typeof OfficeInput> | null>(null)

defineExpose({
  focus: () => inputRef.value?.focus(),
})

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

let composing = false

function onInput(value: string) {
  if (composing) return
  emit("update:query", value)
}

function onCompositionStart() {
  composing = true
}

function onCompositionEnd(event: CompositionEvent) {
  composing = false
  emit("update:query", (event.target as HTMLInputElement).value)
}

function navigate(direction: -1 | 1) {
  if (props.navigationDisabled) return
  if (direction < 0) emit("previous")
  else emit("next")
}

function toggle(key: "caseSensitive" | "wholeWord" | "regex") {
  if (props.disabled) return
  const next = !props[key]
  if (key === "caseSensitive") emit("update:caseSensitive", next)
  else if (key === "wholeWord") emit("update:wholeWord", next)
  else emit("update:regex", next)
}
</script>

<template>
  <OfficeInput
    ref="inputRef"
    :model-value="props.query"
    :type="props.inputType"
    :placeholder="props.placeholder"
    :aria-label="props.placeholder"
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
    >{{ props.statusText }}</span>
    <div class="office-search-popover__navigation">
      <OfficeIconButton
        :label="m.common.search.previous"
        icon="chevron-left"
        variant="outline"
        size="icon-sm"
        :disabled="props.navigationDisabled"
        :data-testid="testId('previous')"
        @click="navigate(-1)"
      />
      <OfficeIconButton
        :label="m.common.search.next"
        icon="chevron-right"
        variant="outline"
        size="icon-sm"
        :disabled="props.navigationDisabled"
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
</template>
