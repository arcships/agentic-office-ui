<template>
  <div class="office-find-bar" role="search" data-testid="office-find-bar">
    <input
      ref="inputRef"
      class="office-find-bar__input"
      type="search"
      :value="query"
      :placeholder="placeholderText"
      :disabled="disabled"
      :aria-label="m.common.search.placeholder"
      :data-testid="inputTestId"
      @input="onInput"
      @keydown="onKeydown"
      @compositionstart="composing = true"
      @compositionend="onCompositionEnd"
    />
    <span
      class="office-find-bar__status"
      role="status"
      aria-live="polite"
      aria-atomic="true"
      data-testid="office-find-status"
    >{{ statusText }}</span>
    <button
      type="button"
      class="office-find-bar__option"
      :class="{ 'office-find-bar__option--active': caseSensitive }"
      :aria-label="m.common.search.caseSensitive"
      :title="m.common.search.caseSensitive"
      :aria-pressed="caseSensitive"
      :disabled="disabled"
      data-testid="office-find-case-sensitive"
      @click="emit('update:caseSensitive', !caseSensitive)"
    >&#65;&#97;</button>
    <button
      type="button"
      class="office-find-bar__option"
      :class="{ 'office-find-bar__option--active': wholeWord }"
      :aria-label="m.common.search.wholeWord"
      :title="m.common.search.wholeWord"
      :aria-pressed="wholeWord"
      :disabled="disabled"
      data-testid="office-find-whole-word"
      @click="emit('update:wholeWord', !wholeWord)"
    >&#97;&#98;|</button>
    <button
      type="button"
      class="office-find-bar__option"
      :class="{ 'office-find-bar__option--active': regex }"
      :aria-label="m.common.search.regex"
      :title="m.common.search.regex"
      :aria-pressed="regex"
      :disabled="disabled"
      data-testid="office-find-regex"
      @click="emit('update:regex', !regex)"
    >.*</button>
    <button
      type="button"
      :aria-label="m.common.search.previous"
      :disabled="navigationDisabled"
      data-testid="office-find-previous"
      @click="emit('previous')"
    >‹</button>
    <button
      type="button"
      :aria-label="m.common.search.next"
      :disabled="navigationDisabled"
      data-testid="office-find-next"
      @click="emit('next')"
    >›</button>
    <button
      type="button"
      :aria-label="m.common.search.close"
      :disabled="disabled"
      data-testid="office-find-close"
      @click="emit('close')"
    >×</button>
    <div v-if="showReplace" class="office-find-bar__replace">
      <input
        ref="replaceInputRef"
        class="office-find-bar__input office-find-bar__input--replace"
        type="text"
        :value="replacement"
        :placeholder="m.common.search.replacePlaceholder"
        :aria-label="m.common.search.replacePlaceholder"
        :disabled="disabled"
        data-testid="office-find-replace-input"
        @input="onReplaceInput"
        @compositionstart="replaceComposing = true"
        @compositionend="onReplaceCompositionEnd"
        @keydown.enter.prevent="emit('replace')"
      />
      <button
        type="button"
        class="office-find-bar__action"
        :disabled="replaceDisabled"
        :aria-label="m.common.search.replace"
        :title="m.common.search.replace"
        data-testid="office-find-replace"
        @click="emit('replace')"
      >{{ m.common.search.replace }}</button>
      <button
        type="button"
        class="office-find-bar__action"
        :disabled="replaceAllDisabled"
        :aria-label="m.common.search.replaceAll"
        :title="m.common.search.replaceAll"
        data-testid="office-find-replace-all"
        @click="emit('replaceAll')"
      >{{ m.common.search.replaceAll }}</button>
      <span
        v-if="replaceStatus"
        class="office-find-bar__replace-status"
        role="status"
        aria-live="polite"
        data-testid="office-find-replace-status"
      >{{ replaceStatus }}</span>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useOfficeUiMessages } from "../i18n"

type SurfaceSearchStatus = "idle" | "searching" | "ready" | "error"

const props = withDefaults(defineProps<{
  query: string
  status: SurfaceSearchStatus
  activeIndex: number
  resultCount: number
  placeholder?: string
  inputTestId?: string
  disabled?: boolean
  caseSensitive?: boolean
  wholeWord?: boolean
  regex?: boolean
  /** Render the replace row (input plus Replace / Replace All). */
  showReplace?: boolean
  /** Current text of the replace field. */
  replacement?: string
  replaceDisabled?: boolean
  replaceAllDisabled?: boolean
  /** Localized summary of the last replace command. */
  replaceStatus?: string
}>(), {
  placeholder: undefined,
  inputTestId: "office-find-input",
  disabled: false,
  caseSensitive: false,
  wholeWord: false,
  regex: false,
  showReplace: false,
  replacement: "",
  replaceDisabled: false,
  replaceAllDisabled: false,
  replaceStatus: undefined,
})

const emit = defineEmits<{
  "update:query": [query: string]
  "update:caseSensitive": [value: boolean]
  "update:wholeWord": [value: boolean]
  "update:regex": [value: boolean]
  "update:replacement": [value: string]
  next: []
  previous: []
  close: []
  replace: []
  replaceAll: []
}>()

const inputRef = ref<HTMLInputElement | null>(null)
const replaceInputRef = ref<HTMLInputElement | null>(null)
const composing = ref(false)
const replaceComposing = ref(false)
const m = useOfficeUiMessages()
const placeholderText = computed(() => props.placeholder ?? m.value.common.search.placeholder)
const navigationDisabled = computed(() =>
  props.disabled || props.status === "searching" || props.resultCount === 0,
)
const statusText = computed(() => {
  if (props.status === "searching") return m.value.common.search.searching
  if (props.status === "error") return m.value.common.search.error
  if (!props.query) return ""
  if (!props.resultCount) return m.value.common.search.noResults
  return formatOfficeMessage(m.value.common.search.resultOf, {
    current: props.activeIndex + 1,
    total: props.resultCount,
  })
})

function inputValue(event: Event): string {
  return (event.target as HTMLInputElement).value
}

function onInput(event: Event): void {
  if (!composing.value) emit("update:query", inputValue(event))
}

function onCompositionEnd(event: CompositionEvent): void {
  composing.value = false
  emit("update:query", inputValue(event))
}

function onReplaceInput(event: Event): void {
  if (!replaceComposing.value) emit("update:replacement", inputValue(event))
}

function onReplaceCompositionEnd(event: CompositionEvent): void {
  replaceComposing.value = false
  emit("update:replacement", inputValue(event))
}

function onKeydown(event: KeyboardEvent): void {
  if (event.key === "Enter" && !navigationDisabled.value) {
    event.preventDefault()
    if (event.shiftKey) emit("previous")
    else emit("next")
  } else if (event.key === "Escape") {
    event.preventDefault()
    emit("close")
  }
}

defineExpose({
  focus: () => inputRef.value?.focus(),
  focusReplace: () => replaceInputRef.value?.focus(),
})
</script>

<style scoped>
.office-find-bar {
  align-items: center;
  display: inline-flex;
  flex-wrap: wrap;
  gap: 4px;
}

.office-find-bar__replace {
  align-items: center;
  display: inline-flex;
  gap: 4px;
}

.office-find-bar__input {
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 7px;
  color: #18181b;
  font: inherit;
  height: 32px;
  padding: 0 8px;
  width: 152px;
}

.office-find-bar__status {
  color: #71717a;
  font-size: 11px;
  min-width: 74px;
  text-align: center;
  white-space: nowrap;
}

.office-find-bar button {
  background: transparent;
  border: 1px solid transparent;
  border-radius: 5px;
  color: inherit;
  cursor: pointer;
  font: inherit;
  font-size: 18px;
  height: 28px;
  line-height: 20px;
  padding: 0;
  width: 26px;
}

.office-find-bar button:hover:not(:disabled) {
  background: #f4f4f5;
}

.office-find-bar button.office-find-bar__option {
  font-size: 11px;
  font-weight: 700;
  line-height: 1;
}

.office-find-bar button.office-find-bar__option--active {
  background: #e0e7ff;
  border-color: #818cf8;
  color: #312e81;
}

.office-find-bar button:disabled {
  cursor: default;
  opacity: .35;
}

.office-find-bar :focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 2px;
}

.office-find-bar button.office-find-bar__action {
  border-color: #d4d4d8;
  font-size: 12px;
  height: 28px;
  padding: 0 8px;
  width: auto;
}

.office-find-bar button.office-find-bar__action:hover:not(:disabled) {
  background: #f4f4f5;
}

.office-find-bar__input--replace {
  width: 132px;
}

.office-find-bar__replace-status {
  color: #71717a;
  font-size: 11px;
  white-space: nowrap;
}
</style>
