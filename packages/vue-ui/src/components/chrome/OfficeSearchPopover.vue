<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useOfficeUiMessages } from "../../i18n"
import OfficeIconButton from "../ui/OfficeIconButton.vue"
import OfficePopover from "../ui/OfficePopover.vue"
import OfficeSearchFields from "./OfficeSearchFields.vue"
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
const fieldsRef = ref<InstanceType<typeof OfficeSearchFields> | null>(null)

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

const fieldsProps = computed(() => ({
  query: query.value,
  resultCount: props.resultCount,
  activeIndex: props.activeIndex,
  statusText: statusText.value,
  caseSensitive: props.caseSensitive,
  wholeWord: props.wholeWord,
  regex: props.regex,
  showOptions: props.showOptions,
  disabled: props.disabled,
  placeholder: placeholderText.value,
  inputType: props.inputType,
  testIdPrefix: props.testIdPrefix,
  navigationDisabled: navigationDisabled.value,
}))

defineExpose({
  focus: () => {
    if (fieldsRef.value) {
      fieldsRef.value.focus()
      return
    }
    contentRoot.value?.querySelector?.("input")?.focus()
  },
  open: () => setOpen(true),
  close: () => setOpen(false),
})
</script>

<template>
  <div
    v-if="props.inline"
    ref="contentRoot"
    :class="officeClass('office-search-popover', 'office-search-popover--inline', props.class)"
    role="search"
    data-slot="search-popover"
    :data-testid="testId('popover')"
    @keydown="onKeydown"
  >
    <OfficeSearchFields
      ref="fieldsRef"
      v-bind="fieldsProps"
      @update:query="onQueryInput"
      @update:case-sensitive="emit('update:caseSensitive', $event)"
      @update:whole-word="emit('update:wholeWord', $event)"
      @update:regex="emit('update:regex', $event)"
      @previous="emit('previous')"
      @next="emit('next')"
    >
      <template #options><slot name="options" /></template>
    </OfficeSearchFields>
  </div>

  <OfficePopover
    v-else
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
      :data-testid="testId('popover')"
      @keydown="onKeydown"
    >
      <OfficeSearchFields
        ref="fieldsRef"
        v-bind="fieldsProps"
        @update:query="onQueryInput"
        @update:case-sensitive="emit('update:caseSensitive', $event)"
        @update:whole-word="emit('update:wholeWord', $event)"
        @update:regex="emit('update:regex', $event)"
        @previous="emit('previous')"
        @next="emit('next')"
      >
        <template #options><slot name="options" /></template>
      </OfficeSearchFields>
    </div>
  </OfficePopover>
</template>
