<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue"
import { useOfficeUiMessages } from "../../i18n"
import OfficeInput from "../ui/OfficeInput.vue"
import { officeClass } from "../ui/types"

export interface OfficePageControlProps {
  /**
   * Current page, 1-based. Leave undefined for uncontrolled use; the
   * component then keeps its own page and still emits `update:modelValue`.
   */
  modelValue?: number
  /** Initial page for uncontrolled use. */
  defaultPage?: number
  /** Last page of the document. `0` renders an em dash. */
  totalPages: number
  disabled?: boolean
  /** Accessible label of the page input; defaults to `common.page.current`. */
  label?: string
  /**
   * Render a native `<input>` instead of `OfficeInput`. Hosts that integrate
   * the control with a form or a renderer without Vue's component patches use
   * this; the visual treatment and keyboard semantics are identical.
   */
  nativeInput?: boolean
  /** Optional `data-testid` for the page input element. */
  inputTestId?: string
  /** Extra classes merged on the root. */
  class?: string
}

const props = withDefaults(defineProps<OfficePageControlProps>(), {
  modelValue: undefined,
  defaultPage: 1,
  disabled: false,
  label: undefined,
  nativeInput: false,
  inputTestId: undefined,
  class: "",
})

const emit = defineEmits<{
  "update:modelValue": [page: number]
  change: [page: number]
}>()

const m = useOfficeUiMessages()

const internalPage = ref(normalize(props.defaultPage))
const focused = ref(false)
const draft = ref(String(currentPage()))
const inputRef = ref<
  InstanceType<typeof OfficeInput> | HTMLInputElement | null
>(null)

function upperBound(): number {
  return Math.max(Math.trunc(props.totalPages) || 0, 0)
}

function normalize(value: number): number {
  const upper = upperBound()
  const integer = Number.isFinite(value) ? Math.trunc(value) : 1
  if (upper < 1) return 1
  return Math.min(Math.max(integer, 1), upper)
}

function currentPage(): number {
  return normalize(props.modelValue ?? internalPage.value)
}

const page = computed(() => currentPage())
const total = computed(() => upperBound())
const displayTotal = computed(() => (total.value > 0 ? String(total.value) : "–"))
const classes = computed(() =>
  officeClass("office-ui", "office-page-control", props.class),
)
const inputClasses = computed(() =>
  officeClass(
    "office-input",
    "office-input--sm",
    "office-page-control__input",
  ),
)

watch(page, (next) => {
  // Keep a focused draft intact while the user is typing a page number.
  if (focused.value) return
  draft.value = String(next)
})

function commit(raw: string = draft.value): void {
  const parsed = Number.parseInt(raw.trim(), 10)
  if (!Number.isInteger(parsed) || total.value < 1) {
    draft.value = String(page.value)
    return
  }
  const next = normalize(parsed)
  draft.value = String(next)
  if (next === page.value) return
  if (props.modelValue === undefined) internalPage.value = next
  emit("update:modelValue", next)
  emit("change", next)
}

function reset(): void {
  draft.value = String(page.value)
}

function step(direction: -1 | 1): void {
  if (props.disabled || total.value < 1) return
  const next = normalize(page.value + direction)
  draft.value = String(next)
  commit(String(next))
}

function onUpdate(value: string): void {
  draft.value = value
}

function onFocus(): void {
  focused.value = true
  nextTick(() => {
    const element = resolveInput()
    element?.select()
  })
}

function onBlur(): void {
  focused.value = false
  commit()
}

function onKeydown(event: KeyboardEvent): void {
  if (props.disabled) return
  switch (event.key) {
    case "Enter":
      event.preventDefault()
      commit()
      break
    case "Escape":
      event.preventDefault()
      reset()
      break
    case "ArrowUp":
      event.preventDefault()
      step(1)
      break
    case "ArrowDown":
      event.preventDefault()
      step(-1)
      break
    case "Home":
      event.preventDefault()
      draft.value = "1"
      commit("1")
      break
    case "End":
      event.preventDefault()
      if (total.value > 0) {
        draft.value = String(total.value)
        commit(draft.value)
      }
      break
    default:
      break
  }
}

/**
 * `OfficeInput` renders the native input as its root, so the component ref is
 * the element in Vue 3. This guard keeps the call safe in test renderers that
 * expose only a proxy.
 */
function resolveInput(): HTMLInputElement | null {
  const instance = inputRef.value as unknown
  if (!instance) return null
  // Non-DOM renderers (SSR, component tests) may not define the global
  // constructors at all; the component must not throw in a deferred tick.
  const hasElement = typeof HTMLInputElement !== "undefined"
  if (hasElement && instance instanceof HTMLInputElement) return instance
  const root = (instance as { $el?: unknown }).$el
  return hasElement && root instanceof HTMLInputElement
    ? (root as HTMLInputElement)
    : null
}

defineExpose({
  focus: () => resolveInput()?.focus(),
})
</script>

<template>
  <div
    :class="classes"
    role="group"
    :aria-label="m.common.page.navigation"
    data-slot="page-control"
    data-testid="office-page-control"
  >
    <span class="office-page-control__label">{{ m.common.page.label }}</span>
    <input
      v-if="props.nativeInput"
      ref="inputRef"
      v-model="draft"
      :class="inputClasses"
      :aria-label="props.label ?? m.common.page.current"
      inputmode="numeric"
      pattern="[0-9]*"
      autocomplete="off"
      :disabled="props.disabled || total < 1"
      data-slot="page-input"
      :data-testid="props.inputTestId ?? 'office-page-input'"
      @focus="onFocus"
      @blur="onBlur"
      @keydown="onKeydown"
    />
    <OfficeInput
      v-else
      ref="inputRef"
      :class="inputClasses"
      :model-value="draft"
      :aria-label="props.label ?? m.common.page.current"
      inputmode="numeric"
      pattern="[0-9]*"
      autocomplete="off"
      :disabled="props.disabled || total < 1"
      data-slot="page-input"
      :data-testid="props.inputTestId ?? 'office-page-input'"
      @update:model-value="onUpdate"
      @focus="onFocus"
      @blur="onBlur"
      @keydown="onKeydown"
    />
    <span
      class="office-page-control__total"
      data-slot="page-total"
      data-testid="office-page-total"
    >{{ m.common.page.of }} {{ displayTotal }}</span>
  </div>
</template>
