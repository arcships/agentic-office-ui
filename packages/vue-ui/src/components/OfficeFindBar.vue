<template>
  <div class="office-find-bar" role="search" data-testid="office-find-bar">
    <input
      ref="inputRef"
      class="office-find-bar__input"
      type="search"
      :value="query"
      :placeholder="placeholder"
      :disabled="disabled"
      aria-label="查找文档内容"
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
      aria-label="上一个搜索结果"
      :disabled="navigationDisabled"
      data-testid="office-find-previous"
      @click="emit('previous')"
    >‹</button>
    <button
      type="button"
      aria-label="下一个搜索结果"
      :disabled="navigationDisabled"
      data-testid="office-find-next"
      @click="emit('next')"
    >›</button>
    <button
      type="button"
      aria-label="关闭查找"
      :disabled="disabled"
      data-testid="office-find-close"
      @click="emit('close')"
    >×</button>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from "vue"

type SurfaceSearchStatus = "idle" | "searching" | "ready" | "error"

const props = withDefaults(defineProps<{
  query: string
  status: SurfaceSearchStatus
  activeIndex: number
  resultCount: number
  placeholder?: string
  inputTestId?: string
  disabled?: boolean
}>(), {
  placeholder: "查找",
  inputTestId: "office-find-input",
  disabled: false,
})

const emit = defineEmits<{
  "update:query": [query: string]
  next: []
  previous: []
  close: []
}>()

const inputRef = ref<HTMLInputElement | null>(null)
const composing = ref(false)
const navigationDisabled = computed(() =>
  props.disabled || props.status === "searching" || props.resultCount === 0,
)
const statusText = computed(() => {
  if (props.status === "searching") return "正在搜索"
  if (props.status === "error") return "搜索失败"
  if (!props.query) return ""
  if (!props.resultCount) return "没有结果"
  return `第 ${props.activeIndex + 1} 项，共 ${props.resultCount} 项`
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
})
</script>

<style scoped>
.office-find-bar {
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

.office-find-bar button:disabled {
  cursor: default;
  opacity: .35;
}

.office-find-bar :focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 2px;
}
</style>
