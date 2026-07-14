<template>
  <div
    :class="['office-region-selector', props.className, { 'is-disabled': props.disabled, 'is-drawing': drawing }]"
    :tabindex="props.disabled ? -1 : 0"
    role="group"
    :aria-label="props.ariaLabel"
    data-testid="office-region-selector"
    @pointerdown="onPointerDown"
    @pointermove="onPointerMove"
    @pointerup="onPointerUp"
    @pointercancel="cancelDrawing"
    @keydown="onKeydown"
  >
    <div
      v-if="currentRect"
      class="office-region-selector__frame"
      :style="rectStyle(currentRect)"
      data-testid="office-region-frame"
      :aria-label="regionDescription"
      @pointerdown.stop
    >
      <span class="office-region-selector__label">选定区域</span>
      <span v-for="handle in handles" :key="handle" :class="['office-region-selector__handle', `is-${handle}`]" />
    </div>
    <span class="office-region-selector__instructions">
      拖动以框选区域。方向键移动区域，按住 Shift 配合方向键调整大小，Enter 确认，Escape 取消。
    </span>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from "vue"
import { normalizedRectFromPoints } from "@arcships/office-interaction"
import type { NormalizedPoint, NormalizedRect } from "@arcships/office-interaction"
import type { OfficeRegionSelectorProps } from "../types"

const props = withDefaults(defineProps<OfficeRegionSelectorProps>(), {
  modelValue: null,
  disabled: false,
  minSize: 0.005,
  keyboardStep: 0.01,
  ariaLabel: "选择文档区域",
  className: "",
})

const emit = defineEmits<{
  "update:modelValue": [rect: NormalizedRect | null]
  "selection-start": [point: NormalizedPoint]
  "selection-change": [rect: NormalizedRect]
  "selection-commit": [rect: NormalizedRect]
  "selection-cancel": [rect: NormalizedRect | null]
}>()

const handles = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const
const start = ref<NormalizedPoint | null>(null)
const draft = ref<NormalizedRect | null>(null)
const original = ref<NormalizedRect | null>(null)
const activePointerId = ref<number | null>(null)
const currentRect = computed(() => draft.value ?? props.modelValue ?? null)
const drawing = computed(() => activePointerId.value !== null)
const regionDescription = computed(() => currentRect.value
  ? `区域，左侧 ${Math.round(currentRect.value.x * 100)}%，顶部 ${Math.round(currentRect.value.y * 100)}%，宽 ${Math.round(currentRect.value.width * 100)}%，高 ${Math.round(currentRect.value.height * 100)}%`
  : "尚未选择区域",
)

function clamp(value: number, min = 0, max = 1): number {
  if (!Number.isFinite(value)) return min
  return Math.min(max, Math.max(min, value))
}

function minimumSize(): number {
  return clamp(props.minSize, Number.EPSILON, 1)
}

function pointFromEvent(event: PointerEvent): NormalizedPoint {
  const target = event.currentTarget as HTMLElement
  const bounds = target.getBoundingClientRect()
  if (bounds.width <= 0 || bounds.height <= 0) throw new RangeError("region container must have positive dimensions")
  return {
    x: clamp((event.clientX - bounds.left) / bounds.width),
    y: clamp((event.clientY - bounds.top) / bounds.height),
  }
}

function rectStyle(rect: NormalizedRect): Record<string, string> {
  return {
    left: `${rect.x * 100}%`,
    top: `${rect.y * 100}%`,
    width: `${rect.width * 100}%`,
    height: `${rect.height * 100}%`,
  }
}

function validDraft(point: NormalizedPoint): NormalizedRect | null {
  if (!start.value) return null
  try {
    const rect = normalizedRectFromPoints(start.value, point)
    const minimum = minimumSize()
    return rect.width >= minimum && rect.height >= minimum ? rect : null
  } catch {
    return null
  }
}

function onPointerDown(event: PointerEvent): void {
  if (props.disabled || event.button !== 0 || event.isPrimary === false) return
  const point = pointFromEvent(event)
  start.value = point
  draft.value = null
  original.value = props.modelValue ? { ...props.modelValue } : null
  activePointerId.value = event.pointerId
  ;(event.currentTarget as HTMLElement).setPointerCapture?.(event.pointerId)
  emit("selection-start", point)
  event.preventDefault()
}

function onPointerMove(event: PointerEvent): void {
  if (props.disabled || event.pointerId !== activePointerId.value || !start.value) return
  const rect = validDraft(pointFromEvent(event))
  if (!rect) return
  draft.value = rect
  emit("update:modelValue", rect)
  emit("selection-change", rect)
}

function releasePointer(event: PointerEvent): void {
  const target = event.currentTarget as HTMLElement
  if (target.hasPointerCapture?.(event.pointerId)) target.releasePointerCapture?.(event.pointerId)
}

function finishDrawing(): void {
  start.value = null
  draft.value = null
  original.value = null
  activePointerId.value = null
}

function onPointerUp(event: PointerEvent): void {
  if (event.pointerId !== activePointerId.value) return
  releasePointer(event)
  const rect = validDraft(pointFromEvent(event)) ?? draft.value
  if (rect) {
    emit("update:modelValue", rect)
    emit("selection-commit", rect)
  } else {
    emit("selection-cancel", original.value)
  }
  finishDrawing()
}

function cancelDrawing(event?: PointerEvent): void {
  if (event && event.pointerId !== activePointerId.value) return
  if (event) releasePointer(event)
  if (activePointerId.value !== null) {
    emit("update:modelValue", original.value)
    emit("selection-cancel", original.value)
  }
  finishDrawing()
}

function updateFromKeyboard(rect: NormalizedRect): void {
  emit("update:modelValue", rect)
  emit("selection-change", rect)
}

function keyboardRect(event: KeyboardEvent, rect: NormalizedRect): NormalizedRect {
  const step = clamp(props.keyboardStep, Number.EPSILON, 1)
  if (event.shiftKey) {
    const widthDelta = event.key === "ArrowRight" ? step : event.key === "ArrowLeft" ? -step : 0
    const heightDelta = event.key === "ArrowDown" ? step : event.key === "ArrowUp" ? -step : 0
    return {
      ...rect,
      width: clamp(rect.width + widthDelta, minimumSize(), 1 - rect.x),
      height: clamp(rect.height + heightDelta, minimumSize(), 1 - rect.y),
    }
  }
  const xDelta = event.key === "ArrowRight" ? step : event.key === "ArrowLeft" ? -step : 0
  const yDelta = event.key === "ArrowDown" ? step : event.key === "ArrowUp" ? -step : 0
  return {
    ...rect,
    x: clamp(rect.x + xDelta, 0, 1 - rect.width),
    y: clamp(rect.y + yDelta, 0, 1 - rect.height),
  }
}

function onKeydown(event: KeyboardEvent): void {
  if (props.disabled) return
  if (event.key === "Escape") {
    event.preventDefault()
    if (drawing.value) cancelDrawing()
    else emit("selection-cancel", currentRect.value)
  } else if (event.key === "Enter" && currentRect.value) {
    event.preventDefault()
    emit("selection-commit", currentRect.value)
  } else if (event.key.startsWith("Arrow") && currentRect.value) {
    event.preventDefault()
    updateFromKeyboard(keyboardRect(event, currentRect.value))
  }
}
</script>

<style scoped>
.office-region-selector {
  cursor: crosshair;
  inset: 0;
  outline: none;
  position: absolute;
  touch-action: none;
  z-index: 25;
}

.office-region-selector.is-disabled {
  cursor: default;
  pointer-events: none;
}

.office-region-selector:focus-visible {
  box-shadow: inset 0 0 0 3px rgb(37 99 235 / 0.75);
}

.office-region-selector__frame {
  background: rgb(37 99 235 / 0.08);
  border: 2px solid #2563eb;
  box-sizing: border-box;
  min-height: 2px;
  min-width: 2px;
  pointer-events: auto;
  position: absolute;
}

.office-region-selector__label {
  background: #1d4ed8;
  border-radius: 3px 3px 0 0;
  bottom: 100%;
  color: #fff;
  font: 600 11px/20px ui-sans-serif, system-ui, sans-serif;
  left: -2px;
  padding: 0 7px;
  position: absolute;
  white-space: nowrap;
}

.office-region-selector__handle {
  background: #fff;
  border: 2px solid #2563eb;
  border-radius: 50%;
  height: 8px;
  position: absolute;
  transform: translate(-50%, -50%);
  width: 8px;
}

.office-region-selector__handle.is-nw { left: 0; top: 0; }
.office-region-selector__handle.is-n { left: 50%; top: 0; }
.office-region-selector__handle.is-ne { left: 100%; top: 0; }
.office-region-selector__handle.is-e { left: 100%; top: 50%; }
.office-region-selector__handle.is-se { left: 100%; top: 100%; }
.office-region-selector__handle.is-s { left: 50%; top: 100%; }
.office-region-selector__handle.is-sw { left: 0; top: 100%; }
.office-region-selector__handle.is-w { left: 0; top: 50%; }

.office-region-selector__instructions {
  height: 1px;
  margin: -1px;
  overflow: hidden;
  padding: 0;
  position: absolute;
  width: 1px;
  clip: rect(0 0 0 0);
  clip-path: inset(50%);
  white-space: nowrap;
}
</style>
