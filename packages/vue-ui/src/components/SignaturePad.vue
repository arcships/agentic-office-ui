<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue"
import { CheckmarkCircle02Icon, Delete02Icon, Pen01Icon } from "@hugeicons/core-free-icons"
import { HugeIcon } from "./HugeIcon"
import OfficeButton from "./ui/OfficeButton.vue"
import { useOfficeUiMessages } from "../i18n"
import type { SignaturePadProps } from "../types"

const props = withDefaults(defineProps<SignaturePadProps>(), {
  width: 400,
  height: 200,
  penColor: "#000000",
  backgroundColor: "transparent",
  className: "",
  mode: "draw",
  modelValue: undefined,
  defaultValue: undefined,
  placeholder: undefined,
  showControls: true,
  showClear: true,
  showCancel: false,
  showConfirm: false,
  confirmDisabled: false,
  disabled: false,
  guide: false,
  loading: false,
  error: undefined,
  ariaLabel: undefined,
})

const emit = defineEmits<{
  "update:signature": [dataUrl: string | null]
  "update:modelValue": [value: string | null]
  clear: []
  cancel: []
  confirm: [value: string | null]
}>()

const m = useOfficeUiMessages()

const canvasRef = ref<HTMLCanvasElement | null>(null)
const typedInput = ref<HTMLInputElement | null>(null)
const typedValue = ref(props.mode === "typed" ? props.modelValue ?? "" : "")
const drawnValue = ref<string | null>(props.defaultValue ?? null)
const isEmpty = ref(true)
let ctx: CanvasRenderingContext2D | null = null
let drawing = false
let canvasRatio = 1

const isTyped = computed(() => props.mode === "typed")
const hasValue = computed(() =>
  isTyped.value ? typedValue.value.trim().length > 0 : !isEmpty.value,
)
const confirmDisabled = computed(
  () => props.disabled || props.confirmDisabled || Boolean(props.error) || !hasValue.value,
)
const placeholder = computed(
  () => props.placeholder ?? m.value.vueUi.signaturePad.typePlaceholder,
)
const ariaLabel = computed(
  () => props.ariaLabel ?? m.value.vueUi.signaturePad.previewAlt,
)

function setupCanvas() {
  const canvas = canvasRef.value
  if (!canvas) return
  canvasRatio =
    typeof window === "undefined" ? 1 : window.devicePixelRatio || 1
  const width = Math.max(1, props.width)
  const height = Math.max(1, props.height)
  canvas.width = Math.floor(width * canvasRatio)
  canvas.height = Math.floor(height * canvasRatio)
  if (canvas.style) {
    canvas.style.width = "100%"
    canvas.style.maxWidth = `${width}px`
    canvas.style.height = `${height}px`
  }
  ctx = canvas.getContext("2d")
  if (!ctx) return
  ctx.setTransform(canvasRatio, 0, 0, canvasRatio, 0, 0)
  ctx.lineCap = "round"
  ctx.lineJoin = "round"
  ctx.lineWidth = 2
  ctx.strokeStyle = props.penColor
  if (props.backgroundColor !== "transparent") {
    ctx.fillStyle = props.backgroundColor
    ctx.fillRect(0, 0, width, height)
  }
}

function point(event: PointerEvent) {
  const rect = canvasRef.value!.getBoundingClientRect()
  return {
    x: event.clientX - rect.left,
    y: event.clientY - rect.top,
  }
}

function emitValue() {
  if (isTyped.value) {
    const value = typedValue.value.trim() ? typedValue.value : null
    emit("update:modelValue", value)
    return
  }
  const value =
    isEmpty.value || !canvasRef.value
      ? null
      : canvasRef.value.toDataURL("image/png")
  drawnValue.value = value
  emit("update:signature", value)
  emit("update:modelValue", value)
}

function onPointerDown(event: PointerEvent) {
  if (props.disabled || !ctx || !canvasRef.value) return
  drawing = true
  canvasRef.value.setPointerCapture?.(event.pointerId)
  const position = point(event)
  ctx.beginPath()
  ctx.moveTo(position.x, position.y)
  event.preventDefault?.()
}

function onPointerMove(event: PointerEvent) {
  if (!drawing || props.disabled || !ctx) return
  const position = point(event)
  ctx.lineTo(position.x, position.y)
  ctx.stroke()
  isEmpty.value = false
  event.preventDefault?.()
}

function onPointerUp(event: PointerEvent) {
  if (!drawing) return
  drawing = false
  canvasRef.value?.releasePointerCapture?.(event.pointerId)
  emitValue()
}

function clearPad(): string | null {
  if (ctx && canvasRef.value) {
    ctx.clearRect(0, 0, props.width, props.height)
    if (props.backgroundColor !== "transparent") {
      ctx.fillStyle = props.backgroundColor
      ctx.fillRect(0, 0, props.width, props.height)
    }
  }
  isEmpty.value = true
  drawnValue.value = null
  if (isTyped.value) typedValue.value = ""
  emit("update:signature", null)
  emit("update:modelValue", null)
  emit("clear")
  return null
}

function onTypedInput(event: Event) {
  typedValue.value = (event.target as HTMLInputElement).value
  emitValue()
}

function confirm() {
  if (confirmDisabled.value) return
  const value = isTyped.value
    ? typedValue.value.trim() || null
    : drawnValue.value ?? canvasRef.value?.toDataURL("image/png") ?? null
  emit("confirm", value)
}

function cancel() {
  emit("cancel")
}

function toDataURL(type?: string, encoderOptions?: number): string {
  return canvasRef.value?.toDataURL(type, encoderOptions) ?? ""
}

function isEmptyState(): boolean {
  return !hasValue.value
}

defineExpose({
  clear: clearPad,
  isEmpty: isEmptyState,
  toDataURL,
  confirm,
  focus: () => (isTyped.value ? typedInput.value?.focus?.() : canvasRef.value?.focus?.()),
})

watch(
  () => props.penColor,
  (color) => {
    if (ctx) ctx.strokeStyle = color
  },
)

watch(
  () => props.backgroundColor,
  () => {
    clearPad()
  },
)

watch(
  () => props.modelValue,
  async (value) => {
    if (isTyped.value) {
      typedValue.value = value ?? ""
      return
    }
    if (value === drawnValue.value) return
    drawnValue.value = value ?? null
    isEmpty.value = !value
    await nextTick()
    // Data-URL restoration is intentionally not painted back into the canvas:
    // the canvas stays a live input, and the emitted value is the host's
    // source of truth. Consumers that need a preview render an <img>.
  },
)

watch(
  () => props.mode,
  () => {
    typedValue.value = props.mode === "typed" ? props.modelValue ?? "" : ""
  },
)

onMounted(setupCanvas)
onBeforeUnmount(() => {
  ctx = null
})
</script>

<template>
  <div :class="['signature-pad', 'office-ui', props.className]" data-slot="signature-pad">
    <div v-if="props.error" class="signature-pad__state signature-pad__state--error" role="alert">
      {{ props.error }}
    </div>
    <div v-else-if="props.loading" class="signature-pad__state" role="status">
      {{ m.vueUi.signaturePad.loading }}
    </div>

    <template v-else>
      <div v-if="isTyped" class="signature-pad__typed">
        <input
          ref="typedInput"
          class="signature-pad__typed-input"
          type="text"
          :value="typedValue"
          :placeholder="placeholder"
          :disabled="props.disabled"
          :aria-label="ariaLabel"
          data-slot="signature-typed-input"
          @input="onTypedInput"
        />
        <div
          :class="['signature-pad__typed-preview', { 'is-empty': !typedValue }]"
          data-slot="signature-typed-preview"
        >
          {{ typedValue || placeholder }}
        </div>
      </div>

      <div v-else :class="['signature-pad__frame', { 'has-guide': props.guide }]">
        <div class="signature-pad__stage">
          <div
            :class="['signature-pad__guide', { 'is-disabled': props.disabled }]"
            :style="{ width: `${props.width}px`, maxWidth: '100%' }"
            data-slot="signature-canvas-guide"
          >
            <canvas
              ref="canvasRef"
              class="signature-pad__canvas"
              :aria-label="ariaLabel"
              role="img"
              :tabindex="props.disabled ? -1 : 0"
              :style="{ touchAction: 'none' }"
              @pointerdown="onPointerDown"
              @pointermove="onPointerMove"
              @pointerup="onPointerUp"
              @pointercancel="onPointerUp"
            />
          </div>
        </div>
      </div>

      <div v-if="props.showControls" class="signature-pad__controls">
        <OfficeButton
          v-if="props.showClear"
          type="button"
          variant="outline"
          :disabled="props.disabled"
          data-slot="signature-clear"
          @click="clearPad"
        >
          <HugeIcon :icon="Delete02Icon" class="signature-pad__control-icon" />
          {{ m.vueUi.signaturePad.clear }}
        </OfficeButton>
        <div class="signature-pad__controls-end">
          <OfficeButton
            v-if="props.showCancel"
            type="button"
            variant="outline"
            :disabled="props.disabled"
            data-slot="signature-cancel"
            @click="cancel"
          >
            {{ m.vueUi.signaturePad.cancel }}
          </OfficeButton>
          <OfficeButton
            v-if="props.showConfirm"
            type="button"
            :disabled="confirmDisabled"
            data-slot="signature-confirm"
            @click="confirm"
          >
            <HugeIcon :icon="hasValue ? CheckmarkCircle02Icon : Pen01Icon" class="signature-pad__control-icon" />
            {{ m.vueUi.signaturePad.confirm }}
          </OfficeButton>
        </div>
      </div>
    </template>
  </div>
</template>

<style scoped>
.signature-pad {
  min-width: 0;
  color: var(--office-foreground);
}
.signature-pad,
.signature-pad * {
  box-sizing: border-box;
}
.signature-pad__state {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: color-mix(in oklab, var(--office-muted) 30%, transparent);
  padding: var(--office-ui-space-3);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.signature-pad__state--error {
  border-color: color-mix(in oklab, var(--office-destructive) 45%, transparent);
  background: color-mix(in oklab, var(--office-destructive) 8%, transparent);
  color: var(--office-destructive);
}
.signature-pad__frame {
  border: var(--office-border-width) solid var(--office-border);
  border-radius: var(--office-radius-xl);
  background: #fff;
  padding: var(--office-ui-space-3);
  color: #020617;
  box-shadow: var(--office-shadow-xs);
}
.signature-pad__stage {
  display: flex;
  height: 224px;
  width: 100%;
  align-items: center;
  justify-content: center;
  overflow: hidden;
  border-radius: var(--office-radius-lg);
  background: #fff;
  padding: var(--office-ui-space-2);
}
.signature-pad__guide {
  position: relative;
  overflow: hidden;
  border-radius: 3px;
  border: 1px dashed color-mix(in oklab, oklch(62.3% .214 259.815) 70%, transparent);
  background: #fff;
  touch-action: none;
}
.signature-pad__guide.is-disabled {
  opacity: 0.64;
  cursor: not-allowed;
}
.signature-pad__canvas {
  display: block;
  width: 100%;
  height: 100%;
  background: #fff;
  touch-action: none;
  cursor: crosshair;
  outline: none;
}
.signature-pad__canvas:focus-visible {
  box-shadow: inset 0 0 0 2px var(--office-ring);
}
.signature-pad__typed {
  display: flex;
  flex-direction: column;
  gap: var(--office-ui-space-3);
}
.signature-pad__typed-input {
  width: 100%;
  height: 34px;
  border: var(--office-border-width) solid var(--office-input);
  border-radius: var(--office-radius-lg);
  background: var(--office-popover);
  padding-inline: calc(var(--office-ui-space-3) - 1px);
  color: var(--office-foreground);
  font: inherit;
  font-size: var(--office-text-sm);
  outline: none;
}
@media (min-width: 640px) {
  .signature-pad__typed-input {
    height: 30px;
  }
}
.signature-pad__typed-input:focus-visible {
  border-color: var(--office-ring);
  box-shadow: 0 0 0 3px color-mix(in oklab, var(--office-ring) 24%, transparent);
}
.signature-pad__typed-input:disabled {
  opacity: 0.64;
}
.signature-pad__typed-preview {
  display: grid;
  min-height: 88px;
  place-items: center;
  border: 1px dashed color-mix(in oklab, oklch(62.3% .214 259.815) 70%, transparent);
  border-radius: 3px;
  background: #fff;
  color: #020617;
  font-family: var(--office-font-handwriting, "Segoe Script", "Brush Script MT", cursive);
  font-size: 28px;
  line-height: 40px;
  text-align: center;
  overflow-wrap: anywhere;
  padding: var(--office-ui-space-2);
}
.signature-pad__typed-preview.is-empty {
  color: color-mix(in oklab, #020617 45%, transparent);
  font-family: var(--office-font-sans);
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
}
.signature-pad__controls {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: space-between;
  gap: var(--office-ui-space-2);
  margin-top: var(--office-ui-space-3);
}
.signature-pad__controls-end {
  display: flex;
  flex-direction: column-reverse;
  gap: var(--office-ui-space-2);
  margin-left: auto;
}
@media (min-width: 640px) {
  .signature-pad__controls-end {
    flex-direction: row;
  }
}
.signature-pad__control-icon {
  width: 16px;
  height: 16px;
}
</style>
