<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, watch } from "vue"

export interface SignaturePadProps {
  width?: number
  height?: number
  penColor?: string
  backgroundColor?: string
  className?: string
}

const props = withDefaults(defineProps<SignaturePadProps>(), {
  width: 400,
  height: 200,
  penColor: "#000000",
  backgroundColor: "transparent",
  className: "",
})

const emit = defineEmits<{
  "update:signature": [dataUrl: string | null]
}>()

const canvasRef = ref<HTMLCanvasElement | null>(null)
let ctx: CanvasRenderingContext2D | null = null
let drawing = false
let empty = true

function setupCanvas() {
  const canvas = canvasRef.value
  if (!canvas) return
  const ratio = window.devicePixelRatio || 1
  canvas.width = props.width * ratio
  canvas.height = props.height * ratio
  canvas.style.width = `${props.width}px`
  canvas.style.height = `${props.height}px`
  ctx = canvas.getContext("2d")
  if (!ctx) return
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0)
  ctx.lineCap = "round"
  ctx.lineJoin = "round"
  ctx.lineWidth = 2
  ctx.strokeStyle = props.penColor
  if (props.backgroundColor !== "transparent") {
    ctx.fillStyle = props.backgroundColor
    ctx.fillRect(0, 0, props.width, props.height)
  }
}

function point(e: PointerEvent) {
  const rect = canvasRef.value!.getBoundingClientRect()
  return { x: e.clientX - rect.left, y: e.clientY - rect.top }
}

function emitSignature() {
  if (!canvasRef.value || empty) emit("update:signature", null)
  else emit("update:signature", canvasRef.value.toDataURL("image/png"))
}

function onPointerDown(e: PointerEvent) {
  if (!ctx || !canvasRef.value) return
  drawing = true
  canvasRef.value.setPointerCapture(e.pointerId)
  const p = point(e)
  ctx.beginPath()
  ctx.moveTo(p.x, p.y)
}

function onPointerMove(e: PointerEvent) {
  if (!drawing || !ctx) return
  const p = point(e)
  ctx.lineTo(p.x, p.y)
  ctx.stroke()
  empty = false
}

function onPointerUp(e: PointerEvent) {
  if (!drawing) return
  drawing = false
  canvasRef.value?.releasePointerCapture(e.pointerId)
  emitSignature()
}

function clearPad() {
  if (!ctx) return
  ctx.clearRect(0, 0, props.width, props.height)
  if (props.backgroundColor !== "transparent") {
    ctx.fillStyle = props.backgroundColor
    ctx.fillRect(0, 0, props.width, props.height)
  }
  empty = true
  emitSignature()
}

watch(() => props.penColor, (color) => {
  if (ctx) ctx.strokeStyle = color
})

watch(() => props.backgroundColor, () => {
  clearPad()
})

onMounted(setupCanvas)
onBeforeUnmount(() => { ctx = null })

defineExpose({
  clear: clearPad,
  isEmpty: () => empty,
  toDataURL: (type?: string, encoderOptions?: number) => canvasRef.value?.toDataURL(type, encoderOptions) ?? "",
})
</script>

<template>
  <div :class="['inline-block', props.className]">
    <canvas
      ref="canvasRef"
      class="block rounded-md border border-border bg-background"
      :style="{ touchAction: 'none' }"
      @pointerdown="onPointerDown"
      @pointermove="onPointerMove"
      @pointerup="onPointerUp"
      @pointercancel="onPointerUp"
    />
    <div class="mt-2 flex items-center gap-2">
      <button
        type="button"
        class="rounded-md border border-border bg-background px-3 py-1 text-xs text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground"
        @click="clearPad"
      >
        Clear
      </button>
    </div>
  </div>
</template>
