<script setup lang="ts">
import { ref, onMounted, onBeforeUnmount, nextTick } from "vue"

export interface TooltipProps {
  content?: string
  side?: "top" | "bottom" | "left" | "right"
  align?: "start" | "center" | "end"
  delayMs?: number
  className?: string
}

const props = withDefaults(defineProps<TooltipProps>(), {
  content: "",
  side: "top",
  align: "center",
  delayMs: 300,
  className: "",
})

const triggerRef = ref<HTMLElement | null>(null)
const tooltipRef = ref<HTMLElement | null>(null)
const isVisible = ref(false)
const floatingStyle = ref<Record<string, string>>({})
let showTimeout: ReturnType<typeof setTimeout> | null = null
let hideTimeout: ReturnType<typeof setTimeout> | null = null

function computePosition() {
  const trigger = triggerRef.value
  const tooltip = tooltipRef.value
  if (!trigger || !tooltip) return

  const gap = 8
  const margin = 8
  const triggerRect = trigger.getBoundingClientRect()
  const tooltipRect = tooltip.getBoundingClientRect()
  let top = 0
  let left = 0

  if (props.side === "bottom") top = triggerRect.bottom + gap
  else if (props.side === "left" || props.side === "right") top = triggerRect.top + triggerRect.height / 2 - tooltipRect.height / 2
  else top = triggerRect.top - tooltipRect.height - gap

  if (props.side === "right") left = triggerRect.right + gap
  else if (props.side === "left") left = triggerRect.left - tooltipRect.width - gap
  else if (props.align === "start") left = triggerRect.left
  else if (props.align === "end") left = triggerRect.right - tooltipRect.width
  else left = triggerRect.left + triggerRect.width / 2 - tooltipRect.width / 2

  left = Math.max(margin, Math.min(left, window.innerWidth - tooltipRect.width - margin))
  top = Math.max(margin, Math.min(top, window.innerHeight - tooltipRect.height - margin))

  floatingStyle.value = {
    position: "fixed",
    top: `${top}px`,
    left: `${left}px`,
  }
}

function show() {
  if (hideTimeout) {
    clearTimeout(hideTimeout)
    hideTimeout = null
  }
  if (!props.content) return
  showTimeout = setTimeout(() => {
    isVisible.value = true
    nextTick(computePosition)
  }, props.delayMs)
}

function hide() {
  if (showTimeout) {
    clearTimeout(showTimeout)
    showTimeout = null
  }
  hideTimeout = setTimeout(() => {
    isVisible.value = false
  }, 100)
}

function onTriggerEnter() {
  show()
}

function onTriggerLeave() {
  hide()
}

function onTooltipEnter() {
  if (hideTimeout) {
    clearTimeout(hideTimeout)
    hideTimeout = null
  }
}

function onTooltipLeave() {
  hide()
}

onMounted(() => {
  const el = triggerRef.value
  if (!el) return
  el.addEventListener("mouseenter", onTriggerEnter)
  el.addEventListener("mouseleave", onTriggerLeave)
  el.addEventListener("focusin", onTriggerEnter)
  el.addEventListener("focusout", onTriggerLeave)
  window.addEventListener("scroll", computePosition, true)
  window.addEventListener("resize", computePosition)
})

onBeforeUnmount(() => {
  const el = triggerRef.value
  if (el) {
    el.removeEventListener("mouseenter", onTriggerEnter)
    el.removeEventListener("mouseleave", onTriggerLeave)
    el.removeEventListener("focusin", onTriggerEnter)
    el.removeEventListener("focusout", onTriggerLeave)
  }
  window.removeEventListener("scroll", computePosition, true)
  window.removeEventListener("resize", computePosition)
  if (showTimeout) clearTimeout(showTimeout)
  if (hideTimeout) clearTimeout(hideTimeout)
})
</script>

<template>
  <div class="tooltip-trigger-wrap">
    <span ref="triggerRef" class="tooltip-trigger">
      <slot />
    </span>
    <Teleport to="body">
      <Transition name="tooltip-fade">
        <div
          v-if="isVisible"
          ref="tooltipRef"
          role="tooltip"
          :class="['tooltip-content', props.className]"
          :style="floatingStyle"
          @mouseenter="onTooltipEnter"
          @mouseleave="onTooltipLeave"
        >
          {{ props.content }}
        </div>
      </Transition>
    </Teleport>
  </div>
</template>

<style scoped>
.tooltip-trigger-wrap,
.tooltip-trigger { display: inline-flex; }
.tooltip-content {
  z-index: 50;
  max-width: min(280px, calc(100vw - 16px));
  overflow-wrap: anywhere;
  border: 1px solid var(--border, #e5e7eb);
  border-radius: var(--radius-md, 8px);
  background: var(--popover, #fff);
  color: var(--popover-foreground, #111827);
  box-shadow: var(--shadow-md, 0 4px 6px -1px rgb(0 0 0 / 0.1));
  font-size: 13px;
  line-height: 1.35;
  padding: 6px 10px;
  pointer-events: auto;
}
.tooltip-fade-enter-active { transition: opacity 0.15s ease, transform 0.15s ease; }
.tooltip-fade-leave-active { transition: opacity 0.1s ease, transform 0.1s ease; }
.tooltip-fade-enter-from,
.tooltip-fade-leave-to { opacity: 0; transform: scale(0.96); }
.tooltip-fade-enter-to,
.tooltip-fade-leave-from { opacity: 1; transform: scale(1); }
</style>
