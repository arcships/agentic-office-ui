<script setup lang="ts">
import { ref, computed, onMounted, onBeforeUnmount } from "vue"

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
let showTimeout: ReturnType<typeof setTimeout> | null = null
let hideTimeout: ReturnType<typeof setTimeout> | null = null

const sideClasses: Record<string, string> = {
  top: "bottom-full left-1/2 -translate-x-1/2 mb-1.5",
  bottom: "top-full left-1/2 -translate-x-1/2 mt-1.5",
  left: "right-full top-1/2 -translate-y-1/2 mr-1.5",
  right: "left-full top-1/2 -translate-y-1/2 ml-1.5",
}

const alignClasses: Record<string, Record<string, string>> = {
  start: {
    top: "left-0 translate-x-0",
    bottom: "left-0 translate-x-0",
    left: "top-0 translate-y-0",
    right: "top-0 translate-y-0",
  },
  center: {
    top: "",
    bottom: "",
    left: "",
    right: "",
  },
  end: {
    top: "left-auto right-0 translate-x-0",
    bottom: "left-auto right-0 translate-x-0",
    left: "top-auto bottom-0 translate-y-0",
    right: "top-auto bottom-0 translate-y-0",
  },
}

const tooltipClass = computed(() => {
  const base = sideClasses[props.side] ?? sideClasses.top
  const align = alignClasses[props.align]?.[props.side] ?? ""
  return [base, align].filter(Boolean).join(" ")
})

function show() {
  if (hideTimeout) {
    clearTimeout(hideTimeout)
    hideTimeout = null
  }
  if (!props.content) return
  showTimeout = setTimeout(() => {
    isVisible.value = true
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
})

onBeforeUnmount(() => {
  const el = triggerRef.value
  if (!el) return
  el.removeEventListener("mouseenter", onTriggerEnter)
  el.removeEventListener("mouseleave", onTriggerLeave)
  el.removeEventListener("focusin", onTriggerEnter)
  el.removeEventListener("focusout", onTriggerLeave)
  if (showTimeout) clearTimeout(showTimeout)
  if (hideTimeout) clearTimeout(hideTimeout)
})
</script>

<template>
  <div class="relative inline-flex">
    <span ref="triggerRef" class="inline-flex">
      <slot />
    </span>
    <Teleport to="body">
      <Transition
        enter-active-class="transition ease-out duration-150"
        enter-from-class="opacity-0 scale-95"
        enter-to-class="opacity-100 scale-100"
        leave-active-class="transition ease-in duration-100"
        leave-from-class="opacity-100 scale-100"
        leave-to-class="opacity-0 scale-95"
      >
        <div
          v-if="isVisible"
          ref="tooltipRef"
          role="tooltip"
          :class="[
            'absolute z-50 overflow-hidden rounded-md border border-border bg-popover px-3 py-1.5 text-sm text-popover-foreground shadow-md',
            tooltipClass,
            props.className,
          ]"
          @mouseenter="onTooltipEnter"
          @mouseleave="onTooltipLeave"
        >
          {{ props.content }}
        </div>
      </Transition>
    </Teleport>
  </div>
</template>
