<template>
  <div
    ref="wrapperRef"
    class="docx-page-wrapper"
    :class="{
      'docx-page-wrapper--active': isActivePage,
    }"
    :data-docx-page-wrapper="true"
    :data-page="page.number"
    :style="wrapperStyle"
  >
    <slot />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted } from "vue"
import {
  resolveClippedBodyRenderedBottomPx,
  resolveClippedMeasuredHeightPx,
} from "@arcships/docx-core"
import type { LayoutPage } from "@arcships/docx-core"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    page: LayoutPage
    pageWidthPx?: number
    pageHeightPx?: number
    isActivePage?: boolean
    margin?: number
    /** Zoom the wrapper (or its ancestors) is rendered at; rects are unscaled. */
    zoomFactor?: number
  }>(),
  {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    isActivePage: false,
    margin: 72,
    zoomFactor: 1,
  }
)

const emit = defineEmits<{
  measure: [event: { heightPx: number }]
}>()

// ── Refs ───────────────────────────────────────────────────────────
const wrapperRef = ref<HTMLElement | undefined>()

// ── Computed ───────────────────────────────────────────────────────
const wrapperStyle = computed(() => ({
  boxSizing: "border-box" as const,
  width: `${props.pageWidthPx}px`,
  minHeight: `${props.pageHeightPx}px`,
  padding: `${props.margin}px`,
  background: props.isActivePage ? "#fefefe" : "#fff",
  border: props.isActivePage ? "2px solid #3b82f6" : "1px solid #d4d4d4",
  boxShadow: props.isActivePage
    ? "0 8px 24px rgba(59, 130, 246, 0.12)"
    : "0 8px 24px rgba(0, 0, 0, 0.08)",
  marginBottom: "16px",
  position: "relative" as const,
  transition: "border-color 150ms ease, box-shadow 150ms ease",
  display: "grid",
  gap: "8px",
  alignContent: "start",
}))

// ── Measurement ────────────────────────────────────────────────────
let resizeObserver: ResizeObserver | null = null

function measure(): void {
  const wrapper = wrapperRef.value
  if (!wrapper) return
  // `scrollHeight` counts content a clipping ancestor already cut off, so a
  // wrapper whose cell or fixed-height container clips overflowing paragraphs
  // reports a taller page than it renders. Prefer the visible bottom, keep
  // `offsetHeight` as the fallback, and never report less than one paper.
  const heightPx = resolveClippedMeasuredHeightPx({
    fallbackHeightPx: wrapper.offsetHeight,
    minimumHeightPx: props.pageHeightPx,
    clippedContentBottomPx: resolveClippedBodyRenderedBottomPx(wrapper, {
      zoomFactor: props.zoomFactor,
    }),
  })
  if (heightPx !== undefined) {
    emit("measure", { heightPx })
  }
}

onMounted(() => {
  if (wrapperRef.value) {
    resizeObserver = new ResizeObserver(() => measure())
    resizeObserver.observe(wrapperRef.value)
    measure()
  }
})

onUnmounted(() => {
  resizeObserver?.disconnect()
  resizeObserver = null
})
</script>

<style scoped>
.docx-page-wrapper {
  box-sizing: border-box;
  background: #fff;
  margin-bottom: 16px;
  display: grid;
  gap: 8px;
  align-content: start;
}
</style>
