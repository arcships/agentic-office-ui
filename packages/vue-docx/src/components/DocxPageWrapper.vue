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
import type { LayoutPage } from "@extend-ai/docx-core"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    page: LayoutPage
    pageWidthPx?: number
    pageHeightPx?: number
    isActivePage?: boolean
    margin?: number
  }>(),
  {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    isActivePage: false,
    margin: 72,
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
  if (!wrapperRef.value) return
  emit("measure", { heightPx: wrapperRef.value.scrollHeight })
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
