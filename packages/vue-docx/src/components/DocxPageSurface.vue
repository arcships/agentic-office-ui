<template>
  <div
    ref="pageSurfaceRef"
    data-docx-page-surface="true"
    class="docx-page-surface"
    :style="surfaceStyle"
    @click="onSurfaceClick"
  >
    <!-- Cover layer (full-page background images) -->
    <div
      v-if="coverEntries.length > 0"
      data-docx-page-cover-layer="true"
      class="docx-page-cover-layer"
    >
      <img
        v-for="cover of coverEntries"
        :key="cover.key"
        :src="cover.src"
        alt=""
        aria-hidden="true"
        draggable="false"
        class="docx-page-cover-image"
        :style="cover.style"
      />
    </div>

    <!-- Border overlay -->
    <div
      v-if="borderStyle"
      data-docx-page-border-overlay="true"
      class="docx-page-border-overlay"
      :style="borderStyle"
    />

    <!-- Header -->
    <DocxPageHeader
      v-if="hasHeader"
      :page-index="pageIndex"
      :header-nodes="[]"
      :page-layout="pageLayout"
      :page-content-width-px="pageContentWidthPx"
      :theme="theme"
      :controller="controller"
    />

    <!-- Body -->
    <DocxPageBody
      :page-index="pageIndex"
      :page-node-segments="pageNodeSegments"
      :page-layout="pageLayout"
      :page-content-width-px="pageContentWidthPx"
      :controller="controller"
      :editable="editable"
      :theme="theme"
      :tracked-changes-enabled="trackedChangesEnabled"
      :comments-enabled="commentsEnabled"
    />

    <!-- Footer -->
    <DocxPageFooter
      v-if="hasFooter"
      :page-index="pageIndex"
      :footer-nodes="[]"
      :page-layout="pageLayout"
      :page-content-width-px="pageContentWidthPx"
      :theme="theme"
      :controller="controller"
    />

    <!-- Floating image layer -->
    <DocxImageLayer
      :page-index="pageIndex"
      :page-width-px="pageWidthPx"
      :page-height-px="pageLayout.pageHeightPx"
      :controller="controller"
    />

    <!-- Form field layer -->
    <DocxFormFieldLayer
      v-if="hasFormFields"
      :page-index="pageIndex"
      :controller="controller"
    />

    <!-- Tracked change / comment gutter -->
    <DocxTrackedChangeGutter
      v-if="showGutter"
      :page-index="pageIndex"
      :page-layout="pageLayout"
      :controller="controller"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onMounted, onUnmounted } from "vue"
import type {
  DocxEditorController,
  DocumentPageNodeSegment,
  ImageRunNode,
} from "@extend-ai/docx-core"
import {
  resolveDocumentLayout,
  resolveRenderableImageSource,
  pageMarginPaddingStyle,
} from "@extend-ai/docx-core"
import DocxPageHeader from "./DocxPageHeader.vue"
import DocxPageFooter from "./DocxPageFooter.vue"
import DocxPageBody from "./DocxPageBody"
import DocxImageLayer from "./DocxImageLayer.vue"
import DocxFormFieldLayer from "./DocxFormFieldLayer.vue"
import DocxTrackedChangeGutter from "./DocxTrackedChangeGutter.vue"

// ── Types ──────────────────────────────────────────────────────────
interface DocxPageSurfaceProps {
  pageIndex: number
  pageLayout: ReturnType<typeof resolveDocumentLayout>
  pageNodeSegments: DocumentPageNodeSegment[]
  controller: DocxEditorController
  editable?: boolean
  pageWidthPx: number
  pageContentWidthPx: number
  trackedChangesEnabled?: boolean
  commentsEnabled?: boolean
  theme?: "light" | "dark"
}

interface CoverEntry {
  key: string
  src: string
  style: Record<string, string>
}

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<DocxPageSurfaceProps>(),
  {
    editable: true,
    trackedChangesEnabled: false,
    commentsEnabled: false,
    theme: "light" as const,
  }
)

const emit = defineEmits<{
  measure: [event: { heightPx: number }]
}>()

// ── Refs ───────────────────────────────────────────────────────────
const pageSurfaceRef = ref<HTMLElement | undefined>()

// ── Computed ───────────────────────────────────────────────────────
const hasHeader = computed(() => false) // Simplified: no header nodes yet
const hasFooter = computed(() => false) // Simplified: no footer nodes yet
const hasFormFields = computed(() => false) // Simplified: form field detection

const showGutter = computed(
  () => props.trackedChangesEnabled || props.commentsEnabled
)

const coverEntries = computed<CoverEntry[]>(() => {
  // Collect full-page cover background images from paragraph segments
  // Simplified: empty for now — full implementation requires image run detection
  return []
})

const borderStyle = computed<Record<string, string> | undefined>(() => {
  // Page border overlay — simplified: no border by default
  return undefined
})

const surfaceStyle = computed(() => ({
  width: `${props.pageWidthPx}px`,
  minHeight: `${props.pageLayout.pageHeightPx}px`,
  backgroundColor: props.theme === "dark" ? "#1f2937" : "#ffffff",
  color: props.theme === "dark" ? "#f9fafb" : "#111827",
  position: "relative" as const,
  boxSizing: "border-box" as const,
  border: "1px solid #d4d4d4",
  boxShadow: "0 8px 24px rgba(0, 0, 0, 0.08)",
  ...pageMarginPaddingStyle(props.pageLayout.marginsPx ?? { top: 72, bottom: 72, left: 72, right: 72 }),
}))

// ── Surface click → select paragraph ───────────────────────────────
function onSurfaceClick(event: MouseEvent): void {
  // Only handle clicks directly on the surface (not on child interactive elements)
  if (event.target === pageSurfaceRef.value || (event.target as HTMLElement)?.dataset?.docxPageSurface === "true") {
    // Don't steal focus from form fields or editable elements
  }
}

// ── Measure page height ────────────────────────────────────────────
let resizeObserver: ResizeObserver | null = null

function measurePage(): void {
  if (!pageSurfaceRef.value) return
  const rect = pageSurfaceRef.value.getBoundingClientRect()
  emit("measure", { heightPx: rect.height })
}

onMounted(() => {
  if (pageSurfaceRef.value) {
    resizeObserver = new ResizeObserver(() => {
      measurePage()
    })
    resizeObserver.observe(pageSurfaceRef.value)
    measurePage()
  }
})

onUnmounted(() => {
  if (resizeObserver) {
    resizeObserver.disconnect()
    resizeObserver = null
  }
})
</script>

<style scoped>
.docx-page-surface {
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
}
.docx-page-cover-layer {
  position: absolute;
  left: 0;
  top: 0;
  right: 0;
  bottom: 0;
  overflow: clip;
  pointer-events: none;
  z-index: 0;
}
.docx-page-cover-image {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  display: block;
  object-fit: cover;
}
.docx-page-border-overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
  z-index: 5;
}
</style>
