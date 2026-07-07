<template>
  <div :class="cn('docx-viewer', props.className)" :style="props.style">
    <!-- Empty state -->
    <div v-if="!model && !file && !props.emptyState" class="flex items-center justify-center h-full min-h-[200px] text-muted-foreground">
      <p>No document loaded</p>
    </div>

    <!-- Custom empty state -->
    <component :is="props.emptyState" v-if="!model && !file && props.emptyState" />

    <!-- Loading state -->
    <div v-if="isLoading" class="flex items-center justify-center h-full min-h-[200px]">
      <slot name="loading">
        <p class="text-muted-foreground">Loading document...</p>
      </slot>
    </div>

    <!-- Error state -->
    <div v-if="loadError && !isLoading" class="flex items-center justify-center h-full min-h-[200px] text-red-600">
      <p>{{ loadError.message }}</p>
    </div>

    <!-- Document content -->
    <div v-if="model && !isLoading && !loadError" class="docx-viewer-content">
      <div
        v-for="page in pages"
        :key="page.number"
        class="docx-page"
        :style="{
          width: `${pageWidth}px`,
          minHeight: `${pageHeight}px`,
          padding: `${marginTop}px ${marginRight}px ${marginBottom}px ${marginLeft}px`,
          margin: '0 auto 24px',
          background: 'white',
        }"
      >
        <div
          v-for="block in page.blocks"
          :key="block.id"
          :style="{
            width: '100%',
            minHeight: `${block.height}px`,
            marginBottom: '6px',
          }"
        >
          <!-- Paragraph block -->
          <p
            v-if="block.kind === 'paragraph'"
            :style="{
              fontSize: block.headingLevel
                ? `${headingScale(block.headingLevel) * 100}%`
                : '100%',
              fontWeight: block.headingLevel ? 'bold' : 'normal',
              fontFamily: block.headingLevel ? 'sans-serif' : 'inherit',
              marginBottom: '4px',
            }"
          >
            <template v-for="run in block.children" :key="run.id">
              <img
                v-if="run.kind === 'image' && run.src"
                :src="run.src"
                :alt="'Document image'"
                :style="{ maxWidth: `${run.widthPx}px`, maxHeight: `${run.heightPx}px` }"
              />
              <span
                v-else
                :style="runTextStyle(run)"
              >{{ run.text }}</span>
            </template>
          </p>

          <!-- Table block -->
          <table
            v-if="block.kind === 'table'"
            class="docx-table"
            style="border-collapse: collapse; width: 100%;"
          >
            <tbody>
              <tr v-for="(row, ri) in block.rows" :key="ri" :style="{ height: `${row.height}px` }">
                <td
                  v-for="(cell, ci) in row.cells"
                  :key="ci"
                  :colSpan="cell.colSpan"
                  :rowSpan="cell.rowSpan"
                  :style="{ width: `${cell.width}px`, border: '1px solid #ccc', padding: '4px', verticalAlign: 'top' }"
                >
                  <template v-for="node in cell.nodes" :key="node.id">
                    <p v-for="run in node.children" :key="run.id" :style="runTextStyle(run)">
                      {{ run.text }}
                    </p>
                  </template>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, watch, computed } from "vue"
import type { DocModel, LayoutOptions, LayoutPage, LayoutRun } from "@extend-ai/docx-core"
import {
  layoutDocument,
  resolveDocumentLayout,
  headingScale as headingScaleFn,
} from "@extend-ai/docx-core"

export interface DocxViewerProps {
  file?: ArrayBuffer
  model?: DocModel | null
  className?: string
  style?: Record<string, string | number>
  layoutOptions?: Partial<LayoutOptions>
  emptyState?: unknown // Component or vnode
}

const props = withDefaults(defineProps<DocxViewerProps>(), {})

const emit = defineEmits<{
  "load-error": [error: Error]
}>()

const isLoading = ref(false)
const loadError = ref<Error | null>(null)
const internalModel = ref<DocModel | null>(null)

const model = computed(() => props.model ?? internalModel.value)

// Watch file prop: auto-parse .docx binary
watch(
  () => props.file,
  async (f) => {
    if (!f) { internalModel.value = null; return }
    isLoading.value = true
    loadError.value = null
    try {
      const { buildDocModelFromBytes } = await import("@extend-ai/docx-core")
      const result = await buildDocModelFromBytes(f)
      internalModel.value = result.model
    } catch (e) {
      const err = e instanceof Error ? e : new Error(String(e))
      loadError.value = err
      internalModel.value = null
      emit("load-error", err)
    } finally {
      isLoading.value = false
    }
  },
  { immediate: true }
)

// Layout
const layoutMetrics = computed(() => {
  const base = model.value
    ? resolveDocumentLayout(model.value)
    : { pageWidth: 816, pageHeight: 1056, marginTop: 72, marginBottom: 72, marginLeft: 72, marginRight: 72 }
  return { ...base, ...(props.layoutOptions ?? {}) }
})

const layoutOptions = computed<Partial<LayoutOptions>>(() => ({
  ...(props.layoutOptions ?? {}),
  pageWidth: layoutMetrics.value.pageWidth,
  pageHeight: layoutMetrics.value.pageHeight,
  marginTop: layoutMetrics.value.marginTop,
  marginBottom: layoutMetrics.value.marginBottom,
  marginLeft: layoutMetrics.value.marginLeft,
  marginRight: layoutMetrics.value.marginRight,
}))

const pageWidth = computed(() => layoutMetrics.value.pageWidth)
const pageHeight = computed(() => layoutMetrics.value.pageHeight)
const marginTop = computed(() => layoutMetrics.value.marginTop)
const marginBottom = computed(() => layoutMetrics.value.marginBottom)
const marginLeft = computed(() => layoutMetrics.value.marginLeft)
const marginRight = computed(() => layoutMetrics.value.marginRight)

const pages = computed<LayoutPage[]>(() => {
  if (!model.value) return []
  return layoutDocument(model.value, layoutOptions.value)
})

// Text style helper
function runTextStyle(run: LayoutRun): Record<string, string | number> {
  const s = run.style
  if (!s) return {}
  return {
    fontWeight: s.bold ? "bold" : "normal",
    fontStyle: s.italic ? "italic" : "normal",
    textDecoration: s.underline ? "underline" : s.strike ? "line-through" : "none",
    color: s.color ?? "inherit",
    fontFamily: s.fontFamily ?? "inherit",
    fontSize: s.fontSizePt ? `${s.fontSizePt}pt` : "inherit",
  }
}

function cn(...classes: (string | undefined | null | false)[]): string {
  return classes.filter(Boolean).join(" ")
}

// Re-export heading scale for template use
const headingScale = headingScaleFn
</script>

<style scoped>
.docx-viewer {
  min-width: 0;
  background: color-mix(in oklch, var(--muted, #f5f5f5) 58%, var(--background, #fff));
}
.docx-viewer-content {
  min-width: 0;
  padding: 24px;
}
.docx-page {
  border: 1px solid color-mix(in oklch, var(--border, #e5e7eb) 75%, transparent);
  box-shadow: 0 12px 28px rgb(15 23 42 / 0.10), 0 2px 6px rgb(15 23 42 / 0.08);
}
@media (max-width: 640px) {
  .docx-viewer-content { padding: 12px; }
}
</style>
