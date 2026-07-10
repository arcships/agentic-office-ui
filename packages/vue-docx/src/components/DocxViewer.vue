<template>
  <div
    class="docx-viewer"
    :class="className"
    data-testid="docx-viewer"
    :data-state="viewerState"
  >
    <!-- Loading state -->
    <div v-if="isLoading" class="docx-viewer-loading">
      <div class="docx-viewer-loading-spinner" />
      <span>Loading DOCX...</span>
    </div>

    <!-- Error state -->
    <div
      v-else-if="error"
      class="docx-viewer-error"
      data-testid="load-error"
      :data-error-code="errorCode(error)"
    >
      <div class="docx-viewer-error-icon">⚠</div>
      <div class="docx-viewer-error-text">Failed to parse DOCX</div>
      <div v-if="errorCode(error)" class="docx-viewer-error-code">
        {{ errorCode(error) }}
      </div>
      <div class="docx-viewer-error-detail">{{ error.message }}</div>
    </div>

    <!-- Empty state -->
    <div v-else-if="!resolvedModel" class="docx-viewer-empty">
      {{ emptyState ?? "No DOCX loaded." }}
    </div>

    <!-- Viewer and editor share the same page/paragraph/table surface. -->
    <DocxDocumentSurface
      v-else
      :model="resolvedModel"
      :layout-options="layoutOptions"
      :editable="false"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onBeforeUnmount } from "vue"
import type { DocModel, DocxImportResult, DocxRuntime, DocxRuntimeLoader } from "@extend-ai/docx-core"
import { createDocxRuntime } from "@extend-ai/docx-core"
import type { LayoutOptions } from "@extend-ai/docx-core"
import DocxDocumentSurface from "./DocxDocumentSurface.vue"

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    file?: ArrayBuffer
    model?: DocModel
    className?: string
    layoutOptions?: LayoutOptions
    emptyState?: string
    runtime?: DocxRuntime
  }>(),
  { emptyState: "No DOCX loaded." }
)

const emit = defineEmits<{
  (event: "load-start"): void
  (event: "load-success", result: DocxImportResult): void
  (event: "load-error", error: Error): void
}>()

// ── State ──────────────────────────────────────────────────────────
const isLoading = ref(false)
const error = ref<Error | undefined>(undefined)
const parsedModel = ref<DocModel | undefined>(undefined)

const resolvedModel = computed(() => props.model ?? parsedModel.value)
const viewerState = computed(() => {
  if (isLoading.value) return "loading"
  if (error.value) return "error"
  if (resolvedModel.value) return "ready"
  return "idle"
})

// ── Import DOCX through one instance-owned loader ──────────────────
const ownedRuntime = createDocxRuntime()
let loaderRuntime: DocxRuntime = props.runtime ?? ownedRuntime
let runtimeLoader: DocxRuntimeLoader = loaderRuntime.createLoader()

function resolveRuntimeLoader(): DocxRuntimeLoader {
  const nextRuntime = props.runtime ?? ownedRuntime
  if (nextRuntime !== loaderRuntime) {
    runtimeLoader.dispose()
    loaderRuntime = nextRuntime
    runtimeLoader = loaderRuntime.createLoader()
  }
  return runtimeLoader
}

watch(
  () => [props.file, props.runtime, props.model] as const,
  async ([newFile], _oldValue, onCleanup) => {
    let active = true
    onCleanup(() => { active = false })
    const loader = resolveRuntimeLoader()

    if (!newFile) {
      loader.cancel()
      parsedModel.value = undefined
      error.value = undefined
      isLoading.value = false
      return
    }
    if (props.model) {
      loader.cancel()
      parsedModel.value = undefined
      error.value = undefined
      isLoading.value = false
      return
    }

    const docxFile = newFile
    isLoading.value = true
    error.value = undefined
    emit("load-start")
    try {
      const result = await loader.load(docxFile, { transferBuffer: false })
      if (!active) return
      parsedModel.value = result.model
      emit("load-success", result)
    } catch (err) {
      if (!active) return
      const normalized = err instanceof Error ? err : new Error("Unknown DOCX parse error")
      error.value = normalized
      emit("load-error", normalized)
    } finally {
      if (active) isLoading.value = false
    }
  },
  { immediate: true }
)

onBeforeUnmount(() => {
  runtimeLoader.dispose()
  ownedRuntime.dispose()
})

function errorCode(value: Error): string | undefined {
  const code = (value as Error & { code?: unknown }).code
  return typeof code === "string" ? code : undefined
}
</script>

<style scoped>
.docx-viewer {
  display: flex;
  flex-direction: column;
  width: 100%;
  min-width: 0;
  background: #f3f4f6;
}
.docx-viewer-loading {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 24px;
  color: #6b7280;
  font-size: 14px;
}
.docx-viewer-loading-spinner {
  width: 20px;
  height: 20px;
  border: 2px solid #e5e7eb;
  border-top-color: #3b82f6;
  border-radius: 50%;
  animation: docx-viewer-spin 0.6s linear infinite;
}
@keyframes docx-viewer-spin {
  to { transform: rotate(360deg); }
}
.docx-viewer-error {
  padding: 24px;
  text-align: center;
}
.docx-viewer-error-icon {
  font-size: 32px;
  margin-bottom: 8px;
}
.docx-viewer-error-text {
  color: #dc2626;
  font-weight: 600;
  margin-bottom: 4px;
}
.docx-viewer-error-code {
  color: #991b1b;
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
  font-size: 12px;
  font-weight: 600;
  margin-bottom: 4px;
}
.docx-viewer-error-detail {
  color: #6b7280;
  font-size: 13px;
}
.docx-viewer-empty {
  padding: 24px;
  color: #6b7280;
  font-size: 14px;
}
</style>
