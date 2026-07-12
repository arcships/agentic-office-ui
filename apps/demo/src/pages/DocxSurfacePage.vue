<template>
  <div class="page" data-testid="docx-surface-page">
    <header class="page-header">
      <div>
        <h2>DOCX Surface — 最小嵌入组件</h2>
        <p class="desc">
          只负责渲染分页文档面 + 批注 gutter，无 toolbar / 缩略图 / 文件加载 / 外壳。
          所有控制按钮由宿主自行实现，通过 props / expose 驱动。
        </p>
      </div>
    </header>

    <div class="toolbar-zone">
      <label class="ctrl">
        示例文档
        <select v-model="selectedSample" data-testid="docx-surface-sample-select" @change="loadSelectedSample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">{{ sample.label }}</option>
        </select>
      </label>
      <button data-testid="docx-surface-load-sample" @click="loadSelectedSample">打开</button>

      <span class="sep" />

      <label class="ctrl">
        <input v-model="showTrackedChanges" data-testid="docx-surface-show-tracked-changes" type="checkbox" />
        修订
      </label>
      <label class="ctrl">
        <input v-model="showComments" data-testid="docx-surface-show-comments" type="checkbox" />
        批注
      </label>

      <span class="sep" />

      <div class="ctrl zoom-group">
        <button :disabled="!ready" @click="zoomDown">−</button>
        <span class="zoom-value">{{ zoom }}%</span>
        <button :disabled="!ready" @click="zoomUp">+</button>
      </div>

      <span class="sep" />

      <div class="ctrl page-group">
        <button :disabled="!ready" @click="surfaceRef?.scrollToPage(0)">⏮</button>
        <button :disabled="!ready" @click="surfaceRef?.scrollToPage(Math.max(0, currentPage - 1))">◀</button>
        <span class="page-value">{{ currentPage + 1 }} / {{ totalPages }}</span>
        <button :disabled="!ready" @click="surfaceRef?.scrollToPage(Math.min(totalPages - 1, currentPage + 1))">▶</button>
        <button :disabled="!ready" @click="surfaceRef?.scrollToPage(Math.max(0, totalPages - 1))">⏭</button>
      </div>

      <span class="sep" />

      <label class="ctrl">
        <input v-model.trim="searchQuery" placeholder="搜索…" data-testid="docx-surface-search" />
      </label>
    </div>

    <p v-if="error" class="error" data-testid="docx-surface-error">{{ error }}</p>

    <div class="surface-container" data-testid="docx-surface-container">
      <DocxDocumentSurface
        v-if="model"
        ref="surfaceRef"
        :key="surfaceKey"
        :model="model"
        :layout-options="{ pageWidth: 816, pageHeight: 1056 }"
        :show-tracked-changes="showTrackedChanges"
        :show-comments="showComments"
        :zoom-scale="zoom"
        :search-query="searchQuery"
        :active-search-node-index="searchMatchNodeIndex"
        style="--docx-surface-bg: transparent; height: 72vh;"
        @page-count-change="onPageCountChange"
        @visible-page-range="onVisiblePageRange"
      />
      <div v-else class="empty" data-testid="docx-surface-empty">
        <p>选择示例文档以查看最小嵌入组件的渲染效果。</p>
      </div>
    </div>

    <div class="status-grid" data-testid="docx-surface-status">
      <div><strong>文件：</strong>{{ displayName || "未打开" }}</div>
      <div><strong>页数：</strong>{{ totalPages }}</div>
      <div><strong>当前页：</strong>{{ currentPage + 1 }}</div>
      <div><strong>执行位置：</strong>{{ sourceKind }}</div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, ref, shallowRef } from "vue"
import {
  bundledDocxWasmUrl,
  createDocxRuntime,
  type DocModel,
  type DocxImportResult,
} from "@arcships/docx-core"
import { DocxDocumentSurface } from "@arcships/vue-docx"

// ── Samples ──────────────────────────────────────────────────────────
const samples = [
  { file: "demo.docx", label: "Demo — master services agreement" },
  { file: "legal-contract.docx", label: "Legal contract — multi-page" },
  { file: "invoice-table.docx", label: "Invoice table — merged cells and totals" },
  { file: "report-with-image.docx", label: "Report with image and rich text" },
  { file: "chinese-mixed.docx", label: "Chinese + English mixed text" },
  { file: "review-comments.docx", label: "Review — 修订、批注、脚注、尾注与搜索" },
]

// ── State ────────────────────────────────────────────────────────────
const selectedSample = ref(samples[0].file)
const model = shallowRef<DocModel | null>(null)
const displayName = ref("")
const sourceKind = ref("pending")
const error = ref<string | null>(null)
const ready = ref(false)
const loadCounter = ref(0)

const showTrackedChanges = ref(true)
const showComments = ref(true)
const zoom = ref(100)
const searchQuery = ref("")

const totalPages = ref(0)
const currentPage = ref(0)
const searchMatchNodeIndex = ref<number | undefined>()

const surfaceRef = ref<InstanceType<typeof DocxDocumentSurface>>()
const surfaceKey = computed(() => `${displayName.value}-${loadCounter.value}`)

// ── Runtime ──────────────────────────────────────────────────────────
const runtime = shallowRef(createDocxRuntime({
  wasmUrl: bundledDocxWasmUrl,
}))

// ── Actions ──────────────────────────────────────────────────────────
async function loadSelectedSample(): Promise<void> {
  const fileName = selectedSample.value
  error.value = null
  ready.value = false
  totalPages.value = 0
  currentPage.value = 0
  try {
    const response = await fetch(`/samples/${encodeURIComponent(fileName)}`)
    if (!response.ok) {
      error.value = `Sample not found: ${fileName}`
      return
    }
    const buffer = await response.arrayBuffer()
    const result = await runtime.value.createLoader().load(buffer, { transferBuffer: false })
    model.value = result.model
    displayName.value = fileName
    sourceKind.value = result.source
    loadCounter.value++
  } catch (cause) {
    error.value = `Failed to load: ${String(cause)}`
  }
}

function zoomDown(): void { zoom.value = Math.max(50, zoom.value - 25) }
function zoomUp(): void { zoom.value = Math.min(200, zoom.value + 25) }

// ── Surface events ───────────────────────────────────────────────────
function onPageCountChange(count: number): void {
  totalPages.value = Math.max(0, count)
  ready.value = true
}

function onVisiblePageRange(range: { startPageIndex: number; endPageIndex: number }): void {
  if (range.endPageIndex >= range.startPageIndex) {
    currentPage.value = range.startPageIndex
  }
}

// ── Teardown ─────────────────────────────────────────────────────────
onBeforeUnmount(() => {
  runtime.value.dispose()
})

void loadSelectedSample()
</script>

<style scoped>
.page { padding: 16px; max-width: 1440px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 0; font-size: 13px; line-height: 1.5; }

/* Host-owned toolbar — not part of the surface component */
.toolbar-zone {
  display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  padding: 10px 12px; margin: 12px 0;
  border: 1px solid var(--border); border-radius: var(--radius);
  background: var(--background);
}
.ctrl { display: flex; align-items: center; gap: 4px; font-size: 13px; white-space: nowrap; }
.ctrl input[type="text"] { width: 120px; padding: 4px 8px; border: 1px solid var(--border); border-radius: 4px; }
.ctrl select { padding: 4px 8px; border: 1px solid var(--border); border-radius: 4px; }
.sep { width: 1px; height: 20px; background: var(--border); }
.toolbar-zone button {
  padding: 4px 10px; border: 1px solid var(--border); border-radius: 4px;
  background: var(--background); cursor: pointer; font-size: 13px;
}
.toolbar-zone button:disabled { opacity: .4; cursor: default; }
.zoom-group, .page-group { gap: 2px; }
.zoom-value, .page-value { min-width: 40px; text-align: center; font-variant-numeric: tabular-nums; }

/* Surface container — hosts the embedded component */
.surface-container {
  border: 1px solid var(--border);
  margin-bottom: 12px;
  /* deliberately NO border-radius — surface should have no decorative chrome */
}

.error { color: #ef4444; font-size: 13px; margin-bottom: 8px; }
.empty { display: flex; align-items: center; justify-content: center; height: 240px; color: var(--muted-foreground); }

.status-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 8px;
  padding: 12px; border: 1px solid var(--border); border-radius: var(--radius);
  background: var(--muted); font-size: 13px;
}
</style>
