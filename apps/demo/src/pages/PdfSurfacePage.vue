<template>
  <div class="page" data-testid="pdf-surface-page">
    <header class="page-header">
      <div>
        <h2>PDF Surface — 最小嵌入组件</h2>
        <p class="desc">
          垂直滚动浏览全部页面，Surface 内无 Toolbar / 缩略图 / 搜索栏 / 旋转。
          下方宿主工具栏通过 expose 控制搜索、zoom 和 scrollToPage。
        </p>
      </div>
    </header>

    <div class="toolbar-zone">
      <label class="ctrl">
        示例 PDF
        <select v-model="selectedSample" data-testid="pdf-surface-sample-select" @change="loadSelectedSample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">
            {{ sample.label }}
          </option>
        </select>
      </label>
      <button data-testid="pdf-surface-load-sample" @click="loadSelectedSample">加载示例</button>
      <label class="ctrl">
        <input ref="fileInputRef" data-testid="pdf-surface-file-input" type="file" accept=".pdf" @change="onFileChange" />
        选择 PDF
      </label>

      <span class="sep" />

      <div class="ctrl">
        <button :disabled="!ready" @click="surfaceRef?.scrollToPage(Math.max(0, visiblePage - 1))">◀</button>
        <span class="page-value">{{ visiblePage + 1 }} / {{ numPages || "—" }}</span>
        <button :disabled="!ready" @click="surfaceRef?.scrollToPage(Math.min((numPages || 1) - 1, visiblePage + 1))">▶</button>
      </div>

      <span class="sep" />

      <div class="ctrl">
        <button :disabled="!ready || zoom <= 0.5" @click="setZoom(zoom - 0.25)">−</button>
        <span class="zoom-value">{{ Math.round(zoom * 100) }}%</span>
        <button :disabled="!ready || zoom >= 2" @click="setZoom(zoom + 0.25)">+</button>
      </div>
      <span class="sep" />

      <label class="ctrl">
        <input v-model="fitWidth" data-testid="pdf-surface-fit-width" type="checkbox" />
        自适应
      </label>

      <span class="sep" />

      <OfficeFindBar
        :query="searchQuery"
        :status="searchState.status"
        :active-index="searchState.activeIndex"
        :result-count="searchState.matches.length"
        :disabled="!ready"
        placeholder="搜索 PDF"
        @update:query="onSearchQuery"
        @previous="void surfaceRef?.searchPrevious()"
        @next="void surfaceRef?.searchNext()"
        @close="closeSearch"
      />
    </div>

    <div class="surface-container" data-testid="pdf-surface-container">
      <PdfSurface
        v-if="source"
        ref="surfaceRef"
        :key="surfaceKey"
        :source="source"
        :url-policy="urlPolicy"
        :zoom="fitWidth ? undefined : zoom"
        :fit-width="fitWidth"
        :is-dark="isDark"
        locale="zh-CN"
        class="pdf-surface-embed"
        show-status-bar
        show-annotation-dialog
        @document-load-success="onLoadSuccess"
        @document-load-error="onLoadError"
        @visible-page-change="visiblePage = $event"
        @context-menu="onContextMenu"
        @selection-change="onSelectionChange"
        @search-state-change="searchState = $event"
        @update:zoom="zoom = $event"
      />
      <div v-else class="empty" data-testid="pdf-surface-empty">
        <p>选择一个 PDF 文件以查看垂直滚动渲染效果。</p>
      </div>
    </div>

    <div class="status-grid">
      <div><strong>状态：</strong><span data-testid="pdf-surface-status" :data-state="status">{{ status }}</span></div>
      <div><strong>文件：</strong>{{ source ? "已加载" : "未选择" }}</div>
      <div><strong>页数：</strong>{{ numPages || "—" }}</div>
      <div><strong>当前页：</strong>{{ visiblePage + 1 }}</div>
      <div><strong>选中：</strong>{{ selectionInfo }}</div>
      <div><strong>右键：</strong>{{ contextMenuInfo }}</div>
    </div>

    <div
      v-if="loadError"
      class="error"
      data-testid="pdf-surface-page-load-error"
      :data-error-code="loadError.code"
    >
      {{ loadError.code }}: {{ loadError.message }}
    </div>
    <div v-if="searchError" class="error" data-testid="pdf-surface-search-error">
      {{ searchError }}
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, shallowRef } from "vue"
import type { PdfLoadError, PdfSearchState, PdfSelectionState, PdfSource, PdfUrlPolicy } from "@arcships/vue-pdf"
import { PdfSurface } from "@arcships/vue-pdf"
import { OfficeFindBar } from "@arcships/vue-ui"
import { readDemoTheme } from "../lib/demo-visual-settings"

const source = ref<PdfSource>()
const numPages = ref(0)
const visiblePage = ref(0)
const ready = ref(false)
const status = ref<"idle" | "loading" | "ready" | "error">("idle")
const loadError = ref<PdfLoadError | null>(null)
const searchError = ref("")
const urlPolicy: PdfUrlPolicy = {
  baseUrl: globalThis.location.href,
  allowRelativeUrl: true,
  allowedProtocols: ["https:"],
  allowedOrigins: [globalThis.location.origin],
  allowHttpOnLocalhost: true,
}
const zoom = ref(1)
const isDark = readDemoTheme() === "dark"
const fitWidth = ref(false)
const loadCounter = ref(0)
const fileInputRef = ref<HTMLInputElement>()
const surfaceRef = ref<InstanceType<typeof PdfSurface>>()
const contextMenuInfo = ref("—")
const selectionInfo = ref("—")
const searchQuery = ref("")
const searchState = shallowRef<PdfSearchState>({
  status: "idle",
  query: "",
  matches: [],
  activeIndex: -1,
})

const surfaceKey = computed(() => `pdf-${loadCounter.value}`)

const samples = [
  { file: "sample.pdf", label: "文本 PDF" },
  { file: "scanned-invoice.pdf", label: "扫描发票（无文本层）" },
  { file: "rotated-pages.pdf", label: "旋转页面" },
  { file: "cjk-report.pdf", label: "中文文本 PDF" },
  { file: "encrypted.pdf", label: "加密 PDF" },
  { file: "corrupted.pdf", label: "损坏文件（失败用例）" },
]
const selectedSample = ref(samples[0].file)

function loadSelectedSample(): void {
  resetLoadState()
  source.value = { kind: "url", url: `/samples/${selectedSample.value}` }
  loadCounter.value++
}

function resetLoadState(): void {
  ready.value = false
  loadError.value = null
  searchError.value = ""
  status.value = "loading"
  numPages.value = 0
  visiblePage.value = 0
}

function onFileChange(event: Event): void {
  const file = (event.target as HTMLInputElement).files?.[0] ?? null
  if (!file) return
  resetLoadState()
  source.value = { kind: "blob", blob: file }
  loadCounter.value++
}

function onLoadSuccess(count: number): void {
  numPages.value = count
  ready.value = true
  loadError.value = null
  status.value = "ready"
}

function onLoadError(error: PdfLoadError): void {
  ready.value = false
  loadError.value = error
  status.value = "error"
}

function setZoom(value: number): void {
  zoom.value = Math.min(2, Math.max(0.5, value))
}

function onSearchQuery(query: string): void {
  searchQuery.value = query
  const task = query.trim()
    ? surfaceRef.value?.search(query)
    : (surfaceRef.value?.clearSearch(), undefined)
  void task?.catch((cause: unknown) => {
    if (!(cause instanceof Error) || cause.name !== "AbortError") {
      searchError.value = cause instanceof Error ? cause.message : String(cause)
    }
  })
}

function closeSearch(): void {
  searchQuery.value = ""
  surfaceRef.value?.clearSearch()
}

function onSelectionChange(selection: PdfSelectionState): void {
  if (selection.kind === "none") selectionInfo.value = "—"
  else if (selection.kind === "text") selectionInfo.value = selection.text.slice(0, 40)
  else selectionInfo.value = selection.reason === "unicode-map" ? "文字映射不可用" : "没有可复制文字"
}

function onContextMenu(ctx: { pageIndex: number; clientX: number; clientY: number }): void {
  contextMenuInfo.value = `第 ${ctx.pageIndex + 1} 页`
}
</script>

<style scoped>
.page { padding: 16px; max-width: 1440px; margin: 0 auto; width: 100%; min-width: 0; display: flex; flex-direction: column; height: calc(100vh - 60px); }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 0; font-size: 13px; line-height: 1.5; }

.toolbar-zone {
  display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  padding: 10px 12px; margin: 12px 0;
  border: 1px solid var(--border); border-radius: var(--radius);
  background: var(--background); flex-shrink: 0;
}
.ctrl { display: flex; align-items: center; gap: 4px; font-size: 13px; white-space: nowrap; }
.ctrl input[type="file"] { font-size: 12px; }
.toolbar-zone button {
  padding: 4px 10px; border: 1px solid var(--border); border-radius: 4px;
  background: var(--background); cursor: pointer; font-size: 13px;
}
.toolbar-zone button:disabled { opacity: .4; cursor: default; }
.page-value, .zoom-value { min-width: 60px; text-align: center; font-variant-numeric: tabular-nums; }
.sep { width: 1px; height: 20px; background: var(--border); }

.surface-container {
  flex: 1; min-height: 0; border: 1px solid var(--border);
}
.pdf-surface-embed { height: 100%; }
.empty { display: flex; align-items: center; justify-content: center; height: 200px; color: var(--muted-foreground); }
.error { color: #dc2626; font-size: 13px; margin: 8px 0 0; }

.status-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 8px;
  padding: 12px; margin-top: 8px; border: 1px solid var(--border); border-radius: var(--radius);
  background: var(--muted); font-size: 13px; flex-shrink: 0;
}
</style>
