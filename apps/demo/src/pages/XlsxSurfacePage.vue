<template>
  <div class="page" data-testid="xlsx-surface-page">
    <header class="page-header">
      <div>
        <h2>Excel/CSV Surface — 最小嵌入组件</h2>
        <p class="desc">
          只负责渲染表格网格 + 图表/图片/绘图/选区叠加层 + 右键菜单，无 Toolbar / Ribbon / FormulaBar。
          宿主自行管理 SheetTabs、文件加载、只读切换。
        </p>
      </div>
    </header>

    <div class="toolbar-zone">
      <label class="ctrl">
        示例工作簿
        <select v-model="selectedSample" data-testid="xlsx-surface-sample-select" @change="loadSelectedSample">
          <option v-for="sample in samples" :key="sample.file" :value="sample.file">{{ sample.label }}</option>
        </select>
      </label>
      <button data-testid="xlsx-surface-load-sample" @click="loadSelectedSample">打开</button>

      <span class="sep" />

      <label class="ctrl">
        <input v-model="readOnly" data-testid="xlsx-surface-read-only" type="checkbox" />
        只读
      </label>

      <span class="sep" />

      <div class="ctrl">
        <button :disabled="zoom <= 0.5" @click="zoom = Math.max(0.5, zoom - 0.25)">−</button>
        <span class="zoom-value">{{ Math.round(zoom * 100) }}%</span>
        <button :disabled="zoom >= 2" @click="zoom = Math.min(2, zoom + 0.25)">+</button>
      </div>

      <span class="sep" />

      <label class="ctrl">
        <input ref="fileInputRef" data-testid="xlsx-surface-file-input" type="file" accept=".xlsx,.xls,.xlsb,.xlsm,.xltx,.xltm,.csv,text/csv" @change="onFileChange" />
        本地文件
      </label>
    </div>

    <p v-if="error" class="error" data-testid="xlsx-surface-error">{{ error }}</p>

    <div class="surface-container" data-testid="xlsx-surface-container">
      <div v-if="loading" class="surface-overlay">加载中…</div>
      <XlsxSurfaceHost
        v-if="surfaceKey"
        :key="surfaceKey"
        :file="fileBuffer"
        :file-name="displayName"
        :read-only="readOnly"
        :zoom="zoom"
        @cellDoubleClick="onCellDoubleClick"
        @contextMenu="onContextMenu"
        @selectionChange="onSelectionChange"
        @update:zoom="zoom = $event"
      />
      <div v-else class="empty" data-testid="xlsx-surface-empty">
        <p>选择示例或打开本地 Excel、CSV 文件。</p>
      </div>
    </div>

    <div class="status-grid" data-testid="xlsx-surface-status">
      <div><strong>文件：</strong>{{ displayName || "未打开" }}</div>
      <div><strong>选中：</strong>{{ selectionInfo }}</div>
      <div><strong>右键：</strong>{{ contextMenuInfo }}</div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, defineComponent, h, ref, type PropType } from "vue"
import type { XlsxViewerController } from "@arcships/xlsx-core"
import {
  useXlsxViewerController,
  XlsxSheetSurface,
} from "@arcships/vue-xlsx"

// ── Inline host component — creates a fresh controller per mount,
// enabling file-switching via reactive key on the parent. ─────────────
const XlsxSurfaceHost = defineComponent({
  props: {
    file: { type: Object as PropType<ArrayBuffer | null> },
    fileName: { type: String, default: "" },
    readOnly: { type: Boolean, default: false },
    zoom: { type: Number, default: 1 },
  },
  emits: ["cellDoubleClick", "contextMenu", "selectionChange", "objectClick", "update:zoom"],
  setup(props, { emit }) {
    const controller: XlsxViewerController = useXlsxViewerController({
      file: props.file ?? undefined,
      fileName: props.fileName,
      get readOnly() { return props.readOnly },
      onDiagnostic: () => {},
    })
    return () =>
      controller.activeSheet
        ? h(XlsxSheetSurface, {
            controller,
            "read-only": props.readOnly,
            zoom: props.zoom,
            style: { flex: "1" },
            onCellDoubleClick: (cell: any) => emit("cellDoubleClick", cell),
            onContextMenu: (ctx: any) => emit("contextMenu", ctx),
            onSelectionChange: (sel: any) => emit("selectionChange", sel),
            onObjectClick: (obj: any) => emit("objectClick", obj),
            "onUpdate:zoom": (value: number) => emit("update:zoom", value),
          })
        : null
  },
})

// ── Samples ──────────────────────────────────────────────────────────
const samples = [
  { file: "financial-model.xlsx", label: "财务模型" },
  { file: "sales-table.xlsx", label: "销售表" },
  { file: "charts-images.xlsx", label: "图表、图片、条件格式、迷你图、超链接、批注、形状与控件" },
  { file: "large-grid.xlsx", label: "大表格" },
]

// ── State ────────────────────────────────────────────────────────────
const readOnly = ref(false)
const selectedSample = ref(samples[0].file)
const fileBuffer = ref<ArrayBuffer | null>(null)
const displayName = ref("")
const loading = ref(false)
const error = ref<string | null>(null)
const loadCounter = ref(0)
const zoom = ref(1)

const surfaceKey = computed(() =>
  fileBuffer.value
    ? `${displayName.value}-${readOnly.value}-${loadCounter.value}`
    : "",
)

const selectionInfo = ref("—")
const contextMenuInfo = ref("—")

// ── Actions ──────────────────────────────────────────────────────────
const fileInputRef = ref<HTMLInputElement>()

async function loadSelectedSample(): Promise<void> {
  const fileName = selectedSample.value
  loading.value = true
  error.value = null
  try {
    const res = await fetch(`/samples/${encodeURIComponent(fileName)}`)
    if (!res.ok) { error.value = `Sample not found: ${fileName}`; return }
    fileBuffer.value = await res.arrayBuffer()
    displayName.value = fileName
    loadCounter.value++
  } catch (e) { error.value = `Load failed: ${String(e)}`
  } finally { loading.value = false }
}

async function onFileChange(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ""
  if (!file) return
  loading.value = true
  error.value = null
  try {
    fileBuffer.value = await file.arrayBuffer()
    displayName.value = file.name
    loadCounter.value++
  } catch (e) { error.value = `Open failed: ${String(e)}`
  } finally { loading.value = false }
}

function onCellDoubleClick() {
  // Host can listen for edit intent
}

function onContextMenu(ctx: { clientX: number; clientY: number; sheetName?: string; selection?: Record<string, unknown> }): void {
  const cell = ctx.selection ? `${String(ctx.selection)}` : ctx.sheetName ?? "sheet"
  contextMenuInfo.value = `${cell}`
}

function onSelectionChange(sel: { kind: string; range?: { start: { row: number; col: number }; end: { row: number; col: number } }; value?: string }): void {
  if (sel.kind === "none") selectionInfo.value = "—"
  else if (sel.range) selectionInfo.value = `${colLabel(sel.range.start.col)}${sel.range.start.row + 1}:${colLabel(sel.range.end.col)}${sel.range.end.row + 1}`
  else selectionInfo.value = sel.kind
}

function colLabel(col: number): string {
  let label = ""
  let n = col
  while (n >= 0) { label = String.fromCharCode(65 + (n % 26)) + label; n = Math.floor(n / 26) - 1 }
  return label
}

void loadSelectedSample()
</script>

<style scoped>
.page { padding: 16px; max-width: 1440px; margin: 0 auto; width: 100%; min-width: 0; }
h2 { margin-bottom: 4px; }
.desc { color: var(--muted-foreground); margin-bottom: 0; font-size: 13px; line-height: 1.5; }

.toolbar-zone {
  display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  padding: 10px 12px; margin: 12px 0;
  border: 1px solid var(--border); border-radius: var(--radius);
  background: var(--background);
}
.ctrl { display: flex; align-items: center; gap: 4px; font-size: 13px; white-space: nowrap; }
.ctrl select { padding: 4px 8px; border: 1px solid var(--border); border-radius: 4px; }
.ctrl input[type="file"] { font-size: 12px; }
.toolbar-zone button {
  padding: 4px 10px; border: 1px solid var(--border); border-radius: 4px;
  background: var(--background); cursor: pointer; font-size: 13px;
}
.toolbar-zone button:disabled { opacity: .4; cursor: default; }
.zoom-value { min-width: 48px; text-align: center; font-variant-numeric: tabular-nums; }
.sep { width: 1px; height: 20px; background: var(--border); }

.surface-container {
  display: flex; flex-direction: column;
  border: 1px solid var(--border); height: 68vh; position: relative;
}
.surface-overlay {
  align-items: center; display: flex; justify-content: center;
  position: absolute; inset: 0; z-index: 10;
  background: color-mix(in oklch, var(--background) 70%, transparent);
  font-size: 13px; color: var(--muted-foreground);
}
.empty { display: flex; align-items: center; justify-content: center; height: 100%; color: var(--muted-foreground); }
.error { color: #ef4444; font-size: 13px; margin-bottom: 8px; }

.status-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 8px;
  padding: 12px; margin-top: 8px; border: 1px solid var(--border); border-radius: var(--radius);
  background: var(--muted); font-size: 13px;
}
</style>
