<template>
  <div :class="cn('xlsx-viewer', props.className)" :style="{ height: '100%', minHeight: '300px', display: 'flex', flexDirection: 'column', overflow: 'hidden' }">
    <!-- Toolbar -->
    <div v-if="showToolbar" class="xlsx-viewer-toolbar" :style="toolbarStyle">
      <div style="display:flex;align-items:center;gap:4px;padding:0 8px;flex-wrap:wrap;">
        <span style="font-size:13px;font-weight:500;margin-right:8px;" :style="{ color: isDarkVal ? '#e5e7eb' : '#374151' }">
          {{ displayFileName }}
        </span>
        <span v-if="isLoadingVal" style="font-size:12px;color:#9ca3af;">Loading...</span>

        <div style="flex:1;" />

        <button :disabled="!canZoomOutVal" @click="zoomOut" :style="btnStyle" title="Zoom out">−</button>
        <span style="font-size:12px;min-width:42px;text-align:center;" :style="{ color: isDarkVal ? '#e5e7eb' : '#374151' }">
          {{ Math.round(zoomScaleVal * 100) }}%
        </span>
        <button :disabled="!canZoomInVal" @click="zoomIn" :style="btnStyle" title="Zoom in">+</button>
        <button @click="resetZoom" :style="btnStyle" title="Reset zoom">100%</button>

        <div style="width:1px;height:20px;background:#d1d5db;margin:0 4px;" />

        <button :disabled="!canUndoVal" @click="undo" :style="btnStyle" title="Undo">↩</button>
        <button :disabled="!canRedoVal" @click="redo" :style="btnStyle" title="Redo">↪</button>

        <div style="width:1px;height:20px;background:#d1d5db;margin:0 4px;" />

        <button v-if="canDownloadVal" @click="download" :style="btnStyle" title="Download">⬇</button>
      </div>
    </div>

    <!-- Sheet tabs -->
    <div v-if="showSheetTabs && tabsVal.length > 0" class="xlsx-viewer-sheet-tabs" :style="sheetTabsStyle">
      <button
        v-for="(tab, idx) in tabsVal"
        :key="tab.id"
        @click="setActiveTabIndex(idx)"
        :style="{
          ...sheetTabBaseStyle,
          background: idx === activeTabIndexVal ? (isDarkVal ? '#374151' : '#ffffff') : 'transparent',
          color: isDarkVal ? '#e5e7eb' : '#374151',
          fontWeight: idx === activeTabIndexVal ? 600 : 400,
          borderBottom: idx === activeTabIndexVal ? `2px solid ${isDarkVal ? '#60a5fa' : '#2563eb'}` : '2px solid transparent',
        }"
      >
        {{ tab.name }}
      </button>
    </div>

    <!-- Grid viewport -->
    <div
      ref="viewportRef"
      class="xlsx-viewer-grid"
      :style="{
        flex: 1,
        overflow: 'auto',
        position: 'relative',
        background: isDarkVal ? '#1f2937' : '#f9fafb',
      }"
    >
      <div v-if="loadErrorVal && !isLoadingVal" style="display:flex;align-items:center;justify-content:center;height:100%;color:#dc2626;font-size:14px;padding:16px;text-align:center;">
        <p>{{ loadErrorVal.message }}</p>
      </div>

      <div v-if="!activeSheetVal && !isLoadingVal && !loadErrorVal" style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;font-size:14px;">
        <p>No workbook loaded</p>
      </div>

      <div v-if="isLoadingVal" style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;font-size:14px;">
        <slot name="loading">
          <p>Loading workbook...</p>
        </slot>
      </div>

      <div v-if="activeSheetVal && !isLoadingVal && !loadErrorVal" style="position:relative;">
        <table
          class="xlsx-viewer-table"
          :style="{
            borderCollapse: 'collapse',
            tableLayout: 'fixed',
            fontSize: `${Math.round(11 * zoomScaleVal)}px`,
          }"
        >
          <colgroup>
            <col :style="{ width: `${ROW_HEADER_WIDTH}px`, minWidth: `${ROW_HEADER_WIDTH}px` }" />
            <col
              v-for="ci in visibleColRange"
              :key="ci"
              :style="{ width: `${colWidth(ci)}px`, minWidth: `${colWidth(ci)}px` }"
            />
          </colgroup>
          <thead>
            <tr>
              <th :style="cornerHeaderStyle"></th>
              <th
                v-for="ci in visibleColRange"
                :key="ci"
                :style="colHeaderStyle(ci)"
                @click="selectColumn(ci)"
              >
                {{ colLetter(ci) }}
              </th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="ri in visibleRowRange" :key="ri">
              <td :style="rowHeaderStyle(ri)" @click="selectRow(ri)">
                {{ ri + 1 }}
              </td>
              <td
                v-for="ci in visibleColRange"
                :key="`${ri}-${ci}`"
                :style="cellStyle(ri, ci)"
                @click="onCellClick(ri, ci)"
                @dblclick="onCellDblClick(ri, ci)"
              >
                <span v-if="editCell && editCell.row === ri && editCell.col === ci">
                  <input
                    v-model="editValue"
                    :style="editInputStyle"
                    @blur="commitEdit"
                    @keydown.enter="commitEdit"
                    @keydown.escape="cancelEdit"
                    @keydown.tab.prevent="commitEditAndMove"
                  />
                </span>
                <span v-else style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
                  {{ cellDisplayValue(ri, ci) }}
                </span>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, provide, inject, type ComputedRef } from "vue"
import type { XlsxViewerController, XlsxCellAddress } from "./composables"
import {
  useXlsxViewerController,
  XLSX_VIEWER_KEY,
  XLSX_VIEWER_DARK_KEY,
} from "./composables"

export interface XlsxViewerProps {
  file?: ArrayBuffer
  src?: string
  controller?: XlsxViewerController
  isDark?: boolean
  className?: string
  readOnly?: boolean
  showToolbar?: boolean
  showSheetTabs?: boolean
  fileName?: string
}

const props = withDefaults(defineProps<XlsxViewerProps>(), {
  showToolbar: true,
  showSheetTabs: true,
  isDark: false,
  readOnly: false,
})

// ── Resolve controller: prop > inject > internal ──
const ctx = inject<XlsxViewerController | ComputedRef<XlsxViewerController> | null>(XLSX_VIEWER_KEY, null)
const injectedCtrl: XlsxViewerController | undefined = ctx
  ? (typeof ctx === "object" && "value" in ctx ? (ctx as ComputedRef<XlsxViewerController>).value : ctx as XlsxViewerController)
  : undefined

const internalCtrl = useXlsxViewerController({
  file: props.file,
  src: props.src,
  readOnly: props.readOnly,
  fileName: props.fileName,
})

const ctrl = computed<XlsxViewerController>(() => {
  return props.controller ?? injectedCtrl ?? internalCtrl
})

provide(XLSX_VIEWER_KEY, ctrl.value)
provide(XLSX_VIEWER_DARK_KEY, { isDark: props.isDark })

// ── Reactive accessors for template ──
const isLoadingVal      = computed(() => ctrl.value.isLoading)
const loadErrorVal      = computed(() => ctrl.value.error)
const canZoomInVal      = computed(() => ctrl.value.canZoomIn)
const canZoomOutVal     = computed(() => ctrl.value.canZoomOut)
const zoomScaleVal      = computed(() => ctrl.value.zoomScale)
const localHistory = ref<Array<{ key: string; previous: string; next: string }>>([])
const localFuture = ref<Array<{ key: string; previous: string; next: string }>>([])
const canUndoVal        = computed(() => ctrl.value.canUndo || localHistory.value.length > 0)
const canRedoVal        = computed(() => ctrl.value.canRedo || localFuture.value.length > 0)
const canDownloadVal    = computed(() => ctrl.value.canDownload)
const displayFileName   = computed(() => ctrl.value.displayFileName)
const activeSheetVal    = computed(() => ctrl.value.activeSheet)
const activeSheetIndexVal = computed(() => ctrl.value.activeSheetIndex)
const activeTabIndexVal = computed(() => ctrl.value.activeTabIndex)
const tabsVal           = computed(() => ctrl.value.tabs)
const activeCellVal     = computed(() => ctrl.value.activeCell)
const selectionVal      = computed(() => ctrl.value.selection)
const isDarkVal         = computed(() => props.isDark)

function zoomIn()  { ctrl.value.zoomIn() }
function zoomOut() { ctrl.value.zoomOut() }
function resetZoom(){ ctrl.value.resetZoom() }
function undo() {
  const entry = localHistory.value.pop()
  if (!entry) { ctrl.value.undo(); return }
  if (entry.previous) cellData.value[entry.key] = entry.previous
  else delete cellData.value[entry.key]
  localFuture.value.push(entry)
}
function redo() {
  const entry = localFuture.value.pop()
  if (!entry) { ctrl.value.redo(); return }
  if (entry.next) cellData.value[entry.key] = entry.next
  else delete cellData.value[entry.key]
  localHistory.value.push(entry)
}
function download(){ ctrl.value.download() }
function setActiveTabIndex(i: number) { ctrl.value.setActiveTabIndex(i) }

// ── Grid constants ──
const ROW_HEADER_WIDTH = 40
const HEADER_HEIGHT = 24
const DEFAULT_COL_WIDTH = 80
const DEFAULT_ROW_HEIGHT = 24
const MAX_VISIBLE_COLS = 100
const MAX_VISIBLE_ROWS = 500

const visibleColRange = computed(() =>
  Array.from({ length: Math.min(activeSheetVal.value?.colCount ?? 0, MAX_VISIBLE_COLS) }, (_, i) => i)
)
const visibleRowRange = computed(() =>
  Array.from({ length: Math.min(activeSheetVal.value?.rowCount ?? 0, MAX_VISIBLE_ROWS) }, (_, i) => i)
)

function colWidth(ci: number): number {
  return Math.round((activeSheetVal.value?.colWidths?.[ci] ?? DEFAULT_COL_WIDTH) * zoomScaleVal.value)
}
function rowHeight(ri: number): number {
  return Math.round((activeSheetVal.value?.rowHeights?.[ri] ?? DEFAULT_ROW_HEIGHT) * zoomScaleVal.value)
}
function colLetter(ci: number): string {
  let label = ""
  let value = ci
  do {
    label = String.fromCharCode(65 + (value % 26)) + label
    value = Math.floor(value / 26) - 1
  } while (value >= 0)
  return label
}

const cellData = ref<Record<string, string>>({})

function cellKey(r: number, c: number): string { return `${r}:${c}` }
function sheetCellText(r: number, c: number): string {
  const sheet = activeSheetVal.value as unknown as { cellText?: Record<string, string> } | null
  return sheet?.cellText?.[cellKey(r, c)] ?? ""
}
function cellDisplayValue(r: number, c: number): string {
  return cellData.value[cellKey(r, c)] ?? sheetCellText(r, c)
}

// ── Editing ──
const editCell = ref<XlsxCellAddress | null>(null)
const editValue = ref("")

function onCellClick(ri: number, ci: number) {
  ctrl.value.selectCell({ row: ri, col: ci })
}
function onCellDblClick(ri: number, ci: number) {
  if (ctrl.value.readOnly) return
  editCell.value = { row: ri, col: ci }
  editValue.value = cellData.value[cellKey(ri, ci)] ?? sheetCellText(ri, ci)
  nextTick(() => {
    const el = document.querySelector<HTMLInputElement>(".xlsx-viewer-grid input")
    el?.focus()
    el?.select()
  })
}
function commitEdit() {
  if (!editCell.value) return
  const key = cellKey(editCell.value.row, editCell.value.col)
  const previous = cellDisplayValue(editCell.value.row, editCell.value.col)
  const next = editValue.value
  if (next) cellData.value[key] = next
  else delete cellData.value[key]
  if (previous !== next) {
    localHistory.value.push({ key, previous, next })
    localFuture.value = []
  }
  ctrl.value.setSelectedCellValue(next)
  editCell.value = null
}
function cancelEdit() { editCell.value = null }
function commitEditAndMove() {
  if (!editCell.value) return
  const { row, col } = editCell.value
  commitEdit()
  ctrl.value.selectCell({ row, col: col + 1 })
}
function selectColumn(ci: number) {
  const maxRow = activeSheetVal.value?.rowCount ?? 0
  ctrl.value.selectRange({ start: { col: ci, row: 0 }, end: { col: ci, row: maxRow - 1 } })
}
function selectRow(ri: number) {
  const maxCol = activeSheetVal.value?.colCount ?? 0
  ctrl.value.selectRange({ start: { col: 0, row: ri }, end: { col: maxCol - 1, row: ri } })
}

// ── Styles ──
const isDarkRef = computed(() => props.isDark)

const toolbarStyle = computed(() => ({
  height: '36px', display: 'flex', alignItems: 'center',
  borderBottom: `1px solid ${isDarkRef.value ? '#374151' : '#e5e7eb'}`,
  background: isDarkRef.value ? '#111827' : '#f3f4f6', flexShrink: 0,
}))
const btnStyle = computed(() => ({
  background: 'transparent', border: `1px solid ${isDarkRef.value ? '#4b5563' : '#d1d5db'}`,
  borderRadius: '4px', cursor: 'pointer', padding: '2px 6px', fontSize: '13px',
  color: isDarkRef.value ? '#e5e7eb' : '#374151', minWidth: '24px',
  display: 'flex', alignItems: 'center', justifyContent: 'center',
}))
const sheetTabsStyle = computed(() => ({
  height: '32px', display: 'flex', alignItems: 'stretch',
  borderBottom: `1px solid ${isDarkRef.value ? '#374151' : '#e5e7eb'}`,
  background: isDarkRef.value ? '#111827' : '#f3f4f6', flexShrink: 0, overflowX: 'auto' as const,
}))
const sheetTabBaseStyle = computed(() => ({
  padding: '0 12px', fontSize: '12px', cursor: 'pointer', border: 'none',
  background: 'transparent', whiteSpace: 'nowrap' as const, display: 'flex',
  alignItems: 'center', outline: 'none', transition: 'background 0.15s',
}))
const cornerHeaderStyle = computed(() => ({
  position: 'sticky' as const, top: 0, left: 0, zIndex: 3,
  background: isDarkRef.value ? '#374151' : '#e5e7eb',
  border: `1px solid ${isDarkRef.value ? '#4b5563' : '#d1d5db'}`,
  height: `${HEADER_HEIGHT}px`, width: `${ROW_HEADER_WIDTH}px`, minWidth: `${ROW_HEADER_WIDTH}px`,
}))
function colHeaderStyle(ci: number) {
  const sel = selectionVal.value
  const isSel = sel && ci >= sel.start.col && ci <= sel.end.col
  return {
    position: 'sticky' as const, top: 0, zIndex: 2,
    background: isSel ? (isDarkRef.value ? '#1e3a5f' : '#dbeafe') : (isDarkRef.value ? '#374151' : '#e5e7eb'),
    border: `1px solid ${isDarkRef.value ? '#4b5563' : '#d1d5db'}`,
    height: `${HEADER_HEIGHT}px`, cursor: 'pointer', textAlign: 'center' as const,
    fontSize: '11px', fontWeight: 400, color: isDarkRef.value ? '#e5e7eb' : '#374151',
    userSelect: 'none' as const, padding: 0,
  }
}
function rowHeaderStyle(ri: number) {
  const sel = selectionVal.value
  const isSel = sel && ri >= sel.start.row && ri <= sel.end.row
  return {
    position: 'sticky' as const, left: 0, zIndex: 1,
    background: isSel ? (isDarkRef.value ? '#1e3a5f' : '#dbeafe') : (isDarkRef.value ? '#374151' : '#e5e7eb'),
    border: `1px solid ${isDarkRef.value ? '#4b5563' : '#d1d5db'}`,
    height: `${rowHeight(ri)}px`, cursor: 'pointer', textAlign: 'center' as const,
    fontSize: '11px', fontWeight: 400, color: isDarkRef.value ? '#e5e7eb' : '#374151',
    userSelect: 'none' as const, padding: 0,
  }
}
function cellStyle(ri: number, ci: number) {
  const ac = activeCellVal.value
  const sel = selectionVal.value
  const isActive = ac?.row === ri && ac?.col === ci
  const isInSel = sel ? ri >= sel.start.row && ri <= sel.end.row && ci >= sel.start.col && ci <= sel.end.col : false
  return {
    border: `1px solid ${isDarkRef.value ? '#374151' : '#e5e7eb'}`,
    height: `${rowHeight(ri)}px`, padding: '0 4px', overflow: 'hidden', cursor: 'cell',
    fontSize: '12px', color: isDarkRef.value ? '#e5e7eb' : '#111827',
    background: isActive
      ? (isDarkRef.value ? '#1e3a5f' : '#dbeafe')
      : isInSel
        ? (isDarkRef.value ? 'rgba(59,130,246,0.15)' : 'rgba(59,130,246,0.08)')
        : (isDarkRef.value ? '#1f2937' : '#ffffff'),
    outline: isActive ? `2px solid ${isDarkRef.value ? '#60a5fa' : '#2563eb'}` : 'none',
    outlineOffset: '-2px',
  }
}
const editInputStyle = computed(() => ({
  width: '100%', height: '100%', border: 'none', outline: 'none', padding: '0 4px',
  fontSize: '12px', fontFamily: 'inherit', background: isDarkRef.value ? '#111827' : '#ffffff',
  color: isDarkRef.value ? '#e5e7eb' : '#111827', boxSizing: 'border-box' as const,
}))

function cn(...classes: (string | undefined | null | false)[]): string {
  return classes.filter(Boolean).join(" ")
}
</script>

<style scoped>
.xlsx-viewer-table {
  width: max-content;
  min-width: 100%;
}
.xlsx-viewer-table th,
.xlsx-viewer-table td {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
</style>
