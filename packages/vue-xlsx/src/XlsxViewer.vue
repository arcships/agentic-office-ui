<template>
  <div :class="cn('xlsx-viewer', props.className)" :style="{ height: '100%', minHeight: '300px', display: 'flex', flexDirection: 'column', overflow: 'hidden' }">
    <!-- Toolbar -->
    <div v-if="showToolbar" class="xlsx-viewer-toolbar" :class="{ dark: isDarkVal }">
      <div class="xlsx-title-cluster">
        <span class="xlsx-file-name">{{ displayFileName }}</span>
        <span v-if="ribbonStatus" class="xlsx-status">{{ ribbonStatus }}</span>
        <span v-if="isLoadingVal" class="xlsx-loading">Loading...</span>
        <input ref="fileInputRef" class="xlsx-hidden-file" type="file" accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" @change="loadOpenedFile" />
      </div>
      <div class="xlsx-ribbon-tabs" role="tablist" aria-label="Workbook ribbon tabs">
        <button
          v-for="tab in ribbonTabs"
          :key="tab"
          class="ribbon-tab"
          :class="{ active: activeRibbonTab === tab }"
          type="button"
          role="tab"
          :aria-selected="activeRibbonTab === tab"
          @click="activeRibbonTab = tab"
        >
          {{ tab }}
        </button>
      </div>
      <div class="xlsx-toolbar-groups" role="toolbar" :aria-label="`${activeRibbonTab} ribbon toolbar`">
        <template v-if="activeRibbonTab === 'Home'">
          <div class="ribbon-group">
            <button class="ribbon-button" :disabled="!canUndoVal" @click="undo" title="Undo" aria-label="Undo">↶</button>
            <button class="ribbon-button" :disabled="!canRedoVal" @click="redo" title="Redo" aria-label="Redo">↷</button>
            <span class="ribbon-label">Clipboard</span>
          </div>
          <div class="ribbon-group ribbon-stack">
            <div class="ribbon-row">
              <select class="ribbon-select font-family" :disabled="!canEditSelection" v-model="fontFamily" @change="applyFontStyle({ name: fontFamily })" aria-label="Font family">
                <option v-for="family in fontFamilies" :key="family" :value="family">{{ family }}</option>
              </select>
              <select class="ribbon-select font-size" :disabled="!canEditSelection" v-model="fontSize" @change="applyFontStyle({ size: Number(fontSize) })" aria-label="Font size">
                <option v-for="size in fontSizes" :key="size" :value="String(size)">{{ size }}</option>
              </select>
            </div>
            <div class="ribbon-row">
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyFontStyle({ bold: true })" aria-label="Bold"><b>B</b></button>
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyFontStyle({ italic: true })" aria-label="Italic"><i>I</i></button>
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyFontStyle({ underline: 'single' })" aria-label="Underline"><u>U</u></button>
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyFontStyle({ strikethrough: true })" aria-label="Strikethrough"><s>S</s></button>
              <input class="ribbon-color" :disabled="!canEditSelection" type="color" v-model="fontColor" @input="applyFontStyle({ color: hexToStyleColor(fontColor) })" aria-label="Text color" />
              <input class="ribbon-color" :disabled="!canEditSelection" type="color" v-model="fillColor" @input="applyStyle({ fill: { fillType: 'solid', color: hexToStyleColor(fillColor) } })" aria-label="Fill color" />
            </div>
            <span class="ribbon-label">Font</span>
          </div>
          <div class="ribbon-group ribbon-stack">
            <div class="ribbon-row">
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyAlignmentStyle({ horizontal: 'left' })" aria-label="Align left">L</button>
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyAlignmentStyle({ horizontal: 'center' })" aria-label="Center">C</button>
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyAlignmentStyle({ horizontal: 'right' })" aria-label="Align right">R</button>
            </div>
            <div class="ribbon-row">
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyAlignmentStyle({ vertical: 'top' })" aria-label="Top align">T</button>
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyAlignmentStyle({ vertical: 'center' })" aria-label="Middle align">M</button>
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyAlignmentStyle({ vertical: 'bottom' })" aria-label="Bottom align">B</button>
              <button class="ribbon-button wide" :disabled="!canEditSelection" @click="applyAlignmentStyle({ wrapText: true })" aria-label="Wrap text">Wrap</button>
            </div>
            <span class="ribbon-label">Alignment</span>
          </div>
          <div class="ribbon-group ribbon-stack">
            <select class="ribbon-select number-format" :disabled="!canEditSelection" v-model="numberFormat" @change="applyNumberFormat(numberFormat)" aria-label="Number format">
              <option v-for="format in numberFormats" :key="format" :value="format">{{ format }}</option>
            </select>
            <div class="ribbon-row">
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyNumberFormat('Currency')">$</button>
              <button class="ribbon-button" :disabled="!canEditSelection" @click="applyNumberFormat('Percent')">%</button>
              <button class="ribbon-button wide" :disabled="!canEditSelection" @click="applyNumberFormat('Thousands')">000</button>
            </div>
            <span class="ribbon-label">Number</span>
          </div>
          <div class="ribbon-group ribbon-stack styles-group">
            <div class="style-presets">
              <button v-for="preset in stylePresets" :key="preset.label" class="style-preset" :disabled="!canEditSelection" @click="applyStyle(preset.style)">
                <span :style="stylePresetSwatch(preset.style)"></span>{{ preset.label }}
              </button>
            </div>
            <span class="ribbon-label">Styles</span>
          </div>
          <div class="ribbon-group ribbon-stack">
            <div class="ribbon-row">
              <button class="ribbon-button wide" :disabled="!hasSelection || isReadOnlyVal" @click="mergeSelection">Merge</button>
              <button class="ribbon-button wide" :disabled="!hasSelection || isReadOnlyVal" @click="unmergeSelection">Unmerge</button>
            </div>
            <div class="ribbon-row">
              <input class="ribbon-color" :disabled="!canEditSelection" type="color" v-model="borderColor" aria-label="Border color" />
              <button class="ribbon-button wide" :disabled="!canEditSelection" @click="applyBorder('bottom')">Bottom</button>
              <button class="ribbon-button wide" :disabled="!canEditSelection" @click="applyBorder('all')">All</button>
              <button class="ribbon-button wide" :disabled="!canEditSelection" @click="applyBorder('none')">None</button>
            </div>
            <span class="ribbon-label">Cells</span>
          </div>
        </template>
        <template v-else-if="activeRibbonTab === 'Insert'">
          <div class="ribbon-group">
            <button class="ribbon-button wide" :disabled="isReadOnlyVal" @click="addSheet">+ Sheet</button>
            <button class="ribbon-button wide" :disabled="tabsVal.length <= 1 || isReadOnlyVal" @click="removeActiveSheet">Delete</button>
            <span class="ribbon-label">Workbook</span>
          </div>
          <div class="ribbon-group">
            <button class="ribbon-button wide" @click="triggerFileOpen">Open</button>
            <span class="ribbon-label">Open</span>
          </div>
          <div class="ribbon-group ribbon-stack source-group">
            <input class="ribbon-url-input" v-model="remoteUrl" placeholder="https://example.com/report.xlsx" @keydown.enter.prevent="loadRemoteUrl" aria-label="Workbook URL" />
            <button class="ribbon-button wide" :disabled="!remoteUrl.trim()" @click="loadRemoteUrl">Load</button>
            <span class="ribbon-label">Source</span>
          </div>
        </template>
        <template v-else-if="activeRibbonTab === 'Page Layout'">
          <div class="ribbon-group">
            <label class="ribbon-switch"><input type="checkbox" v-model="documentDark" /> Document dark</label>
            <span class="ribbon-label">Theme</span>
          </div>
          <div class="ribbon-group">
            <button class="ribbon-button wide" :disabled="!canDownloadVal" @click="download">Source</button>
            <button class="ribbon-button wide" :disabled="!canDownloadVal" @click="exportXlsx">XLSX</button>
            <button class="ribbon-button wide" :disabled="!canDownloadVal" @click="exportCsv">CSV</button>
            <span class="ribbon-label">Export</span>
          </div>
        </template>
        <template v-else-if="activeRibbonTab === 'Formulas'">
          <div class="ribbon-group">
            <button class="ribbon-button wide" :disabled="!canDownloadVal" @click="recalculate">Recalc</button>
            <span class="ribbon-label">Calculation</span>
          </div>
          <div class="ribbon-group ribbon-stack source-group">
            <input class="ribbon-url-input" v-model="namedRangeDraft" placeholder="Named range" @keydown.enter.prevent="defineNamedRange" aria-label="Named range" />
            <button class="ribbon-button wide" :disabled="!hasSelection || !namedRangeDraft.trim() || isReadOnlyVal" @click="defineNamedRange">Define</button>
            <span class="ribbon-label">Defined Names</span>
          </div>
        </template>
        <template v-else-if="activeRibbonTab === 'Data'">
          <div class="ribbon-group ribbon-stack">
            <div class="ribbon-row">
              <button class="ribbon-button wide" :disabled="!activeSheetVal" @click="sortActiveColumn('ascending')">Sort A→Z</button>
              <button class="ribbon-button wide" :disabled="!activeSheetVal" @click="sortActiveColumn('descending')">Sort Z→A</button>
            </div>
            <span class="ribbon-label">Tables</span>
          </div>
          <div class="ribbon-group">
            <button class="ribbon-button wide" :disabled="!canDownloadVal" @click="recalculate">Recalc</button>
            <span class="ribbon-label">Refresh</span>
          </div>
          <div class="ribbon-group">
            <button class="ribbon-button wide" :disabled="!activeSheetVal" @click="clearWorkbook">Clear</button>
            <span class="ribbon-label">Workbook</span>
          </div>
        </template>
        <template v-else>
          <div class="ribbon-group">
            <button class="ribbon-button" :disabled="!activeSheetVal || !canZoomOutVal" @click="zoomOut" aria-label="Zoom out">−</button>
            <select class="ribbon-select zoom-select" :disabled="!activeSheetVal" :value="Math.round(zoomScaleVal * 100)" @change="setZoomFromSelect">
              <option v-for="choice in zoomChoices" :key="choice" :value="choice">{{ choice }}%</option>
            </select>
            <button class="ribbon-button" :disabled="!activeSheetVal || !canZoomInVal" @click="zoomIn" aria-label="Zoom in">+</button>
            <button class="ribbon-button wide" :disabled="!activeSheetVal || Math.round(zoomScaleVal * 100) === 100" @click="resetZoom">Reset</button>
            <span class="ribbon-label">Zoom</span>
          </div>
          <div class="ribbon-group">
            <button class="ribbon-button" :disabled="!activeSheetVal || activeTabIndexVal <= 0" @click="setActiveTabIndex(activeTabIndexVal - 1)" aria-label="Previous sheet">‹</button>
            <button class="ribbon-button" :disabled="!activeSheetVal || activeTabIndexVal >= tabsVal.length - 1" @click="setActiveTabIndex(activeTabIndexVal + 1)" aria-label="Next sheet">›</button>
            <select class="ribbon-select sheet-select" :disabled="tabsVal.length === 0" :value="String(activeTabIndexVal)" @change="setActiveTabFromSelect">
              <option v-for="(tab, idx) in tabsVal" :key="tab.id" :value="String(idx)">{{ tab.name }}</option>
            </select>
            <span class="ribbon-label">Sheets</span>
          </div>
          <div class="ribbon-group">
            <label class="ribbon-switch"><input type="checkbox" v-model="highlightCells" /> Highlight</label>
            <label class="ribbon-switch"><input type="checkbox" :checked="isReadOnlyVal" @change="toggleReadOnly" /> Read only</label>
            <span class="ribbon-label">Display</span>
          </div>
        </template>
      </div>
      <div class="xlsx-formula-bar">
        <input class="name-box" :value="activeCellAddressText" readonly aria-label="Name box" />
        <span class="fx-label">fx</span>
        <input
          class="formula-input"
          v-model="formulaDraft"
          :disabled="!activeCellVal || isReadOnlyVal"
          aria-label="Formula bar"
          @focus="formulaFocusValue = formulaDraft"
          @keydown.enter.prevent="commitFormula"
          @keydown.escape.prevent="cancelFormula"
          @blur="commitFormula"
        />
      </div>
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
                <span>{{ colLetter(ci) }}</span>
                <span class="xlsx-resize-handle col-resize" @mousedown.stop.prevent="startColumnResize(ci, $event)"></span>
              </th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="ri in visibleRowRange" :key="ri">
              <td :style="rowHeaderStyle(ri)" @click="selectRow(ri)">
                <span>{{ ri + 1 }}</span>
                <span class="xlsx-resize-handle row-resize" @mousedown.stop.prevent="startRowResize(ri, $event)"></span>
              </td>
              <td
                v-for="ci in visibleColRange"
                v-show="!isMergedHidden(ri, ci)"
                :key="`${ri}-${ci}`"
                :rowspan="mergeSpan(ri, ci).rowspan"
                :colspan="mergeSpan(ri, ci).colspan"
                :style="cellStyle(ri, ci)"
                @mousedown.prevent="onCellMouseDown(ri, ci, $event)"
                @mouseenter="onCellMouseEnter(ri, ci)"
                @click="onCellClick(ri, ci, $event)"
                @dblclick.stop="onCellDblClick(ri, ci)"
              >
                <span v-if="editCell && editCell.row === ri && editCell.col === ci">
                  <input
                    v-model="editValue"
                    :style="editInputStyle"
                    @blur="commitEdit"
                    @keydown.enter.prevent="commitEdit"
                    @keydown.escape="cancelEdit"
                    @keydown.tab.prevent="commitEditAndMove"
                  />
                </span>
                <span v-if="dataBarPercent(ri, ci) != null" class="xlsx-data-bar" :style="dataBarStyle(ri, ci)"></span>
                <span v-if="validationForCell(ri, ci)" class="xlsx-validation-indicator" :title="validationTitle(ri, ci)">▾</span>
                <span v-if="sparklineForCell(ri, ci)" class="xlsx-sparkline" :title="sparklineTitle(ri, ci)">{{ sparklineGlyph(ri, ci) }}</span>
                <span v-else-if="cellHasFormula(ri, ci)" class="xlsx-formula-badge" title="Formula cell">fx</span>
                <span v-if="!sparklineForCell(ri, ci)" style="position:relative;z-index:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;display:block;">
                  {{ cellDisplayValue(ri, ci) }}
                </span>
              </td>
            </tr>
          </tbody>
        </table>
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
          borderTop: idx === activeTabIndexVal ? `2px solid ${isDarkVal ? '#60a5fa' : '#2563eb'}` : '2px solid transparent',
        }"
      >
        <canvas
          :ref="(el) => setThumbnailCanvas(el as Element | null, idx)"
          class="xlsx-sheet-thumbnail"
          width="144"
          height="88"
          aria-hidden="true"
        ></canvas>
        {{ tab.name }}
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, provide, inject, watch, onMounted, type ComputedRef } from "vue"
import type { XlsxViewerController, XlsxCellAddress } from "./composables"
import {
  useXlsxViewerController,
  XLSX_VIEWER_KEY,
  XLSX_VIEWER_DARK_KEY,
  useXlsxViewerThumbnails,
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
const thumbnailsApi = useXlsxViewerThumbnails()

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
const isDarkVal         = computed(() => props.isDark || documentDark.value)
const isReadOnlyVal     = computed(() => ctrl.value.readOnly || props.readOnly || localReadOnly.value)
const hasSelection      = computed(() => !!selectionVal.value)
const canEditSelection  = computed(() => !!activeCellVal.value && !isReadOnlyVal.value)

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
function setThumbnailCanvas(el: Element | null, index: number) {
  thumbnailCanvases.value[index] = el instanceof HTMLCanvasElement ? el : null
  if (el instanceof HTMLCanvasElement) thumbnailsApi.paintThumbnail(index, el)
}
watch([tabsVal, activeSheetVal, () => ctrl.value.revision], () => {
  void nextTick(() => thumbnailCanvases.value.forEach((canvas, index) => { if (canvas) thumbnailsApi.paintThumbnail(index, canvas) }))
}, { deep: true })
onMounted(() => {
  void nextTick(() => thumbnailCanvases.value.forEach((canvas, index) => { if (canvas) thumbnailsApi.paintThumbnail(index, canvas) }))
})

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

function isFrozenCol(ci: number): boolean { return ci < (activeSheetVal.value?.freezePanes?.col ?? 0) }
function isFrozenRow(ri: number): boolean { return ri < (activeSheetVal.value?.freezePanes?.row ?? 0) }
function frozenColLeft(ci: number): number {
  let left = ROW_HEADER_WIDTH
  for (let col = 0; col < ci; col++) left += colWidth(col)
  return left
}
function frozenRowTop(ri: number): number {
  let top = HEADER_HEIGHT
  for (let row = 0; row < ri; row++) top += rowHeight(row)
  return top
}

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
const isDraggingSelection = ref(false)
const dragAnchor = ref<XlsxCellAddress | null>(null)

function normalizedRange(range: { start: XlsxCellAddress; end: XlsxCellAddress } | null) {
  if (!range) return null
  return {
    start: { row: Math.min(range.start.row, range.end.row), col: Math.min(range.start.col, range.end.col) },
    end: { row: Math.max(range.start.row, range.end.row), col: Math.max(range.start.col, range.end.col) },
  }
}

function finishDragSelection() {
  isDraggingSelection.value = false
  dragAnchor.value = null
  window.removeEventListener("mouseup", finishDragSelection)
}

function cellKey(r: number, c: number): string { return `${r}:${c}` }
function sheetCellText(r: number, c: number): string {
  const sheet = activeSheetVal.value as unknown as { cellText?: Record<string, string> } | null
  return sheet?.cellText?.[cellKey(r, c)] ?? ""
}
function cellHasFormula(r: number, c: number): boolean { return sheetCellText(r, c).startsWith("=") }
function cellDisplayValue(r: number, c: number): string {
  const key = cellKey(r, c)
  const local = cellData.value[key]
  if (local != null) return local.startsWith("=") ? cachedFormulaValue(r, c) ?? local : local
  const raw = sheetCellText(r, c)
  return raw.startsWith("=") ? cachedFormulaValue(r, c) ?? raw : raw
}
function cachedFormulaValue(r: number, c: number): string | null {
  const sheet = activeSheetVal.value as unknown as { cachedFormulaValues?: Record<string, string> } | null
  return sheet?.cachedFormulaValues?.[cellKey(r, c)] ?? null
}
function rangeContains(range: { start: XlsxCellAddress; end: XlsxCellAddress }, row: number, col: number): boolean {
  const n = normalizedRange(range)
  return !!n && row >= n.start.row && row <= n.end.row && col >= n.start.col && col <= n.end.col
}
function numericCellValue(row: number, col: number): number | null {
  const value = cellDisplayValue(row, col).replace(/,/g, "")
  const n = Number(value)
  return Number.isFinite(n) ? n : null
}
function validationForCell(row: number, col: number): Record<string, unknown> | null {
  const sheet = activeSheetVal.value as unknown as { dataValidations?: Array<{ ranges?: Array<{ start: XlsxCellAddress; end: XlsxCellAddress }>; validationType?: string; inputMessage?: string; errorMessage?: string; listSource?: string }> } | null
  return sheet?.dataValidations?.find((rule) => rule.ranges?.some((range) => rangeContains(range, row, col))) as Record<string, unknown> | undefined ?? null
}
function validationTitle(row: number, col: number): string {
  const rule = validationForCell(row, col)
  return rule ? `Validation: ${String(rule.validationType ?? "custom")}${rule.listSource ? ` (${String(rule.listSource)})` : ""}${rule.inputMessage ? ` — ${String(rule.inputMessage)}` : ""}` : ""
}
function sparklineForCell(row: number, col: number): Record<string, unknown> | null {
  const sheet = activeSheetVal.value as unknown as { sparklines?: Array<{ target?: XlsxCellAddress; range?: { start: XlsxCellAddress; end: XlsxCellAddress }; type?: string }> } | null
  return sheet?.sparklines?.find((sparkline) => sparkline.target?.row === row && sparkline.target?.col === col) as Record<string, unknown> | undefined ?? null
}
function sparklineTitle(row: number, col: number): string {
  const spark = sparklineForCell(row, col) as { type?: string; range?: { start: XlsxCellAddress; end: XlsxCellAddress } } | null
  if (!spark?.range) return "Sparkline"
  return `${spark.type ?? "line"} sparkline ${colLetter(spark.range.start.col)}${spark.range.start.row + 1}:${colLetter(spark.range.end.col)}${spark.range.end.row + 1}`
}
function sparklineGlyph(row: number, col: number): string {
  const spark = sparklineForCell(row, col) as { type?: string } | null
  return spark?.type === "column" ? "▁▃▇▅" : spark?.type === "winLoss" ? "+−+" : "⌁⌁⌁"
}
function dataBarRule(row: number, col: number): { ranges?: Array<{ start: XlsxCellAddress; end: XlsxCellAddress }>; color?: { rgb?: string } } | null {
  const sheet = activeSheetVal.value as unknown as { conditionalFormatRules?: Array<{ kind?: string; ranges?: Array<{ start: XlsxCellAddress; end: XlsxCellAddress }>; color?: { rgb?: string } }> } | null
  return sheet?.conditionalFormatRules?.find((rule) => rule.kind === "dataBar" && rule.ranges?.some((range) => rangeContains(range, row, col))) ?? null
}
function dataBarPercent(row: number, col: number): number | null {
  const rule = dataBarRule(row, col)
  const value = numericCellValue(row, col)
  if (!rule || value == null) return null
  const values: number[] = []
  for (const range of rule.ranges ?? []) {
    const n = normalizedRange(range)
    if (!n) continue
    for (let r = n.start.row; r <= n.end.row; r++) for (let c = n.start.col; c <= n.end.col; c++) {
      const v = numericCellValue(r, c)
      if (v != null) values.push(v)
    }
  }
  const min = values.length ? Math.min(...values) : 0
  const max = values.length ? Math.max(...values) : 1
  return max === min ? 100 : Math.max(0, Math.min(100, ((value - min) / (max - min)) * 100))
}
function dataBarStyle(row: number, col: number) {
  const rule = dataBarRule(row, col)
  const rgb = rule?.color?.rgb?.replace(/^FF/i, "") ?? "63C384"
  return { width: `${dataBarPercent(row, col) ?? 0}%`, backgroundColor: `#${rgb}` }
}

// ── Editing ──
const editCell = ref<XlsxCellAddress | null>(null)
const editValue = ref("")
const ribbonTabs = ["Home", "Insert", "Page Layout", "Formulas", "Data", "View"] as const
const activeRibbonTab = ref<(typeof ribbonTabs)[number]>("Home")
const fontFamilies = ["Aptos", "Calibri", "Arial", "Georgia", "Times New Roman", "Courier New"] as const
const fontSizes = [9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36] as const
const numberFormats = ["General", "Number", "Currency", "Percent", "Date", "Time", "Thousands"] as const
const fontFamily = ref<(typeof fontFamilies)[number]>("Aptos")
const fontSize = ref("11")
const fontColor = ref("#1f2937")
const fillColor = ref("#dbeafe")
const borderColor = ref("#4b5563")
const numberFormat = ref<(typeof numberFormats)[number]>("General")
const remoteUrl = ref("")
const namedRangeDraft = ref("")
const documentDark = ref(false)
const highlightCells = ref(false)
const localReadOnly = ref(false)
const definedNames = ref<Array<{ name: string; range: string }>>([])
const ribbonStatus = ref("")
const fileInputRef = ref<HTMLInputElement | null>(null)
const formulaDraft = ref("")
const thumbnailCanvases = ref<Array<HTMLCanvasElement | null>>([])
const formulaFocusValue = ref("")
const activeCellAddressText = computed(() => activeCellVal.value ? `${colLetter(activeCellVal.value.col)}${activeCellVal.value.row + 1}` : "")
const zoomChoices = computed(() => {
  const choices = [50, 75, 100, 125, 150, 200]
  const current = Math.round(zoomScaleVal.value * 100)
  return choices.includes(current) ? choices : [...choices, current].sort((a, b) => a - b)
})

type CellStylePatch = {
  font?: Record<string, unknown>
  alignment?: Record<string, unknown>
  border?: Record<string, unknown>
  fill?: Record<string, unknown>
  numberFormat?: Record<string, unknown>
}

const cellStyles = ref<Record<string, CellStylePatch>>({})
const mergedRanges = ref<Array<{ start: XlsxCellAddress; end: XlsxCellAddress }>>([])
const stylePresets = [
  { label: "Good", style: { fill: { fillType: "solid", color: hexToStyleColor("#E2F0D9") }, font: { color: hexToStyleColor("#375623") } } },
  { label: "Neutral", style: { fill: { fillType: "solid", color: hexToStyleColor("#FFF2CC") }, font: { color: hexToStyleColor("#7F6000") } } },
  { label: "Bad", style: { fill: { fillType: "solid", color: hexToStyleColor("#FCE4D6") }, font: { color: hexToStyleColor("#9C0006") } } },
  { label: "Heading", style: { alignment: { horizontal: "center", vertical: "center" }, border: { bottom: { style: "medium", color: hexToStyleColor("#5B9BD5") } }, font: { bold: true, color: hexToStyleColor("#1F4E79"), size: 14 } } },
] satisfies Array<{ label: string; style: CellStylePatch }>

function hexToStyleColor(hex: string) {
  return { colorType: "rgb", hex: hex.replace("#", "").toUpperCase() }
}
function selectedCells(): XlsxCellAddress[] {
  const range = normalizedRange(selectionVal.value)
  if (range) {
    const out: XlsxCellAddress[] = []
    for (let row = range.start.row; row <= range.end.row; row++) {
      for (let col = range.start.col; col <= range.end.col; col++) out.push({ row, col })
    }
    return out
  }
  return activeCellVal.value ? [activeCellVal.value] : []
}
function mergeStylePatch(base: CellStylePatch, patch: CellStylePatch): CellStylePatch {
  return {
    ...base,
    ...patch,
    font: patch.font ? { ...(base.font ?? {}), ...patch.font } : base.font,
    alignment: patch.alignment ? { ...(base.alignment ?? {}), ...patch.alignment } : base.alignment,
    border: patch.border ? { ...(base.border ?? {}), ...patch.border } : base.border,
    fill: patch.fill ? { ...(base.fill ?? {}), ...patch.fill } : base.fill,
    numberFormat: patch.numberFormat ? { ...(base.numberFormat ?? {}), ...patch.numberFormat } : base.numberFormat,
  }
}
function applyStyle(style: CellStylePatch) {
  if (!canEditSelection.value) return
  const cells = selectedCells()
  for (const cell of cells) {
    const key = cellKey(cell.row, cell.col)
    cellStyles.value[key] = mergeStylePatch(cellStyles.value[key] ?? {}, style)
  }
  if (selectionVal.value) ctrl.value.setRangeStyle(selectionVal.value, style as never)
  else ctrl.value.setSelectedCellStyle(style as never)
  ribbonStatus.value = `Styled ${cells.length} cell${cells.length === 1 ? "" : "s"}`
}
function applyFontStyle(font: Record<string, unknown>) { applyStyle({ font }) }
function applyAlignmentStyle(alignment: Record<string, unknown>) { applyStyle({ alignment }) }
function applyNumberFormat(format: string) {
  numberFormat.value = format as (typeof numberFormats)[number]
  const formatString = format === "Currency" ? "$#,##0.00" : format === "Percent" ? "0.00%" : format === "Number" ? "#,##0.00" : format === "Date" ? "m/d/yyyy" : format === "Time" ? "h:mm AM/PM" : format === "Thousands" ? "#,##0" : "General"
  applyStyle({ numberFormat: { formatCode: formatString } })
}
function sortActiveColumn(direction: "ascending" | "descending") {
  const col = activeCellVal.value?.col ?? selectionVal.value?.start.col ?? 0
  ctrl.value.sortTable(activeSheetVal.value?.name ?? "active", col, direction)
  ribbonStatus.value = `Sorted column ${colLetter(col)} ${direction}`
}
function startColumnResize(col: number, event: MouseEvent) {
  const startX = event.clientX
  const startWidth = colWidth(col)
  const move = (moveEvent: MouseEvent) => ctrl.value.resizeColumn(col, Math.max(24, startWidth + moveEvent.clientX - startX))
  const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up) }
  window.addEventListener("mousemove", move)
  window.addEventListener("mouseup", up)
}
function startRowResize(row: number, event: MouseEvent) {
  const startY = event.clientY
  const startHeight = rowHeight(row)
  const move = (moveEvent: MouseEvent) => ctrl.value.resizeRow(row, Math.max(18, startHeight + moveEvent.clientY - startY))
  const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up) }
  window.addEventListener("mousemove", move)
  window.addEventListener("mouseup", up)
}
function applyBorder(mode: "bottom" | "all" | "none") {
  const none = { style: "none" }
  const line = { style: "thin", color: hexToStyleColor(borderColor.value) }
  const border = mode === "bottom" ? { bottom: line } : mode === "all" ? { top: line, right: line, bottom: line, left: line } : { top: none, right: none, bottom: none, left: none }
  applyStyle({ border })
}
function stylePresetSwatch(style: CellStylePatch) {
  const fill = style.fill as { color?: { hex?: string } } | undefined
  const font = style.font as { color?: { hex?: string } } | undefined
  return { backgroundColor: fill?.color?.hex ? `#${fill.color.hex}` : "transparent", borderColor: font?.color?.hex ? `#${font.color.hex}` : "currentColor" }
}
function mergeSelection() {
  if (!selectionVal.value || isReadOnlyVal.value) return
  mergedRanges.value.push(selectionVal.value)
  ctrl.value.mergeSelection()
  ribbonStatus.value = `Merged ${ctrl.value.selectedRangeAddress ?? "selection"}`
}
function unmergeSelection() {
  if (!selectionVal.value || isReadOnlyVal.value) return
  const range = normalizedRange(selectionVal.value)
  mergedRanges.value = mergedRanges.value.filter((item) => {
    const normalized = normalizedRange(item)
    return !range || !normalized || normalized.start.row !== range.start.row || normalized.start.col !== range.start.col || normalized.end.row !== range.end.row || normalized.end.col !== range.end.col
  })
  ctrl.value.unmergeSelection()
  ribbonStatus.value = `Unmerged ${ctrl.value.selectedRangeAddress ?? "selection"}`
}
function addSheet() {
  if (isReadOnlyVal.value) return
  ctrl.value.addSheet()
  ctrl.value.setActiveTabIndex(ctrl.value.tabs.length - 1)
  ribbonStatus.value = "Sheet added"
}
function removeActiveSheet() {
  if (tabsVal.value.length <= 1 || isReadOnlyVal.value) return
  ctrl.value.removeActiveSheet()
  ribbonStatus.value = "Active sheet removed"
}
function triggerFileOpen() { fileInputRef.value?.click() }
async function loadOpenedFile(event: Event) {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return
  try {
    const buffer = await file.arrayBuffer()
    await ctrl.value.loadWorkbookFromBuffer(buffer, file.name)
    resetLocalWorkbookUiState()
    ribbonStatus.value = `Opened ${file.name}`
  } catch (e) {
    ribbonStatus.value = e instanceof Error ? e.message : "Failed to open workbook"
  } finally {
    input.value = ""
  }
}
async function loadRemoteUrl() {
  const url = remoteUrl.value.trim()
  if (!url) return
  try {
    await ctrl.value.loadWorkbookFromUrl(url)
    resetLocalWorkbookUiState()
    ribbonStatus.value = `Loaded ${ctrl.value.displayFileName}`
  } catch (e) {
    ribbonStatus.value = e instanceof Error ? e.message : "Failed to load remote workbook"
  }
}
function exportXlsx() { ctrl.value.exportXlsx() }
function exportCsv() { ctrl.value.exportCsv() }
function recalculate() { ctrl.value.recalculate(); ribbonStatus.value = "Workbook recalculated" }
function defineNamedRange() {
  const name = namedRangeDraft.value.trim()
  if (!name || !selectionVal.value || isReadOnlyVal.value) return
  ctrl.value.defineNamedRange(name, selectionVal.value)
  definedNames.value.push({ name, range: ctrl.value.selectedRangeAddress ?? "selection" })
  namedRangeDraft.value = ""
  ribbonStatus.value = `Defined ${name}`
}
function resetLocalWorkbookUiState() {
  cellData.value = {}
  cellStyles.value = {}
  mergedRanges.value = []
  localHistory.value = []
  localFuture.value = []
  formulaDraft.value = ""
}
function clearWorkbook() {
  ctrl.value.clearSelection()
  resetLocalWorkbookUiState()
  ribbonStatus.value = "Cleared local edits and selection"
}
function setZoomFromSelect(event: Event) { ctrl.value.setZoomScale(Number((event.target as HTMLSelectElement).value) / 100) }
function setActiveTabFromSelect(event: Event) { setActiveTabIndex(Number((event.target as HTMLSelectElement).value)) }
function toggleReadOnly(event: Event) {
  const target = event.target as HTMLInputElement
  localReadOnly.value = target.checked
}
function isMergedStart(ri: number, ci: number) {
  const sheet = activeSheetVal.value as unknown as { mergedRanges?: Array<{ start: XlsxCellAddress; end: XlsxCellAddress }> } | null
  return [...(sheet?.mergedRanges ?? []), ...mergedRanges.value].some((range) => normalizedRange(range)?.start.row === ri && normalizedRange(range)?.start.col === ci)
}
function isMergedHidden(ri: number, ci: number) {
  const sheet = activeSheetVal.value as unknown as { mergedRanges?: Array<{ start: XlsxCellAddress; end: XlsxCellAddress }> } | null
  return [...(sheet?.mergedRanges ?? []), ...mergedRanges.value].some((range) => {
    const r = normalizedRange(range)
    return !!r && ri >= r.start.row && ri <= r.end.row && ci >= r.start.col && ci <= r.end.col && !(ri === r.start.row && ci === r.start.col)
  })
}
function mergeSpan(ri: number, ci: number) {
  const sheet = activeSheetVal.value as unknown as { mergedRanges?: Array<{ start: XlsxCellAddress; end: XlsxCellAddress }> } | null
  const range = [...(sheet?.mergedRanges ?? []), ...mergedRanges.value].map(normalizedRange).find((r) => r && r.start.row === ri && r.start.col === ci)
  return range ? { rowspan: range.end.row - range.start.row + 1, colspan: range.end.col - range.start.col + 1 } : { rowspan: 1, colspan: 1 }
}
watch([activeCellVal, activeSheetVal], () => {
  if (!activeCellVal.value) { formulaDraft.value = ""; return }
  formulaDraft.value = cellDisplayValue(activeCellVal.value.row, activeCellVal.value.col)
})
function commitFormula() {
  if (!activeCellVal.value || isReadOnlyVal.value) return
  const key = cellKey(activeCellVal.value.row, activeCellVal.value.col)
  const previous = cellDisplayValue(activeCellVal.value.row, activeCellVal.value.col)
  const next = formulaDraft.value
  if (next) cellData.value[key] = next
  else delete cellData.value[key]
  if (previous !== next) {
    localHistory.value.push({ key, previous, next })
    localFuture.value = []
  }
  ctrl.value.setSelectedFormula(next)
  ribbonStatus.value = `Updated ${activeCellAddressText.value}`
}
function cancelFormula() {
  formulaDraft.value = formulaFocusValue.value
}
function onCellClick(ri: number, ci: number, event?: MouseEvent) {
  ctrl.value.selectCell({ row: ri, col: ci }, { extend: !!event?.shiftKey })
}
function onCellMouseDown(ri: number, ci: number, event: MouseEvent) {
  if (event.button !== 0) return
  const cell = { row: ri, col: ci }
  dragAnchor.value = cell
  isDraggingSelection.value = true
  ctrl.value.selectCell(cell, { extend: event.shiftKey })
  window.addEventListener("mouseup", finishDragSelection, { once: true })
}
function onCellMouseEnter(ri: number, ci: number) {
  if (!isDraggingSelection.value || !dragAnchor.value) return
  ctrl.value.selectRange({ start: dragAnchor.value, end: { row: ri, col: ci } })
}
function onCellDblClick(ri: number, ci: number) {
  if (ctrl.value.readOnly || props.readOnly) return
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
  if (ctrl.value.readOnly || props.readOnly) { cancelEdit(); return }
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

const sheetTabsStyle = computed(() => ({
  height: '32px', display: 'flex', alignItems: 'stretch',
  borderTop: `1px solid ${isDarkRef.value ? '#374151' : '#e5e7eb'}`,
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
  const sel = normalizedRange(selectionVal.value)
  const isSel = sel && ci >= sel.start.col && ci <= sel.end.col
  return {
    position: 'sticky' as const,
    top: 0,
    left: isFrozenCol(ci) ? `${frozenColLeft(ci)}px` : undefined,
    zIndex: isFrozenCol(ci) ? 4 : 2,
    background: isSel ? (isDarkRef.value ? '#1e3a5f' : '#dbeafe') : (isDarkRef.value ? '#374151' : '#e5e7eb'),
    border: `1px solid ${isDarkRef.value ? '#4b5563' : '#d1d5db'}`,
    height: `${HEADER_HEIGHT}px`, cursor: 'pointer', textAlign: 'center' as const,
    fontSize: '11px', fontWeight: 400, color: isDarkRef.value ? '#e5e7eb' : '#374151',
    userSelect: 'none' as const, padding: 0,
  }
}
function rowHeaderStyle(ri: number) {
  const sel = normalizedRange(selectionVal.value)
  const isSel = sel && ri >= sel.start.row && ri <= sel.end.row
  return {
    position: 'sticky' as const,
    left: 0,
    top: isFrozenRow(ri) ? `${frozenRowTop(ri)}px` : undefined,
    zIndex: isFrozenRow(ri) ? 4 : 1,
    background: isSel ? (isDarkRef.value ? '#1e3a5f' : '#dbeafe') : (isDarkRef.value ? '#374151' : '#e5e7eb'),
    border: `1px solid ${isDarkRef.value ? '#4b5563' : '#d1d5db'}`,
    height: `${rowHeight(ri)}px`, cursor: 'pointer', textAlign: 'center' as const,
    fontSize: '11px', fontWeight: 400, color: isDarkRef.value ? '#e5e7eb' : '#374151',
    userSelect: 'none' as const, padding: 0,
  }
}
function cellStyle(ri: number, ci: number) {
  const ac = activeCellVal.value
  const sel = normalizedRange(selectionVal.value)
  const isActive = ac?.row === ri && ac?.col === ci
  const isInSel = sel ? ri >= sel.start.row && ri <= sel.end.row && ci >= sel.start.col && ci <= sel.end.col : false
  const sheet = activeSheetVal.value as unknown as { cellStyles?: Record<string, CellStylePatch> } | null
  const patch = cellStyles.value[cellKey(ri, ci)] ?? sheet?.cellStyles?.[cellKey(ri, ci)] ?? {}
  const font = patch.font as Record<string, unknown> | undefined
  const fill = patch.fill as { color?: { hex?: string } } | undefined
  const alignment = patch.alignment as Record<string, unknown> | undefined
  const border = patch.border as Record<string, unknown> | undefined
  const bg = fill?.color?.hex
    ? `#${fill.color.hex}`
    : isActive
      ? (isDarkRef.value ? '#1e3a5f' : '#dbeafe')
      : isInSel
        ? (isDarkRef.value ? 'rgba(59,130,246,0.15)' : 'rgba(59,130,246,0.08)')
        : (isDarkRef.value ? '#1f2937' : '#ffffff')
  return {
    border: border ? `2px solid #${borderColor.value.replace('#', '')}` : `1px solid ${isDarkRef.value ? '#374151' : '#e5e7eb'}`,
    height: `${rowHeight(ri)}px`, padding: '0 4px', overflow: 'hidden', cursor: 'cell', position: (isFrozenRow(ri) || isFrozenCol(ci)) ? 'sticky' as const : 'relative' as const, left: isFrozenCol(ci) ? `${frozenColLeft(ci)}px` : undefined, top: isFrozenRow(ri) ? `${frozenRowTop(ri)}px` : undefined, zIndex: isFrozenRow(ri) && isFrozenCol(ci) ? 3 : (isFrozenRow(ri) || isFrozenCol(ci) ? 2 : 0),
    fontSize: `${font?.size ?? 12}px`,
    fontFamily: String(font?.name ?? 'inherit'),
    fontWeight: font?.bold ? 700 : 400,
    fontStyle: font?.italic ? 'italic' : 'normal',
    textDecoration: `${font?.underline ? 'underline ' : ''}${font?.strikethrough ? 'line-through' : ''}`.trim() || 'none',
    color: (font?.color as { hex?: string } | undefined)?.hex ? `#${(font!.color as { hex: string }).hex}` : (isDarkRef.value ? '#e5e7eb' : '#111827'),
    textAlign: String(alignment?.horizontal ?? 'left') as 'left' | 'center' | 'right',
    verticalAlign: String(alignment?.vertical ?? 'middle'),
    whiteSpace: alignment?.wrapText ? 'normal' : 'nowrap',
    background: highlightCells.value && !fill?.color?.hex ? '#fff7ed' : bg,
    outline: isActive ? `2px solid ${isDarkRef.value ? '#60a5fa' : '#2563eb'}` : 'none',
    outlineOffset: '-2px',
    userSelect: isDraggingSelection.value ? ('none' as const) : ('auto' as const),
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
.xlsx-viewer-toolbar {
  display: flex;
  flex-shrink: 0;
  flex-direction: column;
  gap: 6px;
  border-bottom: 1px solid var(--border, #e5e7eb);
  background: color-mix(in oklch, var(--background, #fff) 85%, var(--muted, #f5f5f5));
  padding: 6px;
}
.xlsx-title-cluster {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  min-width: 0;
}
.xlsx-file-name {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--foreground, #111827);
  font-size: 13px;
  font-weight: 600;
}
.xlsx-loading { color: var(--muted-foreground, #737373); font-size: 12px; }
.xlsx-ribbon-tabs { display: flex; gap: 4px; overflow-x: auto; }
.ribbon-tab {
  border: 1px solid transparent;
  border-radius: var(--radius-md, 8px);
  background: transparent;
  color: var(--muted-foreground, #737373);
  cursor: pointer;
  font-size: 12px;
  padding: 5px 10px;
  white-space: nowrap;
}
.ribbon-tab:hover,
.ribbon-tab.active { background: var(--background, #fff); color: var(--foreground, #111827); border-color: var(--border, #e5e7eb); }
.ribbon-message-group { min-width: min(520px, 100%); justify-content: flex-start; }
.ribbon-message { max-width: 480px; color: var(--foreground, #111827); font-size: 12px; line-height: 1.35; }
.xlsx-toolbar-groups { display: flex; align-items: stretch; gap: 6px; overflow-x: auto; padding-bottom: 1px; }
.ribbon-group {
  display: inline-flex;
  min-height: 56px;
  min-width: fit-content;
  align-items: center;
  gap: 4px;
  border: 1px solid var(--border, #e5e7eb);
  border-radius: var(--radius, 10px);
  background: color-mix(in oklch, var(--background, #fff) 92%, transparent);
  box-shadow: var(--shadow-sm, 0 1px 2px rgb(0 0 0 / 0.05));
  padding: 6px;
  position: relative;
  padding-bottom: 18px;
}
.ribbon-label {
  position: absolute;
  bottom: 4px;
  left: 6px;
  right: 6px;
  color: var(--muted-foreground, #737373);
  font-size: 10px;
  font-weight: 500;
  text-align: center;
}
.ribbon-button {
  display: inline-flex;
  height: 28px;
  min-width: 28px;
  align-items: center;
  justify-content: center;
  border: 1px solid var(--border, #e5e7eb);
  border-radius: var(--radius-md, 8px);
  background: var(--background, #fff);
  color: var(--foreground, #111827);
  cursor: pointer;
  font-size: 12px;
  padding: 0 8px;
}
.ribbon-button.wide { min-width: 44px; }
.ribbon-button:hover:not(:disabled) { background: var(--accent, #f5f5f5); }
.ribbon-button:disabled { cursor: not-allowed; opacity: 0.45; }
.zoom-value { color: var(--foreground, #111827); font-size: 12px; min-width: 42px; text-align: center; }
.xlsx-viewer-grid :deep(button),
.xlsx-viewer-grid :deep(input),
.xlsx-viewer-grid :deep(select),
.xlsx-viewer-sheet-tabs :deep(button) {
  border-radius: 0;
}
.xlsx-viewer-table {
  width: max-content;
  min-width: 100%;
  border-radius: 0 !important;
  border-collapse: collapse !important;
  border-spacing: 0 !important;
  box-shadow: none !important;
}
.xlsx-viewer-table th,
.xlsx-viewer-table td {
  border-radius: 0 !important;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.xlsx-viewer-sheet-tabs button { position: relative; }
.xlsx-sheet-thumbnail {
  display: none;
  position: absolute;
  bottom: calc(100% + 6px);
  left: 0;
  width: 144px;
  height: 88px;
  border: 1px solid #d1d5db;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.18);
  pointer-events: none;
}
.xlsx-viewer-sheet-tabs button:hover .xlsx-sheet-thumbnail,
.xlsx-viewer-sheet-tabs button:focus-visible .xlsx-sheet-thumbnail { display: block; }
.xlsx-resize-handle { position: absolute; display: block; background: transparent; }
.col-resize { top: 0; right: -3px; width: 6px; height: 100%; cursor: col-resize; }
.row-resize { left: 0; right: 0; bottom: -3px; height: 6px; cursor: row-resize; }

.xlsx-data-bar {
  position: absolute;
  inset: 3px auto 3px 3px;
  opacity: 0.28;
  border-radius: 3px;
  pointer-events: none;
  z-index: 0;
}
.xlsx-validation-indicator,
.xlsx-formula-badge,
.xlsx-sparkline {
  position: absolute;
  right: 3px;
  top: 2px;
  z-index: 2;
  font-size: 9px;
  line-height: 1;
  color: #2563eb;
  pointer-events: none;
}
.xlsx-formula-badge { color: #7c3aed; font-weight: 700; }
.xlsx-sparkline {
  inset: 2px 4px auto auto;
  font-size: 13px;
  color: #059669;
  letter-spacing: -1px;
}

</style>
