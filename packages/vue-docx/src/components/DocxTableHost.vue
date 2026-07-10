<template>
  <div
    ref="tableEl"
    class="docx-table-host docx-viewer-table"
    :class="{
      'docx-table-host--editable': editable && !isReadOnly,
      'docx-table-host--active': isActiveTable,
    }"
    :data-docx-table-host="true"
    :data-docx-table-index="tableIndex"
  >
    <!-- Column resize handles -->
    <div
      v-if="showColumnHandles"
      class="docx-table-column-handles"
    >
      <div
        v-for="(boundary, ci) of columnBoundaries"
        :key="`col-handle-${ci}`"
        class="docx-table-column-handle"
        :style="{ left: `${boundary}px` }"
        @mousedown.prevent="onColumnResizeStart(ci, $event)"
      />
    </div>

    <!-- Table -->
    <table class="docx-table-host-table" :style="tableStyle">
      <colgroup>
        <col
          v-for="(width, ci) of resolvedColumnWidths"
          :key="`col-${ci}`"
          :style="{ width: `${width}px` }"
        />
      </colgroup>
      <tbody>
        <tr
          v-for="(row, ri) of visibleRows"
          :key="`row-${sourceRowIndex(ri)}`"
          :style="rowStyle(row, sourceRowIndex(ri))"
        >
          <td
            v-for="(cell, ci) of row.cells"
            :key="`cell-${sourceRowIndex(ri)}-${ci}`"
            :colspan="cell.style?.gridSpan"
            :data-docx-table-cell="true"
            :data-docx-table-index="tableIndex"
            :data-docx-table-row-index="sourceRowIndex(ri)"
            :data-docx-table-cell-index="ci"
            class="docx-table-cell"
            :class="{
              'docx-table-cell--active': isCellActive(sourceRowIndex(ri), ci),
            }"
            :style="cellStyle(cell)"
            @click="onCellClick(sourceRowIndex(ri), ci, $event)"
          >
            <template
              v-for="(node, pi) of cell.nodes"
              :key="`cell-node-${sourceRowIndex(ri)}-${ci}-${pi}`"
            >
              <div
                v-if="node.type === 'paragraph'"
                class="docx-table-cell-paragraph"
                :contenteditable="editable && controller && !isReadOnly ? 'true' : undefined"
                :data-docx-table-cell-paragraph-host="true"
                :data-docx-table-index="tableIndex"
                :data-docx-table-row-index="sourceRowIndex(ri)"
                :data-docx-table-cell-index="ci"
                :data-docx-paragraph-index="pi"
                :suppresscontenteditablewarning="editable && controller ? 'true' : undefined"
                :style="cellParagraphStyle(node)"
                @input="onCellInput(tableIndex, sourceRowIndex(ri), ci, pi, $event)"
                @focus="onCellFocus(tableIndex, sourceRowIndex(ri), ci, pi)"
                @blur="onCellBlur(tableIndex, sourceRowIndex(ri), ci, pi)"
                v-html="cellParagraphHtml(node, sourceRowIndex(ri), ci, pi)"
              />
              <DocxTableHost
                v-else
                :table="node"
                :table-index="tableIndex"
                :editable="false"
                :controller="controller"
                :document-theme="documentTheme"
                :show-tracked-changes="showTrackedChanges"
                :show-comment-highlights="showCommentHighlights"
                :numbering-definitions="numberingDefinitions"
                :page-number="pageNumber"
                :total-pages="totalPages"
                :page-number-format="pageNumberFormat"
                :within-header-footer="withinHeaderFooter"
                :header-footer-region="headerFooterRegion"
              />
            </template>
          </td>
        </tr>
      </tbody>
    </table>

    <!-- Row add button (below table, only when editable) -->
    <div
      v-if="editable && !isReadOnly"
      class="docx-table-add-row"
      @click="onAddRow"
    >
      + Add row
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onUnmounted } from "vue"
import type {
  DocxDocumentTheme,
  DocxEditorController,
  NumberingDefinitionSet,
  ParagraphNode,
  TableNode,
  TableRowRange,
} from "@extend-ai/docx-core"
import { renderParagraphRuns, type ParagraphRunRenderOptions } from "../render/paragraph-runs"
import { renderStaticHtml } from "../render/static-html"

// ── Constants ──────────────────────────────────────────────────────
const HIGHLIGHT_TO_CSS: Record<string, string> = {
  yellow: "#fff59d",
  green: "#bbf7d0",
  cyan: "#a5f3fc",
  magenta: "#f5d0fe",
  blue: "#bfdbfe",
  red: "#fecaca",
  black: "#111827",
  white: "#ffffff",
}
const SCRIPT_FONT_SCALE = 0.65
const DEFAULT_COLUMN_WIDTH = 100
const MIN_COLUMN_WIDTH = 48

defineOptions({ name: "DocxTableHost" })

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    table: TableNode
    tableIndex: number
    editable?: boolean
    controller?: DocxEditorController
    rowRange?: TableRowRange
    documentTheme?: DocxDocumentTheme
    showTrackedChanges?: boolean
    showCommentHighlights?: boolean
    numberingDefinitions?: NumberingDefinitionSet
    pageNumber?: number
    totalPages?: number
    pageNumberFormat?: string
    withinHeaderFooter?: boolean
    headerFooterRegion?: "header" | "footer"
  }>(),
  { editable: false, documentTheme: "light" as DocxDocumentTheme }
)

const emit = defineEmits<{
  cellInput: [tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string]
  cellFocus: [tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number]
}>()

// ── State ──────────────────────────────────────────────────────────
const tableEl = ref<HTMLElement | undefined>()
const hovered = ref(false)
const isResizing = ref(false)
const resizeBoundaryIndex = ref(-1)
const resizeStartX = ref(0)
const resizeStartWidths = ref<number[]>([])

// Local column width overrides (from user resizing)
const localColumnWidths = ref<number[] | null>(null)
let activeColumnMouseMove: ((event: MouseEvent) => void) | undefined
let activeColumnMouseUp: (() => void) | undefined

// ── Computed ───────────────────────────────────────────────────────
const isReadOnly = computed(
  () => !props.editable || !props.controller || props.controller.showTrackedChanges
)

const isActiveTable = computed(() => {
  if (!props.editable || !props.controller) return false
  const sel = props.controller.selection
  if (sel.kind === "table-cell") {
    return sel.tableIndex === props.tableIndex
  }
  return false
})

const columnCount = computed(() => {
  if (props.table.rows.length === 0) return 1
  return props.table.rows[0].cells.length
})

const resolvedColumnWidths = computed(() => {
  if (localColumnWidths.value && localColumnWidths.value.length === columnCount.value) {
    return localColumnWidths.value
  }
  const storedWidths = props.table.style?.columnWidthsTwips
  if (storedWidths?.length === columnCount.value) {
    return storedWidths.map((width) => Math.max(MIN_COLUMN_WIDTH, Math.round(width / 15)))
  }
  return Array.from({ length: columnCount.value }, () => DEFAULT_COLUMN_WIDTH)
})

const columnBoundaries = computed(() => {
  const offsets: number[] = []
  let offset = 0
  for (const w of resolvedColumnWidths.value) {
    offset += w
    offsets.push(offset)
  }
  // Remove the last boundary (right edge of last column)
  return offsets.slice(0, -1)
})

const showColumnHandles = computed(
  () => props.editable && !isReadOnly && (hovered.value || isActiveTable.value)
)

const visibleRowStart = computed(() =>
  Math.max(0, props.rowRange?.startRowIndex ?? 0)
)

const visibleRows = computed(() => {
  const start = visibleRowStart.value
  const end = Math.max(start, props.rowRange?.endRowIndex ?? props.table.rows.length)
  return props.table.rows.slice(start, end)
})

function sourceRowIndex(visibleRowIndex: number): number {
  return visibleRowStart.value + visibleRowIndex
}

const tableStyle = computed(() => ({
  width: "100%",
  borderCollapse: "collapse" as const,
  tableLayout: "fixed" as const,
}))

// ── Cell helpers ───────────────────────────────────────────────────
function rowStyle(_row: any, _ri: number): Record<string, any> {
  return {}
}

function cellStyle(cell: any): Record<string, any> {
  return {
    border: "1px solid #d1d5db",
    padding: "8px",
    verticalAlign: "top",
    minWidth: "0",
    backgroundColor: (cell as any)?.backgroundColor as string | undefined,
    wordWrap: "break-word",
    overflowWrap: "break-word",
    wordBreak: "break-word",
  }
}

function cellParagraphStyle(para: ParagraphNode): Record<string, any> {
  return {
    margin: "0",
    textAlign: para.style?.align,
    fontWeight: para.style?.headingLevel ? "700" : undefined,
    minHeight: "1em",
    outline: "none",
  }
}

function cellParagraphHtml(
  para: ParagraphNode,
  ri: number,
  ci: number,
  pi: number
): string {
  const runOptions: ParagraphRunRenderOptions = {
    showTrackedChanges: props.showTrackedChanges,
    showCommentHighlights: props.showCommentHighlights,
    numberingDefinitions: props.numberingDefinitions,
    withinHeaderFooter: props.withinHeaderFooter,
    headerFooterRegion: props.headerFooterRegion,
    pageNumberFormat: props.pageNumberFormat,
  }
  const rendered = renderParagraphRuns(
    para,
    `table-${props.tableIndex}-r${ri}-c${ci}-p${pi}`,
    props.documentTheme,
    undefined,
    undefined,
    undefined,
    undefined,
    props.pageNumber,
    props.totalPages,
    runOptions
  )
  const nodes = Array.isArray(rendered) ? rendered : [rendered]
  return nodes.map((node) => renderStaticHtml(node)).join("")
}

function isCellActive(rowIndex: number, cellIndex: number): boolean {
  if (!props.editable || !props.controller) return false
  const sel = props.controller.selection
  if (sel.kind === "table-cell") {
    return sel.tableIndex === props.tableIndex &&
      sel.rowIndex === rowIndex &&
      sel.cellIndex === cellIndex
  }
  return false
}

// ── Event handlers ─────────────────────────────────────────────────
function onCellClick(rowIndex: number, cellIndex: number, _event: MouseEvent): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  props.controller.selectTableCell(props.tableIndex, rowIndex, cellIndex)
}

function onCellInput(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number,
  event: Event
): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  const target = event.target as HTMLElement
  const text = target.textContent ?? ""
  emit("cellInput", tableIndex, rowIndex, cellIndex, paragraphIndex, text)
}

function onCellFocus(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number
): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  props.controller.selectTableCell(tableIndex, rowIndex, cellIndex)
  emit("cellFocus", tableIndex, rowIndex, cellIndex, paragraphIndex)
}

function onCellBlur(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  _paragraphIndex: number
): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  // Commit cell text
  const cellEl = tableEl.value?.querySelector(
    `[data-docx-table-cell="true"][data-docx-table-row-index="${rowIndex}"][data-docx-table-cell-index="${cellIndex}"]`
  )
  if (cellEl) {
    const text = cellEl.textContent ?? ""
    props.controller.commitTableCellText(tableIndex, rowIndex, cellIndex, text)
  }
}

function onColumnResizeStart(boundaryIndex: number, event: MouseEvent): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  cleanupColumnResize(false)
  isResizing.value = true
  resizeBoundaryIndex.value = boundaryIndex
  resizeStartX.value = event.clientX
  resizeStartWidths.value = [...resolvedColumnWidths.value]

  activeColumnMouseMove = (e: MouseEvent) => {
    if (!isResizing.value) return
    const delta = e.clientX - resizeStartX.value
    const newWidths = [...resizeStartWidths.value]

    // Adjust two adjacent columns
    const leftIdx = boundaryIndex
    const rightIdx = boundaryIndex + 1

    const leftNew = Math.max(MIN_COLUMN_WIDTH, (resizeStartWidths.value[leftIdx] ?? DEFAULT_COLUMN_WIDTH) + delta)
    const rightNew = Math.max(MIN_COLUMN_WIDTH, (resizeStartWidths.value[rightIdx] ?? DEFAULT_COLUMN_WIDTH) - delta)

    newWidths[leftIdx] = leftNew
    if (rightIdx < newWidths.length) {
      newWidths[rightIdx] = rightNew
    }

    localColumnWidths.value = newWidths
  }

  activeColumnMouseUp = () => {
    cleanupColumnResize(true)
  }

  document.addEventListener("mousemove", activeColumnMouseMove)
  document.addEventListener("mouseup", activeColumnMouseUp)
}

function cleanupColumnResize(commit: boolean): void {
  if (activeColumnMouseMove) {
    document.removeEventListener("mousemove", activeColumnMouseMove)
  }
  if (activeColumnMouseUp) {
    document.removeEventListener("mouseup", activeColumnMouseUp)
  }
  activeColumnMouseMove = undefined
  activeColumnMouseUp = undefined

  const widths = localColumnWidths.value ? [...localColumnWidths.value] : undefined
  const shouldCommit = commit && isResizing.value && widths && props.controller
  isResizing.value = false
  if (shouldCommit) props.controller!.setTableColumnWidths(props.tableIndex, widths)
}

function onAddRow(): void {
  props.controller?.insertTableRow(props.tableIndex, props.table.rows.length, "below")
}

watch(
  () => props.table.style?.columnWidthsTwips,
  () => {
    if (!isResizing.value) localColumnWidths.value = null
  },
  { deep: true },
)

watch([() => props.editable, isReadOnly], ([editable, readOnly]) => {
  if (!editable || readOnly) cleanupColumnResize(false)
})

onUnmounted(() => cleanupColumnResize(false))
</script>

<style scoped>
.docx-table-host {
  position: relative;
  margin-bottom: 8px;
}
.docx-table-host--editable:hover {
  outline: 1px solid #93c5fd;
  outline-offset: 2px;
}
.docx-table-host--active {
  outline: 2px solid #3b82f6;
  outline-offset: 2px;
}

.docx-table-column-handles {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  height: 100%;
  pointer-events: none;
  z-index: 10;
}
.docx-table-column-handle {
  position: absolute;
  top: 0;
  width: 4px;
  height: 100%;
  margin-left: -2px;
  cursor: col-resize;
  pointer-events: auto;
  background: transparent;
}
.docx-table-column-handle:hover {
  background: #3b82f6;
}

.docx-table-host-table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
}
.docx-table-cell {
  border: 1px solid #d1d5db;
  padding: 8px;
  vertical-align: top;
  min-width: 0;
  word-wrap: break-word;
  overflow-wrap: break-word;
  word-break: break-word;
}
.docx-table-cell--active {
  background: rgba(59, 130, 246, 0.04);
}
.docx-table-cell-paragraph {
  margin: 0;
  min-height: 1em;
  outline: none;
}
.docx-table-cell-paragraph[contenteditable="true"] {
  cursor: text;
}
.docx-table-add-row {
  margin-top: 4px;
  padding: 4px 8px;
  font-size: 12px;
  color: #6b7280;
  cursor: pointer;
  border: 1px dashed #d1d5db;
  border-radius: 4px;
  text-align: center;
}
.docx-table-add-row:hover {
  background: #f3f4f6;
  color: #374151;
}
</style>
