<template>
  <div
    ref="tableEl"
    class="docx-table-host docx-viewer-table"
    :class="{
      'docx-table-host--editable': editable && !isReadOnly,
      'docx-table-host--active': isActiveTable,
    }"
    :style="wrapperStyle"
    :data-docx-table-host="true"
    :data-docx-table-index="domTableIndex"
    :data-docx-table-runtime-key="tableRuntimeKey"
  >
    <!-- Column resize handles -->
    <div
      v-if="showColumnHandles"
      class="docx-table-column-handles"
      data-docx-thumbnail-exclude="true"
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
          v-for="row of gridRows"
          :key="`row-${row.rowIndex}`"
          :class="rowClass(row.rowIndex)"
          :data-docx-tracked-change="rowTrackedKind(row.rowIndex)"
          :data-docx-tracked-change-id="rowTrackedId(row.rowIndex)"
          :style="rowStyle(row.row, row.rowIndex)"
        >
          <template
            v-for="cell of row.cells"
            :key="`cell-${cell.rowIndex}-${cell.cellIndex}`"
          >
          <td
            :colspan="cell.colSpan > 1 ? cell.colSpan : undefined"
            :rowspan="cell.rowSpan > 1 ? cell.rowSpan : undefined"
            :data-docx-table-cell="true"
            :data-docx-table-index="domTableIndex"
            :data-docx-table-runtime-key="tableRuntimeKey"
            :data-docx-table-row-index="cell.rowIndex"
            :data-docx-table-cell-index="cell.cellIndex"
            class="docx-table-cell"
            :class="{
              'docx-table-cell--active': isCellActive(cell.rowIndex, cell.cellIndex),
              'docx-table-cell--merge-fallback': cell.continuationFallback,
              'docx-table-cell--contain-floating': cellContainsLayoutInCellFloatingImage(cell),
              'docx-table-cell--insertion': cellTrackedKind(cell.rowIndex, cell.cellIndex) === 'insertion',
              'docx-table-cell--deletion': cellTrackedKind(cell.rowIndex, cell.cellIndex) === 'deletion',
            }"
            :data-docx-tracked-change="cellTrackedKind(cell.rowIndex, cell.cellIndex)"
            :data-docx-tracked-change-id="cellTrackedId(cell.rowIndex, cell.cellIndex)"
            :style="cellStyle(cell)"
            @click.stop="onCellClick(cell.rowIndex, cell.cellIndex, $event)"
          >
            <template v-if="!cell.continuationFallback">
            <template
              v-for="(node, pi) of cell.cell.nodes"
              :key="`cell-node-${cell.rowIndex}-${cell.cellIndex}-${pi}`"
            >
              <div
                v-if="node.type === 'paragraph'"
                class="docx-table-cell-paragraph"
                :contenteditable="editable && controller && !isReadOnly ? 'true' : undefined"
                :data-docx-table-cell-paragraph-host="true"
                :data-docx-table-index="domTableIndex"
                :data-docx-table-runtime-key="tableRuntimeKey"
                :data-docx-table-row-index="cell.rowIndex"
                :data-docx-table-cell-index="cell.cellIndex"
                :data-docx-paragraph-index="pi"
                :suppresscontenteditablewarning="editable && controller ? 'true' : undefined"
                :style="cellParagraphStyle(node)"
                @input="onCellInput(tableIndex, cell.rowIndex, cell.cellIndex, pi, $event)"
                @focus="onCellFocus(tableIndex, cell.rowIndex, cell.cellIndex, pi)"
                @blur="onCellBlur(tableIndex, cell.rowIndex, cell.cellIndex, pi, $event)"
                v-html="cellParagraphHtml(node, cell, pi)"
              />
              <DocxTableHost
                v-else
                :table="node"
                :table-index="tableIndex"
                :table-runtime-key="nestedTableRuntimeKey(cell.rowIndex, cell.cellIndex, pi)"
                :editable="editable && !isReadOnly"
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
                :max-available-width-px="cellContentWidthPx(cell)"
                :note-marker-labels="noteMarkerLabels"
              />
            </template>
            </template>
          </td>
          </template>
        </tr>
      </tbody>
    </table>

    <!-- Row add button (below table, only when editable) -->
    <div
      v-if="editable && !isReadOnly && !isNestedTable"
      class="docx-table-add-row"
      data-docx-thumbnail-exclude="true"
      @click="onAddRow"
    >
      + Add row
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onBeforeUnmount, onMounted, onUnmounted } from "vue"
import {
  mergeTableSpacing,
  resolveTableCellBorderCss,
  resolveTableCellDiagonalOverlayCss,
  resolveTableRenderMetricsPx,
  resolveTableRowHeightCss,
  resolveTableSpacingPaddingPx,
  tableSpacingPaddingStyle,
  tableElementBorderStyle,
  tableWrapperStyle,
  twipsToPixels,
  type DocxDocumentTheme,
  type DocxEditorController,
  type DocxEditorSelection,
  type NumberingDefinitionSet,
  type ParagraphNode,
  type TableNode,
  type TableRowNode,
  type TableRowRange,
  paragraphText,
  paragraphTypographyStyle,
  resolveTableTrackedMarkup,
} from "@arcships/docx-core"
import { renderParagraphRuns, type ParagraphRunRenderOptions } from "../render/paragraph-runs"
import { renderStaticHtml } from "../render/static-html"
import { readDocxTextHostText } from "../composables/editor-selection"
import { resolveTableGrid, type TableGridCell } from "./table-grid"

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
    tableRuntimeKey?: string
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
    maxAvailableWidthPx?: number
    noteMarkerLabels?: { footnote?: Map<number, string>; endnote?: Map<number, string> }
  }>(),
  {
    editable: false,
    documentTheme: "light" as DocxDocumentTheme,
  }
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

// Cell text drafts. Each editable cell paragraph keeps its uncommitted DOM
// text here until blur, export, or unmount; committing the whole cell's
// textContent would duplicate every paragraph into the first one.
interface CellDraftEntry {
  text: string
  baseText: string
  documentLoadNonce: number
}
const cellDrafts = new Map<string, CellDraftEntry>()
let unregisterDraftFlusher: (() => void) | undefined

function cellDraftKey(rowIndex: number, cellIndex: number, paragraphIndex: number): string {
  return `${rowIndex}:${cellIndex}:${paragraphIndex}`
}

function parseCellDraftKey(
  key: string
): { rowIndex: number; cellIndex: number; paragraphIndex: number } | undefined {
  const [row, cell, paragraph] = key.split(":").map((value) => Number(value))
  if (!Number.isInteger(row) || !Number.isInteger(cell) || !Number.isInteger(paragraph)) {
    return undefined
  }
  return { rowIndex: row, cellIndex: cell, paragraphIndex: paragraph }
}

function cellParagraphAt(
  rowIndex: number, cellIndex: number, paragraphIndex: number
): ParagraphNode | undefined {
  // Nested tables carry their own node; the host table index points at the
  // parent table and cannot be used to resolve inner row/cell coordinates.
  const modelTable = props.controller?.model.nodes[props.tableIndex]
  const table = props.tableRuntimeKey
    ? props.table
    : modelTable?.type === "table"
      ? modelTable
      : props.table
  const cell = table.rows[rowIndex]?.cells[cellIndex]
  if (!cell) return undefined
  return cell.nodes.filter((node): node is ParagraphNode => node.type === "paragraph")[paragraphIndex]
}

function readCellParagraphDomText(root: HTMLElement): string {
  // Shared reader: skips list markers, image children and — while revisions are
  // on screen — inline deleted/moved-from runs, which must not be committed.
  return readDocxTextHostText(root)
}

function commitCellParagraph(
  rowIndex: number, cellIndex: number, paragraphIndex: number, text: string
): void {
  const controller = props.controller
  if (!controller) return
  const paragraph = cellParagraphAt(rowIndex, cellIndex, paragraphIndex)
  if (!paragraph || paragraphText(paragraph) === text) return
  if (props.tableRuntimeKey) {
    const nestedController = controller as DocxEditorController &
      RuntimeKeyTableCellController
    nestedController.commitTableCellParagraphTextByRuntimeKey?.(
      props.tableRuntimeKey, rowIndex, cellIndex, paragraphIndex, text
    )
    return
  }
  controller.commitTableCellParagraphTextRecursive(
    props.tableIndex, rowIndex, cellIndex, paragraphIndex, text
  )
}

function flushCellDrafts(): void {
  const controller = props.controller
  if (!controller) return
  for (const [key, draft] of [...cellDrafts.entries()]) {
    const parsed = parseCellDraftKey(key)
    cellDrafts.delete(key)
    if (!parsed) continue
    const paragraph = cellParagraphAt(parsed.rowIndex, parsed.cellIndex, parsed.paragraphIndex)
    if (
      !paragraph ||
      draft.documentLoadNonce !== controller.documentLoadNonce ||
      paragraphText(paragraph) !== draft.baseText ||
      paragraphText(paragraph) === draft.text
    ) continue
    commitCellParagraph(
      parsed.rowIndex, parsed.cellIndex, parsed.paragraphIndex, draft.text
    )
  }
}
// ── Computed ───────────────────────────────────────────────────────
type RuntimeKeyTableCellController = {
  selectTableCellByRuntimeKey?: (
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number
  ) => void
  commitTableCellTextByRuntimeKey?: (
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number,
    text: string
  ) => void
  commitTableCellParagraphTextByRuntimeKey?: (
    tableRuntimeKey: string,
    rowIndex: number,
    cellIndex: number,
    paragraphIndex: number,
    text: string
  ) => void
}

const isNestedTable = computed(() => Boolean(props.tableRuntimeKey))

const domTableIndex = computed(() => props.tableRuntimeKey ?? props.tableIndex)

// Visible revisions no longer imply a read-only table: the document stays
// editable and revisions are resolved through accept/reject.
const isReadOnly = computed(() => !props.editable || !props.controller)

function selectionRuntimeKey(
  selection: Extract<DocxEditorSelection, { kind: "table-cell" }>
): string | undefined {
  return (selection as { tableRuntimeKey?: string }).tableRuntimeKey
}

const isActiveTable = computed(() => {
  if (!props.editable || !props.controller) return false
  const sel = props.controller.selection
  if (sel.kind === "table-cell") {
    return sel.tableIndex === props.tableIndex &&
      selectionRuntimeKey(sel) === props.tableRuntimeKey
  }
  return false
})

const tableMetrics = computed(() =>
  resolveTableRenderMetricsPx(props.table, props.maxAvailableWidthPx)
)

const columnCount = computed(() => tableMetrics.value.columnCount)

const wrapperStyle = computed(() =>
  tableWrapperStyle(props.table, twipsToPixels(props.table.style?.indentTwips) ?? 0)
)

const renderedTableWidthPx = computed(() => {
  if (localColumnWidths.value && localColumnWidths.value.length === columnCount.value) {
    return localColumnWidths.value.reduce((sum, width) => sum + width, 0)
  }
  return tableMetrics.value.tableWidthPx
})

const resolvedColumnWidths = computed(() => {
  if (localColumnWidths.value && localColumnWidths.value.length === columnCount.value) {
    return localColumnWidths.value
  }
  return tableMetrics.value.columnWidthsPx.map((width) => Math.max(1, Math.round(width)))
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
  () =>
    props.editable &&
    !isReadOnly.value &&
    !isNestedTable.value &&
    (hovered.value || isActiveTable.value)
)

const gridRows = computed(() =>
  resolveTableGrid(props.table, {
    startRowIndex: props.rowRange?.startRowIndex,
    endRowIndex: props.rowRange?.endRowIndex,
  })
)

const tableTrackedMarkup = computed(() =>
  props.showTrackedChanges ? resolveTableTrackedMarkup(props.table) : undefined
)

function rowTrackedRevision(rowIndex: number) {
  return tableTrackedMarkup.value?.rowChangeByIndex.get(rowIndex)
}

function cellTrackedRevision(rowIndex: number, cellIndex: number) {
  return tableTrackedMarkup.value?.cellChangeByIndex.get(
    `${rowIndex}:${cellIndex}`
  )
}

function rowTrackedKind(rowIndex: number): string | undefined {
  return rowTrackedRevision(rowIndex)?.kind
}

function rowTrackedId(rowIndex: number): string | undefined {
  return rowTrackedRevision(rowIndex)?.stableId
}

function cellTrackedKind(rowIndex: number, cellIndex: number): string | undefined {
  return cellTrackedRevision(rowIndex, cellIndex)?.kind
}

function cellTrackedId(rowIndex: number, cellIndex: number): string | undefined {
  return cellTrackedRevision(rowIndex, cellIndex)?.stableId
}

function rowClass(rowIndex: number): string | undefined {
  const kind = rowTrackedKind(rowIndex)
  return kind === "insertion" || kind === "deletion"
    ? `docx-table-row--${kind}`
    : undefined
}

const tableStyle = computed(() => ({
  width: `${Math.max(1, Math.round(renderedTableWidthPx.value))}px`,
  maxWidth: "100%",
  tableLayout: "fixed" as const,
  ...tableElementBorderStyle(props.table),
}))

// ── Cell helpers ───────────────────────────────────────────────────
function rowStyle(
  row: TableRowNode,
  rowIndex: number
): Record<string, string | number | undefined> {
  const style: Record<string, string | number | undefined> = {}
  const explicitHeightPx = twipsToPixels(row.style?.heightTwips)
  const heightCss = resolveTableRowHeightCss(row, explicitHeightPx)
  if (heightCss) Object.assign(style, heightCss)
  if (row.style?.backgroundColor) style.backgroundColor = row.style.backgroundColor
  if (row.style?.cantSplit) style.breakInside = "avoid"
  const trackedKind = rowTrackedKind(rowIndex)
  if (trackedKind === "deletion") {
    style.backgroundColor = "rgba(185, 28, 28, 0.07)"
    style.opacity = 0.72
  } else if (trackedKind === "insertion") {
    style.backgroundColor = "rgba(220, 38, 38, 0.08)"
  }
  return style
}

// Border resolution needs the grid columns a cell covers, so a merged cell is
// treated as one outer edge. The materialized grid already knows them.
function cellStyle(cell: TableGridCell): Record<string, string | number | undefined> {
  const startColumnIndex = cell.columnIndex
  const endColumnIndex = Math.min(
    Math.max(0, columnCount.value - 1),
    cell.columnIndex + cell.colSpan - 1
  )
  const borders = resolveTableCellBorderCss(
    props.table.style?.borders,
    cell.cell.style?.borders,
    cell.rowIndex,
    props.table.rows.length,
    startColumnIndex,
    endColumnIndex,
    columnCount.value
  )
  const trackedKind = cellTrackedKind(cell.rowIndex, cell.cellIndex)
  return {
    ...borders,
    ...tableSpacingPaddingStyle(
      mergeTableSpacing(
        props.table.style?.cellMarginTwips,
        cell.cell.style?.marginTwips
      )
    ),
    ...resolveTableCellDiagonalOverlayCss(
      props.table.style?.borders,
      cell.cell.style?.borders
    ),
    verticalAlign: cell.cell.style?.verticalAlign ?? "top",
    minWidth: "0",
    backgroundColor:
      trackedKind === "deletion"
        ? "rgba(185, 28, 28, 0.08)"
        : trackedKind === "insertion"
        ? "rgba(220, 38, 38, 0.08)"
        : (cell.cell.style?.backgroundColor as string | undefined),
    color: trackedKind === "deletion" ? "#b91c1c" : undefined,
    textDecoration: trackedKind === "deletion" ? "line-through" : undefined,
    opacity: trackedKind === "deletion" ? 0.8 : undefined,
    wordWrap: "break-word",
    overflowWrap: "break-word",
    wordBreak: "break-word",
  }
}

// Width a nested table may occupy: the parent cell box minus its padding.
function cellContentWidthPx(cell: TableGridCell): number | undefined {
  const endColumnIndex = Math.min(
    Math.max(0, columnCount.value - 1),
    cell.columnIndex + cell.colSpan - 1
  )
  const padding = resolveTableSpacingPaddingPx(
    mergeTableSpacing(
      props.table.style?.cellMarginTwips,
      cell.cell.style?.marginTwips
    )
  )
  const spannedWidthPx = resolvedColumnWidths.value
    .slice(cell.columnIndex, endColumnIndex + 1)
    .reduce((sum, width) => sum + width, 0)
  return Math.max(1, Math.round(spannedWidthPx - padding.left - padding.right))
}

function cellParagraphStyle(para: ParagraphNode): Record<string, any> {
  return {
    margin: "0",
    ...paragraphTypographyStyle(para.style),
    fontWeight: para.style?.headingLevel ? "700" : undefined,
    minHeight: "1em",
    outline: "none",
  }
}

function cellContainsLayoutInCellFloatingImage(cell: TableGridCell): boolean {
  return cell.cell.nodes.some((node) =>
    node.type === "paragraph" &&
    node.children.some(
      (child) =>
        child.type === "image" &&
        Boolean(child.floating) &&
        child.floating?.layoutInCell !== false
    )
  )
}

function cellParagraphHtml(
  para: ParagraphNode,
  cell: TableGridCell,
  pi: number
): string {
  const runOptions: ParagraphRunRenderOptions = {
    showTrackedChanges: props.showTrackedChanges,
    showCommentHighlights: props.showCommentHighlights,
    numberingDefinitions: props.numberingDefinitions,
    withinHeaderFooter: props.withinHeaderFooter,
    headerFooterRegion: props.headerFooterRegion,
    pageNumberFormat: props.pageNumberFormat,
    noteMarkerLabels: props.noteMarkerLabels,
    tableCellLayout: { widthPx: cellContentWidthPx(cell) },
  }
  const rendered = renderParagraphRuns(
    para,
    `table-${domTableIndex.value}-r${cell.rowIndex}-c${cell.cellIndex}-p${pi}`,
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
      selectionRuntimeKey(sel) === props.tableRuntimeKey &&
      sel.rowIndex === rowIndex &&
      sel.cellIndex === cellIndex
  }
  return false
}

function nestedTableRuntimeKey(
  rowIndex: number,
  cellIndex: number,
  contentIndex: number
): string {
  if (props.tableRuntimeKey) {
    return `${props.tableRuntimeKey}-nested-table-${rowIndex}-${cellIndex}-${contentIndex}`
  }
  return `body-cell-nested-table-${props.tableIndex}-${rowIndex}-${cellIndex}-${contentIndex}`
}

function selectCell(rowIndex: number, cellIndex: number): void {
  if (!props.controller) return
  if (props.tableRuntimeKey) {
    const nestedController = props.controller as DocxEditorController &
      RuntimeKeyTableCellController
    if (nestedController.selectTableCellByRuntimeKey) {
      nestedController.selectTableCellByRuntimeKey(
        props.tableRuntimeKey,
        rowIndex,
        cellIndex
      )
      return
    }
  }
  props.controller.selectTableCell(props.tableIndex, rowIndex, cellIndex)
}

// ── Event handlers ─────────────────────────────────────────────────
function onCellClick(rowIndex: number, cellIndex: number, _event: MouseEvent): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  selectCell(rowIndex, cellIndex)
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
  const text = readCellParagraphDomText(target)
  const key = cellDraftKey(rowIndex, cellIndex, paragraphIndex)
  const existing = cellDrafts.get(key)
  const paragraph = cellParagraphAt(rowIndex, cellIndex, paragraphIndex)
  if (paragraph) {
    cellDrafts.set(key, {
      text,
      baseText: existing?.baseText ?? paragraphText(paragraph),
      documentLoadNonce: existing?.documentLoadNonce ?? props.controller.documentLoadNonce,
    })
  }
  emit("cellInput", tableIndex, rowIndex, cellIndex, paragraphIndex, text)
}

function onCellFocus(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number
): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  selectCell(rowIndex, cellIndex)
  emit("cellFocus", tableIndex, rowIndex, cellIndex, paragraphIndex)
}

function onCellBlur(
  _tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number,
  event?: Event
): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  const key = cellDraftKey(rowIndex, cellIndex, paragraphIndex)
  const draft = cellDrafts.get(key)
  cellDrafts.delete(key)
  if (draft && draft.documentLoadNonce !== props.controller.documentLoadNonce) return

  const target = event?.target as HTMLElement | undefined
  const runtimeKeySelector = props.tableRuntimeKey
    ? `[data-docx-table-runtime-key="${props.tableRuntimeKey}"]`
    : ""
  const element = target ?? tableEl.value?.querySelector<HTMLElement>(
    `[data-docx-table-cell-paragraph-host="true"]${runtimeKeySelector}` +
    `[data-docx-table-row-index="${rowIndex}"]` +
    `[data-docx-table-cell-index="${cellIndex}"]` +
    `[data-docx-paragraph-index="${paragraphIndex}"]`
  )
  const text = element ? readCellParagraphDomText(element) : draft?.text
  if (text === undefined) return
  commitCellParagraph(rowIndex, cellIndex, paragraphIndex, text)
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

onMounted(() => {
  unregisterDraftFlusher = props.controller?.registerTextDraftFlusher(flushCellDrafts)
})

onBeforeUnmount(() => {
  flushCellDrafts()
})

onUnmounted(() => {
  unregisterDraftFlusher?.()
  unregisterDraftFlusher = undefined
  cleanupColumnResize(false)
})
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
  border-collapse: collapse;
  table-layout: fixed;
}
.docx-table-cell {
  vertical-align: top;
  min-width: 0;
  word-wrap: break-word;
  overflow-wrap: break-word;
  word-break: break-word;
}
.docx-table-cell--active {
  background: rgba(59, 130, 246, 0.04);
}
.docx-table-row--insertion .docx-table-cell {
  box-shadow: inset 0 2px 0 rgba(220, 38, 38, 0.55);
}
.docx-table-row--deletion .docx-table-cell {
  text-decoration: line-through;
}
.docx-table-row--deletion .docx-table-cell-paragraph {
  opacity: 0.72;
}
.docx-table-cell--deletion {
  text-decoration: line-through;
}
.docx-table-cell--insertion {
  box-shadow: inset 0 2px 0 rgba(220, 38, 38, 0.55);
}
.docx-table-cell--merge-fallback {
  background: transparent;
}
.docx-table-cell--contain-floating {
  position: relative;
  overflow: hidden;
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
