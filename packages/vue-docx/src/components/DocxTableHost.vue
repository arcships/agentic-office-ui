<template>
  <div
    ref="tableEl"
    class="docx-table-host docx-viewer-table"
    :class="{
      'docx-table-host--editable': editable && !isReadOnly,
      'docx-table-host--active': isActiveTable,
      'docx-table-host--bookmark-target': bookmarkTarget,
    }"
    :style="wrapperStyle"
    :data-docx-table-host="true"
    :data-docx-table-index="domTableIndex"
    :data-docx-table-node-index="tableIndex"
    :data-docx-table-runtime-key="tableRuntimeKey"
    :data-docx-table-nesting-path="tableNestingPath || undefined"
    @pointerenter="onTablePointerEnter"
    @pointerleave="onTablePointerLeave"
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

    <!-- Row resize handles -->
    <div
      v-if="showRowHandles"
      class="docx-table-row-handles"
      data-docx-thumbnail-exclude="true"
      data-docx-table-row-handles="true"
    >
      <div
        v-for="boundary of rowBoundaries"
        :key="`row-handle-${boundary.rowIndex}`"
        class="docx-table-row-handle"
        :data-docx-table-row-handle="boundary.rowIndex"
        :style="{ top: `${boundary.bottomPx}px` }"
        @mousedown.prevent="onRowResizeStart(boundary.rowIndex, $event)"
      />
    </div>

    <!-- Table move handle (top-left) and the drag ghost it moves -->
    <div
      v-if="showMoveHandle"
      class="docx-table-move-handle"
      data-docx-table-move-handle="true"
      data-docx-thumbnail-exclude="true"
      :title="m.docx.editor.moveTableHandle"
      :aria-label="m.docx.editor.moveTableHandle"
      @mousedown.prevent="onTableMoveStart($event)"
    />
    <div
      v-if="moveGhost"
      class="docx-table-drag-ghost"
      data-docx-table-drag-ghost="true"
      data-docx-thumbnail-exclude="true"
      :style="moveGhostStyle"
    />

    <!-- Table -->
    <table
      class="docx-table-host-table"
      :style="tableStyle"
      :data-docx-source-id="sourceId"
    >
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
          :data-docx-table-row-index="row.rowIndex"
          :data-docx-source-id="rowSourceId(row.rowIndex)"
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
            :data-docx-source-id="cellSourceId(cell.rowIndex, cell.cellIndex)"
            :data-docx-table-runtime-key="tableRuntimeKey"
            :data-docx-table-row-index="cell.rowIndex"
            :data-docx-table-cell-index="cell.cellIndex"
            class="docx-table-cell"
            :class="{
              'docx-table-cell--active': isCellActive(cell.rowIndex, cell.cellIndex),
              'docx-table-cell--range-selected': isCellInSelectedRange(cell.rowIndex, cell.cellIndex),
              'docx-table-cell--range-focus': isCellRangeFocus(cell.rowIndex, cell.cellIndex),
              'docx-table-cell--merge-fallback': cell.continuationFallback,
              'docx-table-cell--contain-floating': cellContainsLayoutInCellFloatingImage(cell),
              'docx-table-cell--insertion': cellTrackedKind(cell.rowIndex, cell.cellIndex) === 'insertion',
              'docx-table-cell--deletion': cellTrackedKind(cell.rowIndex, cell.cellIndex) === 'deletion',
              'docx-table-cell--bookmark-target': isBookmarkTargetCell(cell.rowIndex, cell.cellIndex),
            }"
            :data-docx-tracked-change="cellTrackedKind(cell.rowIndex, cell.cellIndex)"
            :data-docx-tracked-change-id="cellTrackedId(cell.rowIndex, cell.cellIndex)"
            :style="{ ...cellStyle(cell), ...tableCellSliceStyle(cell.rowIndex) }"
            @mousedown="onCellMouseDown(cell.rowIndex, cell.cellIndex, $event)"
            @mouseenter="onCellMouseEnter(cell.rowIndex, cell.cellIndex, $event)"
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
                :data-docx-table-node-index="tableIndex"
                :data-docx-source-id="cellParagraphSourceId(cell, pi)"
                :data-docx-table-runtime-key="tableRuntimeKey"
                :data-docx-table-row-index="cell.rowIndex"
                :data-docx-table-cell-index="cell.cellIndex"
                :data-docx-paragraph-index="cellParagraphIndex(cell.cell.nodes, pi)"
                :suppresscontenteditablewarning="editable && controller ? 'true' : undefined"
                :style="{ ...cellParagraphStyle(node), ...cellParagraphSliceStyle(cell.rowIndex) }"
                @input="onCellInput(tableIndex, cell.rowIndex, cell.cellIndex, cellParagraphIndex(cell.cell.nodes, pi), $event)"
                @focus="onCellFocus(tableIndex, cell.rowIndex, cell.cellIndex, cellParagraphIndex(cell.cell.nodes, pi))"
                @blur="onCellBlur(tableIndex, cell.rowIndex, cell.cellIndex, cellParagraphIndex(cell.cell.nodes, pi), $event)"
                v-html="cellParagraphHtml(node, cell, pi)"
              />
              <DocxTableHost
                v-else
                :table="node"
                :table-index="tableIndex"
                :source-id="cellNodeSourceId(cell, pi)"
                :table-runtime-key="nestedTableRuntimeKey(cell.rowIndex, cell.cellIndex, pi)"
                :nesting-path="nestedTableNestingPath(cell.rowIndex, cell.cellIndex, pi)"
                :editable="editable && !isReadOnly"
                :controller="controller"
                :document-theme="documentTheme"
                :show-tracked-changes="showTrackedChanges"
                :revision-display="effectiveRevisionDisplay"
                :show-comment-highlights="showCommentHighlights"
                :numbering-definitions="numberingDefinitions"
                :page-number="pageNumber"
                :total-pages="totalPages"
                :page-number-format="pageNumberFormat"
                :section-pages="sectionPages"
                :resolve-style-ref-field-value="resolveStyleRefFieldValue"
                :resolve-bookmark-field-value="resolveBookmarkFieldValue"
                :resolve-bookmark-page-number="resolveBookmarkPageNumber"
                :model="model"
                :within-header-footer="withinHeaderFooter"
                :header-footer-region="headerFooterRegion"
                :max-available-width-px="cellContentWidthPx(cell)"
                :note-marker-labels="noteMarkerLabels"
                :floating-page-origin-px="floatingPageOriginPx"
                :on-internal-link-click="onInternalLinkClick"
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
import { ref, computed, watch, nextTick, onBeforeUnmount, onMounted, onUnmounted } from "vue"
import {
  mergeTableSpacing,
  resolveTableCellBorderCss,
  resolveTableCellDiagonalOverlayCss,
  resolveTableFloatingDropPatch,
  resolveTableRenderMetricsPx,
  resolveTableRowHeightCss,
  resolveTableSpacingPaddingPx,
  tableSpacingPaddingStyle,
  tableElementBorderStyle,
  tableWrapperStyle,
  twipsToPixels,
  type DocxFloatingPageOriginPx,
  type DocModel,
  type DocxDocumentTheme,
  type DocxEditorController,
  type DocxEditorSelection,
  type DocxRevisionDisplay,
  type DocxTableCellSelectionRange,
  type NumberingDefinitionSet,
  type ParagraphNode,
  type TableNode,
  type TableRowNode,
  type TableRowRange,
  paragraphText,
  paragraphTypographyStyle,
  isCellSelected,
  isCellWithinTableSelectionRange,
  tableCellSelectionRangeHasMultipleCells,
  resolveDefaultTabStopPx,
  resolveTableTrackedMarkup,
  docxTableCellSourceId,
  docxTableRowSourceId,
  docxCellParagraphSourceId,
  docxNestedTableSourceId,
} from "@arcships/docx-core"
import { renderParagraphRuns, type ParagraphRunRenderOptions } from "../render/paragraph-runs"
import { renderStaticHtml } from "../render/static-html"
import { readDocxTextHostText } from "../composables/editor-selection"
import { useDocxMessages } from "../i18n"
import { docxBookmarkNameFromClick } from "../composables/docx-bookmarks"
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
const MIN_ROW_HEIGHT_PX = 20
const FALLBACK_ROW_HEIGHT_PX = 24
/** Pointer travel below this stays a click: the handle then only selects. */
const TABLE_MOVE_THRESHOLD_PX = 3

defineOptions({ name: "DocxTableHost" })

// ── Props ──────────────────────────────────────────────────────────
const props = withDefaults(
  defineProps<{
    table: TableNode
    tableIndex: number
    /** Source id consumed by block/word geometry (`data-docx-source-id`). */
    sourceId?: string
    editable?: boolean
    controller?: DocxEditorController
    tableRuntimeKey?: string
    /**
     * Slash-separated `rowIndex-cellIndex-contentIndex` steps from the
     * top-level table to this nested table. Search uses it to project matches
     * into the DOM independently of the page-specific runtime key.
     */
    nestingPath?: string
    rowRange?: TableRowRange
    /** Partial row slice produced by the pagination engine. */
    rowSlice?: {
      rowIndex: number
      startOffsetPx: number
      sliceHeightPx: number
      totalRowHeightPx: number
    }
    documentTheme?: DocxDocumentTheme
    showTrackedChanges?: boolean
    /** `inline`, `balloons` or `none`; omitted keeps historical inline markup. */
    revisionDisplay?: DocxRevisionDisplay
    showCommentHighlights?: boolean
    numberingDefinitions?: NumberingDefinitionSet
    pageNumber?: number
    totalPages?: number
    pageNumberFormat?: string
    sectionPages?: number
    resolveStyleRefFieldValue?: (
      target: string,
      beforeNodeIndex?: number,
    ) => string | undefined
    resolveBookmarkFieldValue?: (target: string) => string | undefined
    resolveBookmarkPageNumber?: (target: string) => number | undefined
    model?: DocModel
    /** Document `w:defaultTabStop`; resolved from `model` when omitted. */
    defaultTabStopPx?: number
    withinHeaderFooter?: boolean
    headerFooterRegion?: "header" | "footer"
    maxAvailableWidthPx?: number
    noteMarkerLabels?: { footnote?: Map<number, string>; endnote?: Map<number, string> }
    /** Page-relative origin used to place margin/page/column anchored images. */
    floatingPageOriginPx?: DocxFloatingPageOriginPx
    /** Row resize handles; hosts that manage row heights themselves set false. */
    showRowResizeHandles?: boolean
    /** Table move handle (drag to float the table); set false to render none. */
    showTableMoveHandle?: boolean
    /** Internal `#bookmark` link clicked inside a cell paragraph. */
    onInternalLinkClick?: (bookmarkName: string) => void
    /** Transient highlight after the viewer revealed a bookmark in this table. */
    bookmarkTarget?: boolean
    /** Cell coordinates of a table bookmark, when known. */
    bookmarkCell?: { rowIndex: number; cellIndex: number }
  }>(),
  {
    editable: false,
    documentTheme: "light" as DocxDocumentTheme,
    revisionDisplay: "inline" as DocxRevisionDisplay,
    showRowResizeHandles: true,
    showTableMoveHandle: true,
  }
)

const emit = defineEmits<{
  cellInput: [tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string]
  cellFocus: [tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number]
  measure: [tableIndex: number, rowHeightsPx: number[]]
}>()

const m = useDocxMessages()

// Cell paragraphs are measured and rendered on the document's tab grid.
const documentTabStopPx = computed(() =>
  Number.isFinite(props.defaultTabStopPx)
    ? Math.max(1, Math.round(props.defaultTabStopPx as number))
    : resolveDefaultTabStopPx(props.model?.metadata?.defaultTabStopTwips)
)

// ── State ──────────────────────────────────────────────────────────
const tableEl = ref<HTMLElement | undefined>()
let rowResizeObserver: ResizeObserver | undefined
const hovered = ref(false)
const isResizing = ref(false)
const resizeBoundaryIndex = ref(-1)
const resizeStartX = ref(0)
const resizeStartWidths = ref<number[]>([])

// Local column width overrides (from user resizing)
const localColumnWidths = ref<number[] | null>(null)
let activeColumnMouseMove: ((event: MouseEvent) => void) | undefined
let activeColumnMouseUp: (() => void) | undefined

// Local row height overrides (from user resizing) and the measured row heights
// the row handles are positioned from.
const localRowHeights = ref<Record<number, number>>({})
const measuredRowHeights = ref<number[]>([])
const tableOffsetTopPx = ref(0)
const isRowResizing = ref(false)
let activeRowMouseMove: ((event: MouseEvent) => void) | undefined
let activeRowMouseUp: (() => void) | undefined
let rowResizeTarget: { rowIndex: number; heightPx: number } | undefined

// Table move drag: live pointer offset while the handle is held, committed as a
// floating position on drop.
const isMovingTable = ref(false)
const moveDelta = ref<{ x: number; y: number } | null>(null)
let moveStartClientX = 0
let moveStartClientY = 0
let activeMoveMouseMove: ((event: MouseEvent) => void) | undefined
let activeMoveMouseUp: (() => void) | undefined

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

function cellParagraphIndex(
  nodes: TableGridCell["cell"]["nodes"],
  contentIndex: number
): number {
  // `paragraphIndex` addresses paragraphs, not all cell content nodes. Nested
  // tables occupy content slots but must not shift the index of a following
  // paragraph, otherwise the host commits into the wrong paragraph.
  let paragraphIndex = 0
  for (let index = 0; index < contentIndex; index += 1) {
    if (nodes[index]?.type === "paragraph") paragraphIndex += 1
  }
  return paragraphIndex
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

const tableNestingPath = computed(() => props.nestingPath ?? "")

function nestedTableNestingPath(
  rowIndex: number,
  cellIndex: number,
  contentIndex: number,
): string {
  const step = `${rowIndex}-${cellIndex}-${contentIndex}`
  return tableNestingPath.value ? `${tableNestingPath.value}/${step}` : step
}

// Source ids shared with `extractDocumentBlocks`; geometry resolves the rendered
// table, rows, cells and cell paragraphs through these attributes.
function rowSourceId(rowIndex: number): string | undefined {
  return props.sourceId ? docxTableRowSourceId(props.sourceId, rowIndex) : undefined
}

function cellSourceId(rowIndex: number, cellIndex: number): string | undefined {
  return props.sourceId
    ? docxTableCellSourceId(props.sourceId, rowIndex, cellIndex)
    : undefined
}

function cellParagraphSourceId(
  cell: TableGridCell,
  contentIndex: number,
): string | undefined {
  const cellId = cellSourceId(cell.rowIndex, cell.cellIndex)
  return cellId
    ? docxCellParagraphSourceId(cellId, cellParagraphIndex(cell.cell.nodes, contentIndex))
    : undefined
}

function cellNodeSourceId(
  cell: TableGridCell,
  contentIndex: number,
): string | undefined {
  const cellId = cellSourceId(cell.rowIndex, cell.cellIndex)
  return cellId ? docxNestedTableSourceId(cellId, contentIndex) : undefined
}

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

const showRowHandles = computed(
  () =>
    props.showRowResizeHandles !== false &&
    props.editable &&
    !isReadOnly.value &&
    !isNestedTable.value &&
    (hovered.value || isActiveTable.value)
)

/**
 * Tables that can float: a body table, on its first page fragment only. A
 * nested table lives inside a cell, and a header/footer table has no body node
 * `setTableFloating` could address.
 */
const isFloatableTable = computed(
  () =>
    !isNestedTable.value &&
    !props.withinHeaderFooter &&
    (props.rowRange?.startRowIndex ?? 0) === 0
)

const showMoveHandle = computed(
  () =>
    props.showTableMoveHandle !== false &&
    isFloatableTable.value &&
    props.editable &&
    !isReadOnly.value &&
    (hovered.value || isActiveTable.value)
)

const moveGhost = computed(() => isMovingTable.value && moveDelta.value !== null)

const moveGhostStyle = computed<Record<string, string>>(() => {
  const delta = moveDelta.value ?? { x: 0, y: 0 }
  const boundaries = rowBoundaries.value
  const lastBottomPx = boundaries[boundaries.length - 1]?.bottomPx ?? 0
  return {
    width: `${Math.max(1, Math.round(renderedTableWidthPx.value))}px`,
    height: `${Math.max(1, Math.round(lastBottomPx - tableOffsetTopPx.value))}px`,
    top: `${Math.round(tableOffsetTopPx.value)}px`,
    left: "0px",
    transform: `translate(${Math.round(delta.x)}px, ${Math.round(delta.y)}px)`,
  }
})

/**
 * Bottom edge of every rendered row, in table-host coordinates. Measured row
 * heights win over the model's `w:trHeight`, because a row that grew to fit its
 * content is taller than the height stored in the document.
 */
const rowBoundaries = computed<Array<{ rowIndex: number; bottomPx: number; heightPx: number }>>(
  () => {
    let top = tableOffsetTopPx.value
    return gridRows.value.map((row) => {
      const measured = measuredRowHeights.value[row.rowIndex]
      const modelHeightPx = twipsToPixels(row.row.style?.heightTwips)
      const heightPx = Math.max(
        MIN_ROW_HEIGHT_PX,
        Math.round(
          localRowHeights.value[row.rowIndex] ??
            (Number.isFinite(measured) && (measured as number) > 0
              ? (measured as number)
              : Number.isFinite(modelHeightPx) && (modelHeightPx as number) > 0
                ? (modelHeightPx as number)
                : FALLBACK_ROW_HEIGHT_PX)
        )
      )
      top += heightPx
      return { rowIndex: row.rowIndex, bottomPx: top, heightPx }
    })
  }
)

const gridRows = computed(() => {
  const range = props.rowRange
  const bodyRows = resolveTableGrid(props.table, {
    startRowIndex: range?.startRowIndex,
    endRowIndex: range?.endRowIndex,
  })
  const repeatHeaderRowCount = Math.max(
    0,
    Math.round(range?.repeatHeaderRowCount ?? 0)
  )
  const rangeStartsInBody = (range?.startRowIndex ?? 0) > 0
  if (
    repeatHeaderRowCount === 0 ||
    !rangeStartsInBody ||
    repeatHeaderRowCount >= (range?.startRowIndex ?? 0)
  ) {
    return bodyRows
  }
  const headerRows = resolveTableGrid(props.table, {
    startRowIndex: 0,
    endRowIndex: repeatHeaderRowCount,
  })
  return [...headerRows, ...bodyRows]
})

const tableTrackedMarkup = computed(() =>
  props.showTrackedChanges && props.revisionDisplay !== "none"
    ? resolveTableTrackedMarkup(props.table)
    : undefined
)

// Hidden revisions mean the final text, regardless of the preferred markup
// mode chosen for the gutter selector.
const effectiveRevisionDisplay = computed<DocxRevisionDisplay>(() =>
  props.showTrackedChanges ? props.revisionDisplay : "none"
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
function rowSlicedFrame(rowIndex: number) {
  return props.rowSlice?.rowIndex === rowIndex ? props.rowSlice : undefined
}

function tableCellSliceStyle(rowIndex: number): Record<string, string | number | undefined> {
  const slice = rowSlicedFrame(rowIndex)
  if (!slice) return {}
  return {
    height: `${Math.max(1, Math.round(slice.sliceHeightPx))}px`,
    maxHeight: `${Math.max(1, Math.round(slice.sliceHeightPx))}px`,
    overflow: "hidden",
    paddingTop: "0px",
    paddingBottom: "0px",
  }
}

function cellParagraphSliceStyle(rowIndex: number): Record<string, string | number | undefined> {
  const slice = rowSlicedFrame(rowIndex)
  if (!slice) return {}
  return { marginTop: `${-Math.max(0, Math.round(slice.startOffsetPx))}px` }
}

function rowStyle(
  row: TableRowNode,
  rowIndex: number
): Record<string, string | number | undefined> {
  const style: Record<string, string | number | undefined> = {}
  const explicitHeightPx = twipsToPixels(row.style?.heightTwips)
  const heightCss = resolveTableRowHeightCss(row, explicitHeightPx)
  if (heightCss) Object.assign(style, heightCss)
  const draggedHeightPx = localRowHeights.value[rowIndex]
  if (draggedHeightPx !== undefined) {
    // Live drag feedback only; the committed height comes from the model.
    style.height = `${Math.round(draggedHeightPx)}px`
    style.minHeight = `${Math.round(draggedHeightPx)}px`
  }
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
    revisionDisplay: effectiveRevisionDisplay.value,
    showCommentHighlights: props.showCommentHighlights,
    numberingDefinitions: props.numberingDefinitions,
    paragraphSourceId: cellParagraphSourceId(cell, pi),
    emitTextRangeAttributes: true,
    withinHeaderFooter: props.withinHeaderFooter,
    headerFooterRegion: props.headerFooterRegion,
    pageNumberFormat: props.pageNumberFormat,
    sectionPages: props.sectionPages,
    resolveStyleRefFieldValue: props.resolveStyleRefFieldValue
      ? (target: string) =>
          props.resolveStyleRefFieldValue?.(target, undefined)
      : undefined,
    resolveBookmarkFieldValue: props.resolveBookmarkFieldValue,
    resolveBookmarkPageNumber: props.resolveBookmarkPageNumber,
    noteMarkerLabels: props.noteMarkerLabels,
    tableCellLayout: { widthPx: cellContentWidthPx(cell) },
    imageAltText: m.value.docx.image.alt,
    imageMissingText: m.value.docx.image.missing,
    defaultTabStopPx: documentTabStopPx.value,
    emitLinkRangeAttributes:
      props.editable && !isReadOnly.value && !props.withinHeaderFooter,
  }
  const rendered = renderParagraphRuns(
    para,
    `table-${domTableIndex.value}-r${cell.rowIndex}-c${cell.cellIndex}-p${pi}`,
    props.documentTheme,
    undefined,
    props.onInternalLinkClick,
    props.floatingPageOriginPx,
    undefined,
    props.pageNumber,
    props.totalPages,
    runOptions
  )
  const nodes = Array.isArray(rendered) ? rendered : [rendered]
  return nodes.map((node) => renderStaticHtml(node)).join("")
}

function isBookmarkTargetCell(rowIndex: number, cellIndex: number): boolean {
  return Boolean(
    props.bookmarkTarget &&
    props.bookmarkCell &&
    props.bookmarkCell.rowIndex === rowIndex &&
    props.bookmarkCell.cellIndex === cellIndex,
  )
}

function isCellActive(rowIndex: number, cellIndex: number): boolean {
  if (!props.editable || !props.controller) return false
  const sel = props.controller.selection
  if (sel.kind !== "table-cell") return false
  if (
    sel.tableIndex !== props.tableIndex ||
    selectionRuntimeKey(sel) !== props.tableRuntimeKey
  ) {
    return false
  }
  return isCellSelected(sel, props.tableIndex, rowIndex, cellIndex, props.table)
}

/** True for every cell the current multi-cell range highlights. */
function isCellInSelectedRange(rowIndex: number, cellIndex: number): boolean {
  if (!props.editable || !props.controller) return false
  const sel = props.controller.selection
  if (
    sel.kind !== "table-cell" ||
    sel.tableIndex !== props.tableIndex ||
    selectionRuntimeKey(sel) !== props.tableRuntimeKey ||
    !sel.cellRange ||
    !tableCellSelectionRangeHasMultipleCells(sel.cellRange)
  ) {
    return false
  }
  return isCellWithinTableSelectionRange(
    props.table,
    sel.cellRange,
    rowIndex,
    cellIndex
  )
}

/** The active (focus) cell of a multi-cell range. */
function isCellRangeFocus(rowIndex: number, cellIndex: number): boolean {
  if (!props.editable || !props.controller) return false
  const sel = props.controller.selection
  return Boolean(
    sel.kind === "table-cell" &&
      sel.tableIndex === props.tableIndex &&
      selectionRuntimeKey(sel) === props.tableRuntimeKey &&
      sel.cellRange &&
      tableCellSelectionRangeHasMultipleCells(sel.cellRange) &&
      sel.cellRange.focusRowIndex === rowIndex &&
      sel.cellRange.focusCellIndex === cellIndex
  )
}

function cellInsideRange(
  rowIndex: number,
  cellIndex: number,
  range: DocxTableCellSelectionRange
): boolean {
  return isCellWithinTableSelectionRange(
    props.table,
    range,
    rowIndex,
    cellIndex
  )
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
function onCellClick(rowIndex: number, cellIndex: number, event: MouseEvent): void {
  const bookmarkName = docxBookmarkNameFromClick(event?.target)
  if (bookmarkName && props.onInternalLinkClick) {
    event.preventDefault()
    event.stopPropagation()
    props.onInternalLinkClick(bookmarkName)
    return
  }
  if (!props.editable || !props.controller || isReadOnly.value) return
  const sel = props.controller.selection
  if (
    sel.kind === "table-cell" &&
    sel.tableIndex === props.tableIndex &&
    selectionRuntimeKey(sel) === props.tableRuntimeKey &&
    sel.cellRange &&
    selectionRangeHasMultipleCells(sel.cellRange) &&
    cellInsideRange(rowIndex, cellIndex, sel.cellRange)
  ) {
    return
  }
  selectCell(rowIndex, cellIndex)
}

function selectionRangeHasMultipleCells(range: DocxTableCellSelectionRange): boolean {
  return tableCellSelectionRangeHasMultipleCells(range)
}

const rangeAnchor = ref<{ rowIndex: number; cellIndex: number } | null>(null)
const isRangeSelecting = ref(false)
let activeRangeMouseUp: (() => void) | undefined

function selectCellRange(
  anchor: { rowIndex: number; cellIndex: number },
  focus: { rowIndex: number; cellIndex: number }
): void {
  if (!props.controller) return
  props.controller.selectTableCellRange(props.tableIndex, {
    tableIndex: props.tableIndex,
    anchorRowIndex: anchor.rowIndex,
    anchorCellIndex: anchor.cellIndex,
    focusRowIndex: focus.rowIndex,
    focusCellIndex: focus.cellIndex,
  }, props.tableRuntimeKey)
}

function finishRangeSelection(): void {
  isRangeSelecting.value = false
  rangeAnchor.value = null
  if (activeRangeMouseUp) {
    document.removeEventListener("mouseup", activeRangeMouseUp)
    activeRangeMouseUp = undefined
  }
}

function onCellMouseDown(
  rowIndex: number,
  cellIndex: number,
  event: MouseEvent
): void {
  if (docxBookmarkNameFromClick(event?.target)) {
    event.preventDefault()
    return
  }
  if (!props.editable || !props.controller || isReadOnly.value) return
  if (event.button !== 0) return
  // Nested cell mousedowns bubble to the host table; keep the range anchor in
  // the table whose cell the pointer is actually over.
  event.stopPropagation?.()
  const target = event.target as HTMLElement | undefined
  const insideEditable = target?.closest?.("[contenteditable='true']")
  if (insideEditable && !event.shiftKey) return

  const selection = props.controller.selection
  const selectionMatchesTable =
    selection.kind === "table-cell" &&
    selection.tableIndex === props.tableIndex &&
    selectionRuntimeKey(selection) === props.tableRuntimeKey
  const currentRange =
    selectionMatchesTable && selection.kind === "table-cell"
      ? selection.cellRange
      : undefined
  // Shift extends from the range anchor, or from the single active cell when
  // no rectangle exists yet. Clicking a cell without Shift starts a new anchor.
  const anchor = event.shiftKey
    ? currentRange
      ? {
          rowIndex: currentRange.anchorRowIndex,
          cellIndex: currentRange.anchorCellIndex,
        }
      : selectionMatchesTable && selection.kind === "table-cell"
        ? { rowIndex: selection.rowIndex, cellIndex: selection.cellIndex }
        : { rowIndex, cellIndex }
    : { rowIndex, cellIndex }

  event.preventDefault()
  rangeAnchor.value = anchor
  isRangeSelecting.value = true
  selectCellRange(anchor, { rowIndex, cellIndex })

  if (!activeRangeMouseUp) {
    activeRangeMouseUp = () => finishRangeSelection()
    document.addEventListener("mouseup", activeRangeMouseUp)
  }
}

function onCellMouseEnter(
  rowIndex: number,
  cellIndex: number,
  event: MouseEvent
): void {
  if (!isRangeSelecting.value || !rangeAnchor.value) return
  if (event.buttons === 0) {
    finishRangeSelection()
    return
  }
  selectCellRange(rangeAnchor.value, { rowIndex, cellIndex })
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
  // Focusing a cell inside the selected rectangle must not collapse it: the
  // click that completes a drag or a Shift+click lands in the focus cell and
  // would otherwise discard the range before the next command runs.
  const selection = props.controller.selection
  if (
    selection.kind === "table-cell" &&
    selection.tableIndex === props.tableIndex &&
    selectionRuntimeKey(selection) === props.tableRuntimeKey &&
    selection.cellRange &&
    tableCellSelectionRangeHasMultipleCells(selection.cellRange) &&
    cellInsideRange(rowIndex, cellIndex, selection.cellRange)
  ) {
    emit("cellFocus", tableIndex, rowIndex, cellIndex, paragraphIndex)
    return
  }
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
  // Blur closes the typing burst: the next edit in this cell is a new undo step.
  props.controller.endTextCommitBurst?.()
}

function onTablePointerEnter(): void {
  if (!props.editable) return
  hovered.value = true
}

function onTablePointerLeave(): void {
  hovered.value = false
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

function onRowResizeStart(rowIndex: number, event: MouseEvent): void {
  if (!props.editable || !props.controller || isReadOnly.value) return
  cleanupRowResize(false)
  const startHeight =
    rowBoundaries.value.find((boundary) => boundary.rowIndex === rowIndex)?.heightPx ??
    FALLBACK_ROW_HEIGHT_PX
  const startY = event.clientY
  isRowResizing.value = true
  rowResizeTarget = { rowIndex, heightPx: startHeight }

  activeRowMouseMove = (e: MouseEvent) => {
    if (!isRowResizing.value) return
    const heightPx = Math.max(
      MIN_ROW_HEIGHT_PX,
      Math.round(startHeight + (e.clientY - startY))
    )
    // A drag that grows the row also has to grow the host, otherwise the
    // measured table height wins and the row snaps back mid-drag.
    localRowHeights.value = { ...localRowHeights.value, [rowIndex]: heightPx }
    rowResizeTarget = { rowIndex, heightPx }
  }

  activeRowMouseUp = () => cleanupRowResize(true)

  document.addEventListener("mousemove", activeRowMouseMove)
  document.addEventListener("mouseup", activeRowMouseUp)
}

function cleanupRowResize(commit: boolean): void {
  if (activeRowMouseMove) {
    document.removeEventListener("mousemove", activeRowMouseMove)
  }
  if (activeRowMouseUp) {
    document.removeEventListener("mouseup", activeRowMouseUp)
  }
  activeRowMouseMove = undefined
  activeRowMouseUp = undefined

  const target = rowResizeTarget
  const shouldCommit =
    commit && isRowResizing.value && target && props.controller && !props.tableRuntimeKey
  isRowResizing.value = false
  rowResizeTarget = undefined
  localRowHeights.value = {}
  if (shouldCommit && target) {
    props.controller!.setTableRowHeight(props.tableIndex, target.rowIndex, target.heightPx)
  }
}

function onAddRow(): void {
  const lastRowIndex = Math.max(0, props.table.rows.length - 1)
  props.controller?.insertTableRow(props.tableIndex, lastRowIndex, "below")
}

/**
 * Selects every cell of the table, the way Word's table selector does: the
 * table gets its active outline and the cell-range commands (merge, clear,
 * borders) apply to the whole table after a click on the move handle.
 */
function selectWholeTable(): void {
  const controller = props.controller
  if (!controller || isNestedTable.value) return
  const lastRowIndex = Math.max(0, props.table.rows.length - 1)
  const lastCellIndex = Math.max(0, (props.table.rows[lastRowIndex]?.cells.length ?? 1) - 1)
  controller.selectTableCellRange(
    props.tableIndex,
    {
      tableIndex: props.tableIndex,
      anchorRowIndex: 0,
      anchorCellIndex: 0,
      focusRowIndex: lastRowIndex,
      focusCellIndex: lastCellIndex,
    },
    props.tableRuntimeKey
  )
}

function onTableMoveStart(event: MouseEvent): void {
  if (!showMoveHandle.value) return
  if (typeof event.button === "number" && event.button !== 0) return
  event.preventDefault()
  selectWholeTable()
  cleanupTableMove(false)
  isMovingTable.value = true
  moveStartClientX = event.clientX
  moveStartClientY = event.clientY
  moveDelta.value = { x: 0, y: 0 }

  activeMoveMouseMove = (e: MouseEvent) => {
    if (!isMovingTable.value) return
    moveDelta.value = {
      x: e.clientX - moveStartClientX,
      y: e.clientY - moveStartClientY,
    }
  }
  activeMoveMouseUp = () => cleanupTableMove(true)

  document.addEventListener("mousemove", activeMoveMouseMove)
  document.addEventListener("mouseup", activeMoveMouseUp)
}

function cleanupTableMove(commit: boolean): void {
  if (activeMoveMouseMove) {
    document.removeEventListener("mousemove", activeMoveMouseMove)
  }
  if (activeMoveMouseUp) {
    document.removeEventListener("mouseup", activeMoveMouseUp)
  }
  activeMoveMouseMove = undefined
  activeMoveMouseUp = undefined

  const delta = moveDelta.value
  const shouldCommit =
    commit &&
    isMovingTable.value &&
    Boolean(delta) &&
    Boolean(props.controller) &&
    !isReadOnly.value &&
    isFloatableTable.value
  isMovingTable.value = false
  moveDelta.value = null
  if (!shouldCommit || !delta || !props.controller) return
  // A click (no real travel) only selects the table; it must not turn a table
  // floating at its own position.
  if (
    Math.abs(delta.x) < TABLE_MOVE_THRESHOLD_PX &&
    Math.abs(delta.y) < TABLE_MOVE_THRESHOLD_PX
  ) {
    return
  }

  props.controller.setTableFloating(
    props.tableIndex,
    resolveTableFloatingDropPatch({
      startXPx: twipsToPixels(props.table.style?.floating?.xTwips) ?? 0,
      startYPx: twipsToPixels(props.table.style?.floating?.yTwips) ?? 0,
      deltaXPx: delta.x,
      deltaYPx: delta.y,
      pageWidthPx: props.floatingPageOriginPx?.pageWidth,
      tableWidthPx: renderedTableWidthPx.value,
    })
  )
}

watch(
  () => props.table.style?.columnWidthsTwips,
  () => {
    if (!isResizing.value) localColumnWidths.value = null
  },
  { deep: true },
)

watch([() => props.editable, isReadOnly], ([editable, readOnly]) => {
  if (!editable || readOnly) {
    cleanupColumnResize(false)
    cleanupRowResize(false)
    cleanupTableMove(false)
  }
})

function emitMeasuredRowHeights(): void {
  const root = tableEl.value
  if (!root || typeof root.querySelectorAll !== "function") return
  const rowHeightsPx: number[] = []
  root
    .querySelectorAll<HTMLElement>("tbody tr[data-docx-table-row-index]")
    .forEach((rowElement) => {
      const rowIndex = Number(rowElement.getAttribute("data-docx-table-row-index"))
      const heightPx = rowElement.offsetHeight
      if (Number.isSafeInteger(rowIndex) && heightPx > 0) {
        rowHeightsPx[rowIndex] = Math.round(heightPx)
      }
    })
  const tableElement = root.querySelector?.("table") as HTMLElement | null | undefined
  const offsetTop = tableElement?.offsetTop
  if (Number.isFinite(offsetTop)) tableOffsetTopPx.value = Math.max(0, Math.round(offsetTop as number))
  if (rowHeightsPx.some((heightPx) => Number.isFinite(heightPx) && heightPx > 0)) {
    measuredRowHeights.value = rowHeightsPx
    emit("measure", props.tableIndex, rowHeightsPx)
  }
}

watch(
  () => [props.table.rows, props.maxAvailableWidthPx] as const,
  () => { void nextTick(emitMeasuredRowHeights) },
  { flush: "post" },
)

onMounted(() => {
  unregisterDraftFlusher = props.controller?.registerTextDraftFlusher(flushCellDrafts)
  void nextTick(emitMeasuredRowHeights)
  if (typeof ResizeObserver !== "undefined" && tableEl.value) {
    const observer = new ResizeObserver(() => emitMeasuredRowHeights())
    observer.observe(tableEl.value)
    rowResizeObserver = observer
  }
})

onBeforeUnmount(() => {
  flushCellDrafts()
  rowResizeObserver?.disconnect()
  rowResizeObserver = undefined
})

onUnmounted(() => {
  unregisterDraftFlusher?.()
  unregisterDraftFlusher = undefined
  finishRangeSelection()
  cleanupColumnResize(false)
  cleanupRowResize(false)
  cleanupTableMove(false)
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

.docx-table-row-handles {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  height: 100%;
  pointer-events: none;
  z-index: 10;
}
.docx-table-row-handle {
  position: absolute;
  left: 0;
  right: 0;
  height: 5px;
  margin-top: -2px;
  cursor: row-resize;
  pointer-events: auto;
  background: transparent;
}
.docx-table-row-handle:hover {
  background: #3b82f6;
}

.docx-table-move-handle {
  position: absolute;
  top: -12px;
  left: -12px;
  width: 12px;
  height: 12px;
  border: 1px solid #93c5fd;
  border-radius: 2px;
  background: #eff6ff;
  cursor: move;
  pointer-events: auto;
  z-index: 12;
}
.docx-table-move-handle:hover {
  background: #bfdbfe;
  border-color: #3b82f6;
}

.docx-table-drag-ghost {
  position: absolute;
  border: 1px dashed #3b82f6;
  border-radius: 2px;
  background: rgba(59, 130, 246, 0.08);
  pointer-events: none;
  z-index: 11;
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
.docx-table-cell--range-selected {
  background: rgba(59, 130, 246, 0.16);
}
.docx-table-cell--range-focus {
  box-shadow: inset 0 0 0 1px rgba(37, 99, 235, 0.55);
}
.docx-table-cell--bookmark-target {
  background: rgba(250, 204, 21, 0.22);
  box-shadow: inset 0 0 0 2px rgba(250, 204, 21, 0.55);
}
.docx-table-host--bookmark-target .docx-table-host-table {
  outline: 2px solid rgba(250, 204, 21, 0.55);
  outline-offset: 2px;
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
