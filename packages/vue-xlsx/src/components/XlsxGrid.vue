<template>
  <div
    ref="containerRef"
    class="xlsx-grid"
    data-testid="xlsx-grid"
    :data-display-row-count="effectiveRowHeights.length"
    :data-display-column-count="effectiveColWidths.length"
    :data-first-column-width="effectiveColWidths[0] ?? DEFAULT_COL_WIDTH"
    :data-first-row-height="effectiveRowHeights[0] ?? DEFAULT_ROW_HEIGHT"
    :data-merged-region-count="activeSheet?.mergedRegions?.length ?? 0"
    :data-worker-backed="String(Boolean(controller.isWorkerBacked))"
    :data-cut-range="cutRangeAddress || undefined"
    role="grid"
    :aria-label="gridAriaLabel"
    :aria-rowcount="gridRowCount"
    :aria-colcount="gridColCount"
    :aria-readonly="gridAriaReadOnly"
    :aria-activedescendant="activeCellDescendantId"
    tabindex="0"
    :style="containerStyle"
    @keydown="onGridKeydown"
    @scroll="onScroll"
    @pointerdown="onPointerDown"
    @pointermove="onPointerMove"
    @pointerup="onPointerUp"
    @pointercancel="onPointerUp"
    @pointerleave="onPointerLeave"
    @dblclick="onDblClick"
  >
    <div
      v-if="activeCellAria"
      class="xlsx-grid__a11y"
      role="row"
      :aria-rowindex="activeCellAria.rowIndex"
    >
      <div
        :id="activeCellDescendantId"
        class="xlsx-grid__a11y-cell"
        role="gridcell"
        :aria-colindex="activeCellAria.colIndex"
        :aria-label="activeCellAria.label"
        aria-selected="true"
      >{{ activeCellAria.text }}</div>
    </div>
    <div
      class="xlsx-grid__scroll-space"
      data-testid="xlsx-grid-scroll-space"
      :style="scrollSpaceStyle"
    >
      <div class="xlsx-grid__viewport" :style="viewportStyle">
        <!-- Column header canvas -->
        <canvas
          ref="colHeaderCanvasRef"
          class="xlsx-grid__col-header"
          :style="colHeaderCanvasStyle"
        />
        <!-- Row header canvas -->
        <canvas
          ref="rowHeaderCanvasRef"
          class="xlsx-grid__row-header"
          :style="rowHeaderCanvasStyle"
        />
        <!-- Corner canvas -->
        <canvas
          ref="cornerCanvasRef"
          class="xlsx-grid__corner"
          :style="cornerCanvasStyle"
        />
        <!-- Main body canvas -->
        <canvas
          ref="bodyCanvasRef"
          class="xlsx-grid__body"
          :style="bodyCanvasStyle"
        />
        <button
          v-if="columnMenuTriggerStyle && !editingLocked"
          class="xlsx-grid__axis-menu-trigger"
          data-testid="xlsx-column-menu"
          type="button"
          tabindex="-1"
          :style="columnMenuTriggerStyle"
          :aria-label="m.xlsx.contextMenu.selectColumn"
          @pointerdown.stop.prevent="openHeaderMenu('column', $event)"
          @dblclick.stop.prevent
        >▾</button>
        <button
          v-if="rowMenuTriggerStyle && !editingLocked"
          class="xlsx-grid__axis-menu-trigger"
          data-testid="xlsx-row-menu"
          type="button"
          tabindex="-1"
          :style="rowMenuTriggerStyle"
          :aria-label="m.xlsx.contextMenu.selectRow"
          @pointerdown.stop.prevent="openHeaderMenu('row', $event)"
          @dblclick.stop.prevent
        >▾</button>
        <XlsxHeaderMenu
          v-if="headerMenu"
          :controller="controller"
          :kind="headerMenu.kind"
          :index="headerMenu.index"
          :x="headerMenu.x"
          :y="headerMenu.y"
          @close="headerMenu = null"
        />
        <div
          v-if="pageBreakOverlay"
          class="xlsx-grid__page-breaks"
          data-testid="xlsx-page-break-preview"
          aria-hidden="true"
        >
          <i
            v-for="(line, index) in pageBreakOverlay.lines"
            :key="`line-${index}`"
            class="xlsx-grid__page-break-line"
            :class="`xlsx-grid__page-break-line--${line.kind}`"
            :style="line.style"
          />
          <b
            v-for="(label, index) in pageBreakOverlay.labels"
            :key="`label-${index}`"
            class="xlsx-grid__page-break-label"
            :style="label.style"
          >{{ label.text }}</b>
        </div>
        <!-- Editing input (hidden unless editing). The textarea makes a draft
             multi-line; the hidden mirror under it measures how many wrapped
             lines the draft needs so the editor grows instead of scrolling. -->
        <template v-if="editingCell">
          <div
            ref="editMirrorRef"
            class="xlsx-grid__edit-mirror"
            data-testid="xlsx-cell-editor-mirror"
            aria-hidden="true"
            :style="editMirrorStyle"
          >{{ editingValue }}</div>
          <textarea
            ref="editInputRef"
            v-model="editingValue"
            class="xlsx-grid__edit-input"
            data-testid="xlsx-cell-editor"
            :data-editor-lines="editorLineCount"
            rows="1"
            wrap="soft"
            spellcheck="false"
            :style="editInputStyle"
            @keydown="onEditKeydown"
            @paste="onEditPaste"
            @compositionstart="onEditCompositionStart"
            @compositionend="onEditCompositionEnd"
            @blur="commitEdit"
          />
        </template>
        <div
          v-if="hoveredComment"
          class="xlsx-grid__comment-popover"
          data-testid="xlsx-comment-popover"
          :style="hoveredComment.style"
        >
          <strong v-if="hoveredComment.author">{{ hoveredComment.author }}</strong>
          <span>{{ hoveredComment.text }}</span>
        </div>
        <div
          v-if="validationHint"
          class="xlsx-grid__validation-hint"
          data-testid="xlsx-validation-hint"
          :style="validationHint.style"
        >
          <strong v-if="validationHint.title">{{ validationHint.title }}</strong>
          <span>{{ validationHint.message }}</span>
        </div>
        <button
          v-if="validationDropdownTriggerStyle"
          class="xlsx-grid__validation-trigger"
          data-testid="xlsx-validation-dropdown-trigger"
          :style="validationDropdownTriggerStyle"
          type="button"
          tabindex="-1"
          :aria-label="m.xlsx.grid.openValidation"
          @pointerdown.stop.prevent="openValidationDropdown"
          @dblclick.stop.prevent
        >
          ▾
        </button>
        <div
          v-if="validationDropdownOpen"
          class="xlsx-grid__validation-dropdown"
          data-testid="xlsx-validation-dropdown"
          role="listbox"
          :style="validationDropdownStyle"
          @pointerdown.stop.prevent
          @keydown.escape.stop.prevent="closeValidationDropdown"
        >
          <button
            v-for="option in validationDropdownOptions"
            :key="option"
            class="xlsx-grid__validation-option"
            data-testid="xlsx-validation-option"
            type="button"
            role="option"
            @click.stop="chooseValidationOption(option)"
          >
            {{ option }}
          </button>
        </div>
        <div
          v-if="validationFeedbackView"
          class="xlsx-grid__validation-alert"
          :class="`xlsx-grid__validation-alert--${validationFeedbackView.kind}`"
          data-testid="xlsx-validation-alert"
          role="alert"
          :style="validationFeedbackView.style"
        >
          <strong v-if="validationFeedbackView.title">{{ validationFeedbackView.title }}</strong>
          <span>{{ validationFeedbackView.message }}</span>
          <button
            class="xlsx-grid__validation-dismiss"
            data-testid="xlsx-validation-alert-dismiss"
            type="button"
            :aria-label="m.xlsx.grid.closeValidation"
            @pointerdown.stop.prevent
            @click.stop="props.controller.dismissDataValidationFeedback()"
          >
            ×
          </button>
        </div>
      </div>
    </div>
    <XlsxClipboardNotice :message="clipboardNotice" />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, onMounted, onUnmounted, nextTick, type CSSProperties } from "vue";
import {
  resolveWorkbookColor,
  resolveWorkbookFillStyle,
  type XlsxViewerController,
  type XlsxCellAddress,
  type XlsxCellRange,
  type XlsxCellStyleContext,
  type XlsxCellHyperlink,
  type XlsxCellComment,
  type XlsxConditionalFormatRule,
  type XlsxDataValidation,
  type XlsxDataValidationFeedback,
  type XlsxSparkline,
  type XlsxPrintLayout,
  type XlsxResolvedCellStyle,
  type XlsxSheetData,
} from "@arcships/xlsx-core";
import { sanitizeOfficeLink } from "@arcships/office-runtime/link-policy";
import { GRID_HEADER_HEIGHT_PX, GRID_ROW_HEADER_WIDTH_PX, gridZoomFactor } from "./grid-metrics";
import { nextGridA11yId } from "./grid-a11y";
import XlsxClipboardNotice from "./XlsxClipboardNotice.vue";
import { useClipboardCopyNotice } from "../composables/useClipboardCopyNotice";
import {
  sortConditionalRules as orderConditionalRules,
  createConditionalEnvironment,
  resolveColorScaleFill,
  resolveDataBarGeometry,
  resolveIconGeometry,
  withAlpha,
  type ConditionalEnvironment,
} from "../render/conditional-format";
import {
  cellInNormalizedRange,
  collectConditionalFormatValues,
  createScopedCache,
  normalizeConditionalFormatRanges,
  type NormalizedRange,
} from "./conditional-format-stats";
import { paneAxisContentOffset, paneViewOffset, resolvePaneContentPoint, resolvePaneOverlayRect, resolvePaneViewPoint, type PaneOverlayRect } from "./grid-pane-geometry";
import {
  MAX_CANVAS_DIMENSION_PX,
  cellKey,
  intersectsViewport,
  readDevicePixelRatio,
  resolveAxisResizeSize,
  resolveCanvasPixelRatio,
  resolveOverflowTextWidth,
  sheetCellKey,
} from "./grid-paint";
import { isDataValidationListDropdownVisible, localizeDataValidationFeedback } from "../composables/validation";
import { localizeFormattedNumber, parseDisplayedNumber } from "../composables/cell-input";
import { rangeToA1 } from "../composables/selection";
import { formatOfficeMessage } from "@arcships/office-interaction";
import { useXlsxMessages } from "../i18n";
import XlsxHeaderMenu from "./XlsxHeaderMenu.vue";

const props = defineProps<{
  controller: XlsxViewerController;
  getCellStyle?: ((cell: XlsxCellAddress, context?: XlsxCellStyleContext) => Partial<CSSProperties> | undefined) | null;
  isDark?: boolean;
  readOnly?: boolean;
  selectionColor?: string;
  selectionFillColor?: string;
}>();

const m = useXlsxMessages();

const emit = defineEmits<{
  cellDoubleClick: [cell: XlsxCellAddress];
  viewportChange: [viewport: { scrollLeft: number; scrollTop: number; width: number; height: number }];
}>();

const { notice: clipboardNotice, copySelection } = useClipboardCopyNotice(() => props.controller, {
  failureMessage: () => m.value.xlsx.clipboard.copyFailed,
});

// ── Constants ──────────────────────────────────────────────────────────
const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const HEADER_HEIGHT = GRID_HEADER_HEIGHT_PX;
const ROW_HEADER_WIDTH = GRID_ROW_HEADER_WIDTH_PX;
const OVERSCAN = 240;
const WORKER_ROW_BATCH_SIZE = 64;
const MIN_DISPLAY_ROWS = 200;
const MIN_DISPLAY_COLS = 50;
const AXIS_RESIZE_DRAG_THRESHOLD = 3;
// Inline editor chrome, in CSS pixels. A `<textarea>` sizes its content box
// between padding and border, so the auto-grown box height is
// `lines × lineHeight + vertical padding + border`.
const EDITOR_BORDER_PX = 2;
const EDITOR_PADDING_X_PX = 4;
const EDITOR_PADDING_Y_PX = 2;
const EDITOR_BOX_CHROME_PX = EDITOR_BORDER_PX * 2 + EDITOR_PADDING_Y_PX * 2;
const EDITOR_LINE_HEIGHT_RATIO = 1.25;
const EDITOR_MIN_LINE_HEIGHT_PX = 12;
// Excel's marching ants: a 4/4 dash whose phase advances a few times a second,
// slow enough to read and cheap enough to repaint while a cut is pending.
const CUT_ANTS_DASH = [4, 4];
const CUT_ANTS_DASH_PERIOD = CUT_ANTS_DASH[0] + CUT_ANTS_DASH[1];
const CUT_ANTS_STEP_MS = 120;

// ── Refs ───────────────────────────────────────────────────────────────
const containerRef = ref<HTMLDivElement | null>(null);
const bodyCanvasRef = ref<HTMLCanvasElement | null>(null);
const colHeaderCanvasRef = ref<HTMLCanvasElement | null>(null);
const rowHeaderCanvasRef = ref<HTMLCanvasElement | null>(null);
const cornerCanvasRef = ref<HTMLCanvasElement | null>(null);
const editInputRef = ref<HTMLTextAreaElement | null>(null);
const editMirrorRef = ref<HTMLDivElement | null>(null);
// Box height the mirror measured for the current draft. Kept separate from the
// style so growing the editor never feeds back into its own measurement.
const editorMeasuredHeight = ref(0);

const scrollTop = ref(0);
const scrollLeft = ref(0);
const containerWidth = ref(0);
const containerHeight = ref(0);

const editingCell = ref<XlsxCellAddress | null>(null);
const editingValue = ref("");
// Exposed so the viewer can keep sheet-level shortcuts (undo, Ctrl+PageUp/Down)
// out of an open inline editor.
const isEditingCell = computed(() => editingCell.value !== null);
const isEditComposing = ref(false);
let commitEditOnCompositionEnd = false;
const hoveredCell = ref<XlsxCellAddress | null>(null);
const hoveredAxis = ref<{
  actualIndex: number;
  displayIndex: number;
  kind: "column" | "row";
} | null>(null);
const headerMenu = ref<{
  index: number;
  kind: "column" | "row";
  x: number;
  y: number;
} | null>(null);
const validationDropdownOpen = ref(false);
const resizePreview = ref<{
  actualIndex: number;
  displayIndex: number;
  kind: "column" | "row";
  moved: boolean;
  pointerId: number;
  size: number;
  startPosition: number;
  startSize: number;
} | null>(null);

const isSelecting = ref(false);
const selectionAnchor = ref<XlsxCellAddress | null>(null);
const cutAntsStep = ref(0);
let cutAntsFrame: number | null = null;
let cutAntsLastStepAt = 0;
const workerRowsRevision = ref(0);
const devicePixelRatioRevision = ref(0);
interface GridCellSnapshot {
  style?: XlsxResolvedCellStyle;
  value: string;
}

interface GridMergeRegion {
  end: XlsxCellAddress;
  start: XlsxCellAddress;
}

const workerCellCache = new Map<number, GridCellSnapshot>();
const loadedWorkerBatches = new Set<string>();
const pendingWorkerBatches = new Map<string, AbortController>();
// Cells referenced by `formula` cfvos that the Worker has already been asked for.
const requestedConditionalFormulaCells = new Set<number>();
let workerCacheSheetIndex: number | null = null;
let paintedMergeRegions: GridMergeRegion[] = [];
// Anchor -> owning region and anchor -> top-left cell, both keyed by `cellKey`
// so the per-cell paint path is a map hit instead of a scan over every region.
let paintedMergeAnchors = new Map<number, XlsxCellAddress>();
let paintedMergeRegionsByAnchor = new Map<number, GridMergeRegion>();
// Conditional-format work is data-sized: a rule range can cover the whole used
// extent, so its rule order, per-cell bounds, sampled values and resolved
// environment are cached per sheet revision instead of being rebuilt on every
// paint. `conditionalFormatScope` changes identity whenever the sheet, the
// theme or the revision moves, which drops all of these caches at once.
const conditionalRuleOrder = createScopedCache<XlsxConditionalFormatRule[]>();
const conditionalRuleRanges = createScopedCache<NormalizedRange[]>();
const conditionalRuleEnvironments = createScopedCache<ConditionalEnvironment>();

// ── Computed from controller ──────────────────────────────────────────
const activeSheet = computed(() => props.controller.activeSheet as XlsxSheetData | null);

// ── Grid accessibility ────────────────────────────────────────────────
// The grid paints cells onto canvases, so assistive technology would otherwise
// see one anonymous focusable div. The container is now a real `grid` whose
// `aria-activedescendant` points at a hidden cell carrying the active address
// and value; keyboard navigation already lives in `onGridKeydown`.
const gridA11yId = nextGridA11yId();
const gridAriaLabel = computed(() => {
  const name = activeSheet.value?.name?.trim();
  return name
    ? formatOfficeMessage(m.value.xlsx.grid.worksheetNamed, { name })
    : m.value.xlsx.grid.worksheet;
});
const gridAriaReadOnly = computed(() => (props.readOnly ? "true" : "false"));
// Prefer the worksheet's real size; the display counts include row/column
// filler so the canvas can always paint a full viewport.
const gridRowCount = computed(() => {
  const declared = activeSheet.value?.rowCount ?? 0;
  return declared > 0 ? declared : effectiveRowHeights.value.length;
});
const gridColCount = computed(() => {
  const declared = activeSheet.value?.colCount ?? 0;
  return declared > 0 ? declared : effectiveColWidths.value.length;
});
const activeCellAria = computed(() => {
  const cell = activeCell.value;
  if (!cell) return null;
  const address = `${columnLabel(cell.col)}${cell.row + 1}`;
  const value = props.controller.getCellDisplayValue(cell);
  const display = value && value.trim() ? value : m.value.xlsx.grid.blank;
  return {
    colIndex: cell.col + 1,
    label: `${address}，${display}`,
    rowIndex: cell.row + 1,
    text: `${address}: ${display}`,
  };
});
const activeCellDescendantId = computed(() =>
  activeCellAria.value ? `${gridA11yId}-active-cell` : undefined,
);
// Per-cell metadata is keyed by `cellKey` (one number) rather than a
// `${row}:${col}` string, because the paint loop looks these up for every
// visible cell on every frame.
const hyperlinksByCell = computed(() => new Map((activeSheet.value?.hyperlinks ?? []).map((link) => [cellKey(link.cell.row, link.cell.col), link])));
const commentsByCell = computed(() => new Map((activeSheet.value?.comments ?? []).map((comment) => [cellKey(comment.cell.row, comment.cell.col), comment])));
const sparklinesByCell = computed(() => new Map((activeSheet.value?.sparklines ?? []).map((sparkline) => [cellKey(sparkline.target.row, sparkline.target.col), sparkline])));
const hoveredComment = computed(() => {
  const cell = hoveredCell.value;
  if (!cell) return null;
  const comment = commentsByCell.value.get(cellKey(cell.row, cell.col));
  const display = getDisplayCell(cell);
  if (!comment || !display) return null;
  const anchor = paneOverlayRect(cell, {
    contentX: getColOffsetSum(display.col + 1),
    contentY: getRowOffsetSum(display.row),
    width: 0,
    height: 0,
  });
  if (!anchor) return null;
  return {
    ...comment,
    style: {
      left: `${ROW_HEADER_WIDTH + anchor.left + 6}px`,
      top: `${HEADER_HEIGHT + anchor.top + 4}px`,
    } as CSSProperties,
  };
});
const zoomScale = computed(() => props.controller.zoomScale);
const activeCell = computed(() => props.controller.activeCell);
const selection = computed(() => props.controller.selection);
const revision = computed(() => props.controller.revision);
// The pending cut only gets a marquee on the sheet it was made on; switching
// sheets keeps the move alive but hides its source outline.
const cutRange = computed<XlsxCellRange | null>(() => {
  const cut = props.controller.cutState;
  const sheet = props.controller.activeSheet;
  if (!cut || !sheet || cut.sheetIndex !== sheet.workbookSheetIndex) return null;
  return cut.range;
});
const cutRangeAddress = computed(() => (cutRange.value ? rangeToA1(cutRange.value) : ""));
const editingLocked = computed(() => props.readOnly === true || props.controller.isEditingLocked === true);
const activeValidation = computed<XlsxDataValidation | null>(() => {
  if (editingLocked.value) return null;
  return props.controller.getCellDataValidation?.(activeCell.value ?? null) ?? null;
});
const activeListValidation = computed(() => {
  const validation = activeValidation.value;
  return validation && isDataValidationListDropdownVisible(validation) ? validation : null;
});
const validationDropdownOptions = computed(() => {
  const validation = activeListValidation.value;
  return validation ? props.controller.getDataValidationListOptions?.(validation) ?? [] : [];
});
const validationHint = computed(() => {
  const validation = activeValidation.value;
  const cell = activeCell.value;
  if (!validation || !cell || editingCell.value || validation.showInputMessage !== true || !validation.inputMessage?.trim()) return null;
  const display = getDisplayCell(cell);
  if (!display) return null;
  const anchor = paneOverlayRect(cell, {
    contentY: getRowOffsetSum(display.row) + (effectiveRowHeights.value[display.row] || DEFAULT_ROW_HEIGHT),
    width: 0,
    height: 0,
  });
  if (!anchor) return null;
  return {
    message: validation.inputMessage,
    title: validation.inputTitle,
    style: {
      left: `${ROW_HEADER_WIDTH + anchor.left}px`,
      top: `${HEADER_HEIGHT + anchor.top + 4}px`,
    } as CSSProperties,
  };
});
const validationDropdownTriggerStyle = computed<CSSProperties | null>(() => {
  if (!activeListValidation.value || validationDropdownOptions.value.length === 0 || editingCell.value) return null;
  const cell = activeCell.value;
  const display = cell ? getDisplayCell(cell) : null;
  if (!cell || !display) return null;
  const rowHeight = effectiveRowHeights.value[display.row] || DEFAULT_ROW_HEIGHT;
  const colWidth = effectiveColWidths.value[display.col] || DEFAULT_COL_WIDTH;
  const anchor = paneOverlayRect(cell, {
    contentX: getColOffsetSum(display.col) + colWidth - 20,
    contentY: getRowOffsetSum(display.row) + Math.max(2, (rowHeight - 18) / 2),
    width: 20,
    height: 18,
  });
  if (!anchor) return null;
  return {
    left: `${ROW_HEADER_WIDTH + anchor.left}px`,
    top: `${HEADER_HEIGHT + anchor.top}px`,
  };
});
const validationDropdownStyle = computed<CSSProperties>(() => {
  const cell = activeCell.value;
  const display = cell ? getDisplayCell(cell) : null;
  if (!cell || !display) return {};
  const rowHeight = effectiveRowHeights.value[display.row] || DEFAULT_ROW_HEIGHT;
  const colWidth = effectiveColWidths.value[display.col] || DEFAULT_COL_WIDTH;
  const anchor = paneOverlayRect(cell, {
    contentY: getRowOffsetSum(display.row) + rowHeight,
    width: Math.max(120, colWidth),
    height: 0,
  });
  return {
    left: `${ROW_HEADER_WIDTH + (anchor?.left ?? 0)}px`,
    minWidth: `${Math.max(120, colWidth)}px`,
    top: `${HEADER_HEIGHT + (anchor?.top ?? 0)}px`,
  };
});
const validationFeedbackView = computed(() => {
  const feedback = props.controller.dataValidationFeedback as XlsxDataValidationFeedback | null | undefined;
  if (!feedback) return null;
  const display = getDisplayCell(feedback.cell);
  if (!display) return null;
  const anchor = paneOverlayRect(feedback.cell, {
    contentY: getRowOffsetSum(display.row) + (effectiveRowHeights.value[display.row] || DEFAULT_ROW_HEIGHT),
    width: 0,
    height: 0,
  });
  if (!anchor) return null;
  return {
    ...feedback,
    message: localizeDataValidationFeedback(feedback, m.value.xlsx.validation),
    style: {
      left: `${ROW_HEADER_WIDTH + anchor.left}px`,
      top: `${HEADER_HEIGHT + anchor.top + 4}px`,
    } as CSSProperties,
  };
});

interface ConditionalFormatScope {
  maxUsedCol: number;
  maxUsedRow: number;
  revision: number;
  rowsRevision: number;
  themePalette: XlsxSheetData["themePalette"];
  workbookSheetIndex: number;
}

let conditionalFormatScope: ConditionalFormatScope | null = null;

/** Identity-stable scope token: cached rule data survives repaints but not edits. */
function currentConditionalFormatScope(sheet: XlsxSheetData): ConditionalFormatScope {
  const current = conditionalFormatScope;
  const revisionValue = revision.value;
  const rowsRevision = workerRowsRevision.value;
  if (
    current
    && current.workbookSheetIndex === sheet.workbookSheetIndex
    && current.maxUsedRow === sheet.maxUsedRow
    && current.maxUsedCol === sheet.maxUsedCol
    && current.revision === revisionValue
    && current.rowsRevision === rowsRevision
    && current.themePalette === sheet.themePalette
  ) {
    return current;
  }
  conditionalFormatScope = {
    maxUsedCol: sheet.maxUsedCol,
    maxUsedRow: sheet.maxUsedRow,
    revision: revisionValue,
    rowsRevision,
    themePalette: sheet.themePalette,
    workbookSheetIndex: sheet.workbookSheetIndex,
  };
  return conditionalFormatScope;
}

const effectiveColWidths = computed(() => {
  const sheet = activeSheet.value;
  if (!sheet) return [];
  const z = gridZoomFactor(zoomScale.value);
  const count = Math.max(sheet.colWidths?.length ?? 0, MIN_DISPLAY_COLS);
  return Array.from({ length: count }, (_, index) => {
    const actualIndex = sheet.visibleCols?.[index] ?? index;
    const preview = resizePreview.value;
    const size = preview?.kind === "column" && preview.actualIndex === actualIndex
      ? preview.size
      : sheet.colWidths?.[index] || sheet.colWidthOverridesPx?.[actualIndex] || sheet.defaultColWidthPx || DEFAULT_COL_WIDTH;
    return size * z;
  });
});

const effectiveRowHeights = computed(() => {
  const sheet = activeSheet.value;
  if (!sheet) return [];
  const z = gridZoomFactor(zoomScale.value);
  const count = Math.max(sheet.rowHeights?.length ?? 0, MIN_DISPLAY_ROWS);
  return Array.from({ length: count }, (_, index) => {
    const actualIndex = sheet.visibleRows?.[index] ?? index;
    const preview = resizePreview.value;
    const size = preview?.kind === "row" && preview.actualIndex === actualIndex
      ? preview.size
      : sheet.rowHeights?.[index] || sheet.rowHeightOverridesPx?.[actualIndex] || sheet.defaultRowHeightPx || DEFAULT_ROW_HEIGHT;
    return size * z;
  });
});

function createOffsetIndex(sizes: number[]): number[] {
  const offsets = new Array<number>(sizes.length + 1);
  offsets[0] = 0;
  for (let index = 0; index < sizes.length; index += 1) {
    offsets[index + 1] = offsets[index] + Math.max(0, sizes[index] || 0);
  }
  return offsets;
}

function findDisplayIndexAtOffset(offsets: number[], offset: number): number {
  const itemCount = offsets.length - 1;
  const total = offsets[itemCount] ?? 0;
  if (itemCount <= 0 || offset < 0 || offset >= total) return -1;

  let low = 1;
  let high = itemCount;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if ((offsets[middle] ?? 0) > offset) high = middle;
    else low = middle + 1;
  }
  return low - 1;
}

function resolveVisibleRange(offsets: number[], startOffset: number, endOffset: number) {
  const itemCount = offsets.length - 1;
  if (itemCount <= 0) return { start: 0, end: -1 };

  const total = offsets[itemCount] ?? 0;
  const start = findDisplayIndexAtOffset(offsets, Math.max(0, startOffset));
  const end = findDisplayIndexAtOffset(offsets, Math.max(0, endOffset));
  return {
    start: start >= 0 ? start : itemCount - 1,
    end: end >= 0 && endOffset < total ? end : itemCount - 1,
  };
}

function resolveDisplayIndices(indices: number[] | undefined, count: number): number[] {
  if (!indices?.length) return Array.from({ length: count }, (_, index) => index);
  if (indices.length >= count) return indices.slice(0, count);
  const resolved = [...indices];
  let nextIndex = Math.max(-1, ...resolved) + 1;
  while (resolved.length < count) resolved.push(nextIndex++);
  return resolved;
}

function createSheetToDisplayIndex(indices: number[]): Map<number, number> {
  return new Map(indices.map((sheetIndex, displayIndex) => [sheetIndex, displayIndex]));
}

function resolveAdjacentSheetCell(
  cell: XlsxCellAddress,
  rowDelta: number,
  colDelta: number,
  rows: number[],
  cols: number[],
  rowIndex: Map<number, number>,
  colIndex: Map<number, number>,
): XlsxCellAddress | null {
  const displayRow = rowIndex.get(cell.row);
  const displayCol = colIndex.get(cell.col);
  if (displayRow === undefined || displayCol === undefined || rows.length === 0 || cols.length === 0) return null;
  const nextDisplayRow = Math.max(0, Math.min(displayRow + rowDelta, rows.length - 1));
  const nextDisplayCol = Math.max(0, Math.min(displayCol + colDelta, cols.length - 1));
  return { row: rows[nextDisplayRow], col: cols[nextDisplayCol] };
}

const colOffsets = computed(() => createOffsetIndex(effectiveColWidths.value));
const rowOffsets = computed(() => createOffsetIndex(effectiveRowHeights.value));
const displayCols = computed(() => resolveDisplayIndices(activeSheet.value?.visibleCols, effectiveColWidths.value.length));
const displayRows = computed(() => resolveDisplayIndices(activeSheet.value?.visibleRows, effectiveRowHeights.value.length));
const sheetColToDisplayIndex = computed(() => createSheetToDisplayIndex(displayCols.value));
const sheetRowToDisplayIndex = computed(() => createSheetToDisplayIndex(displayRows.value));

const visibleRowRange = computed(() => {
  const bodyHeight = Math.max(0, containerHeight.value - HEADER_HEIGHT);
  return resolveVisibleRange(
    rowOffsets.value,
    scrollTop.value - OVERSCAN,
    scrollTop.value + bodyHeight + OVERSCAN,
  );
});

const visibleColRange = computed(() => {
  const bodyWidth = Math.max(0, containerWidth.value - ROW_HEADER_WIDTH);
  return resolveVisibleRange(
    colOffsets.value,
    scrollLeft.value - OVERSCAN,
    scrollLeft.value + bodyWidth + OVERSCAN,
  );
});

// ── Helpers ───────────────────────────────────────────────────────────
function getColOffsetSum(col: number): number {
  return colOffsets.value[Math.max(0, Math.min(col, effectiveColWidths.value.length))] ?? 0;
}

function getRowOffsetSum(row: number): number {
  return rowOffsets.value[Math.max(0, Math.min(row, effectiveRowHeights.value.length))] ?? 0;
}

function getSheetCell(displayRow: number, displayCol: number): XlsxCellAddress | null {
  const row = displayRows.value[displayRow];
  const col = displayCols.value[displayCol];
  if (row === undefined || col === undefined) return null;
  return { row, col };
}

function getDisplayCell(cell: XlsxCellAddress): XlsxCellAddress | null {
  const row = sheetRowToDisplayIndex.value.get(cell.row);
  const col = sheetColToDisplayIndex.value.get(cell.col);
  if (row === undefined || col === undefined) return null;
  return { row, col };
}

type XlsxZoomAnchor = {
  row?: number;
  col?: number;
  rowRatio: number;
  colRatio: number;
  clientX: number;
  clientY: number;
  scrollX: boolean;
  scrollY: boolean;
};

interface FrozenPaneMetrics {
  /** Pixel width of the frozen column strip. */
  width: number;
  /** Pixel height of the frozen row strip. */
  height: number;
  /** Number of frozen display columns. */
  cols: number;
  /** Number of frozen display rows. */
  rows: number;
}

function frozenPaneMetrics(): FrozenPaneMetrics {
  const sheet = activeSheet.value;
  const freezeRow = sheet?.freezePanes?.row ?? 0;
  const freezeCol = sheet?.freezePanes?.col ?? 0;
  let width = 0;
  let height = 0;
  let cols = 0;
  let rows = 0;
  for (let index = 0; index < displayCols.value.length; index += 1) {
    if ((displayCols.value[index] ?? 0) >= freezeCol) break;
    width += effectiveColWidths.value[index] || 0;
    cols += 1;
  }
  for (let index = 0; index < displayRows.value.length; index += 1) {
    if ((displayRows.value[index] ?? 0) >= freezeRow) break;
    height += effectiveRowHeights.value[index] || 0;
    rows += 1;
  }
  return { width, height, cols, rows };
}

/**
 * Grid-body-local point (headers excluded) mapped into sheet content space.
 * Frozen strips keep content coordinates; the scrolling pane adds the live
 * scroll offset, so a click on a pinned cell resolves to that cell instead
 * of whatever scrolled underneath it.
 */
function resolvePanePoint(clientX: number, clientY: number) {
  const container = containerRef.value;
  if (!container) return null;
  const rect = container.getBoundingClientRect();
  const localX = clientX - rect.left - ROW_HEADER_WIDTH;
  const localY = clientY - rect.top - HEADER_HEIGHT;
  if (localX < 0 || localY < 0) return null;
  const frozen = frozenPaneMetrics();
  return resolvePaneContentPoint(
    localX,
    localY,
    container.scrollLeft,
    container.scrollTop,
    frozen.width,
    frozen.height,
  );
}

/** Body viewport size the canvases and overlays share (headers excluded). */
function bodyViewportSize(): { width: number; height: number } {
  return {
    width: Math.max(0, containerWidth.value - ROW_HEADER_WIDTH),
    height: Math.max(0, containerHeight.value - HEADER_HEIGHT),
  };
}

interface PageBreakOverlayLine {
  kind: "column" | "row";
  style: CSSProperties;
}

interface PageBreakOverlayLabel {
  style: CSSProperties;
  text: string;
}

const pageBreakOverlay = computed<{
  labels: PageBreakOverlayLabel[];
  lines: PageBreakOverlayLine[];
} | null>(() => {
  const controller = props.controller as XlsxViewerController & {
    isPageBreakPreview?: boolean;
    printLayout?: XlsxPrintLayout | null;
  };
  if (!controller.isPageBreakPreview) return null;
  const layout = controller.printLayout;
  const sheet = activeSheet.value;
  if (!layout || !sheet || layout.pages.length === 0) return null;
  const frozen = frozenPaneMetrics();
  const viewport = bodyViewportSize();
  const lines = new Map<string, PageBreakOverlayLine>();
  const labels: PageBreakOverlayLabel[] = [];

  for (const page of layout.pages) {
    if (page.colPage > 0) {
      const displayCol = sheetColToDisplayIndex.value.get(page.range.start.col);
      if (displayCol !== undefined) {
        const contentX = getColOffsetSum(displayCol);
        const x = paneViewOffset(contentX, scrollLeft.value, displayCol < frozen.cols);
        if (displayCol >= frozen.cols && x < frozen.width) continue;
        const key = `column:${Math.round(x)}`;
        if (!lines.has(key)) {
          lines.set(key, {
            kind: "column",
            style: {
              height: `${viewport.height}px`,
              left: `${ROW_HEADER_WIDTH + x}px`,
              top: `${HEADER_HEIGHT}px`,
            },
          });
        }
      }
    }
    if (page.rowPage > 0) {
      const displayRow = sheetRowToDisplayIndex.value.get(page.range.start.row);
      if (displayRow !== undefined) {
        const contentY = getRowOffsetSum(displayRow);
        const y = paneViewOffset(contentY, scrollTop.value, displayRow < frozen.rows);
        if (displayRow >= frozen.rows && y < frozen.height) continue;
        const key = `row:${Math.round(y)}`;
        if (!lines.has(key)) {
          lines.set(key, {
            kind: "row",
            style: {
              left: `${ROW_HEADER_WIDTH}px`,
              top: `${HEADER_HEIGHT + y}px`,
              width: `${viewport.width}px`,
            },
          });
        }
      }
    }

    const displayRow = sheetRowToDisplayIndex.value.get(page.range.start.row);
    const displayCol = sheetColToDisplayIndex.value.get(page.range.start.col);
    if (displayRow === undefined || displayCol === undefined) continue;
    const x = paneViewOffset(getColOffsetSum(displayCol), scrollLeft.value, displayCol < frozen.cols);
    const y = paneViewOffset(getRowOffsetSum(displayRow), scrollTop.value, displayRow < frozen.rows);
    if (displayCol >= frozen.cols && x < frozen.width) continue;
    if (displayRow >= frozen.rows && y < frozen.height) continue;
    if (x > viewport.width || y > viewport.height || x < -120 || y < -60) continue;
    labels.push({
      style: {
        left: `${ROW_HEADER_WIDTH + Math.max(0, x) + 6}px`,
        top: `${HEADER_HEIGHT + Math.max(0, y) + 4}px`,
      },
      text: String(page.pageNumber),
    });
  }

  return { labels, lines: [...lines.values()] };
});

/**
 * Viewport box for a cell-anchored overlay (editor, comment popover,
 * validation hint/dropdown). The box is mapped into the cell's pane and
 * clamped there, so an overlay never paints over the frozen strips.
 * `contentX`/`contentY` default to the cell's top-left corner.
 */
function paneOverlayRect(
  cell: XlsxCellAddress,
  options: { contentX?: number; contentY?: number; width?: number; height?: number } = {},
): PaneOverlayRect | null {
  const display = getDisplayCell(cell);
  if (!display) return null;
  const frozen = frozenPaneMetrics();
  const viewport = bodyViewportSize();
  return resolvePaneOverlayRect({
    contentX: options.contentX ?? getColOffsetSum(display.col),
    contentY: options.contentY ?? getRowOffsetSum(display.row),
    width: options.width ?? (effectiveColWidths.value[display.col] || DEFAULT_COL_WIDTH),
    height: options.height ?? (effectiveRowHeights.value[display.row] || DEFAULT_ROW_HEIGHT),
    displayCol: display.col,
    displayRow: display.row,
    frozenCols: frozen.cols,
    frozenRows: frozen.rows,
    frozenWidth: frozen.width,
    frozenHeight: frozen.height,
    viewportWidth: viewport.width,
    viewportHeight: viewport.height,
    scrollLeft: scrollLeft.value,
    scrollTop: scrollTop.value,
  });
}

const columnMenuTriggerStyle = computed<CSSProperties | null>(() => {
  const axis = hoveredAxis.value;
  if (!axis || axis.kind !== "column") return null;
  const anchor = paneOverlayRect(
    { col: axis.actualIndex, row: 0 },
    { contentX: getColOffsetSum(axis.displayIndex + 1), contentY: 0, height: 0, width: 0 },
  );
  if (!anchor) return null;
  return {
    left: `${ROW_HEADER_WIDTH + anchor.left - 16}px`,
    top: "3px",
  };
});

const rowMenuTriggerStyle = computed<CSSProperties | null>(() => {
  const axis = hoveredAxis.value;
  if (!axis || axis.kind !== "row") return null;
  const rowHeight = effectiveRowHeights.value[axis.displayIndex] || DEFAULT_ROW_HEIGHT;
  const anchor = paneOverlayRect(
    { col: 0, row: axis.actualIndex },
    { contentX: 0, contentY: getRowOffsetSum(axis.displayIndex) + rowHeight / 2, height: 0, width: 0 },
  );
  if (!anchor) return null;
  return {
    left: "3px",
    top: `${HEADER_HEIGHT + anchor.top - 7}px`,
  };
});

function openHeaderMenu(kind: "column" | "row", event: PointerEvent) {
  const axis = hoveredAxis.value;
  if (!axis || axis.kind !== kind) return;
  const target = event.currentTarget as HTMLElement | null;
  const rect = typeof target?.getBoundingClientRect === "function"
    ? target.getBoundingClientRect()
    : { bottom: 0, left: 0 };
  headerMenu.value = {
    index: axis.actualIndex,
    kind,
    x: rect.left,
    y: rect.bottom + 2,
  };
}

function captureZoomAnchor(clientX: number, clientY: number): XlsxZoomAnchor | null {
  const container = containerRef.value;
  if (!container) return null;
  const rect = container.getBoundingClientRect();
  const localX = clientX - rect.left;
  const localY = clientY - rect.top;
  const bodyX = localX - ROW_HEADER_WIDTH;
  const bodyY = localY - HEADER_HEIGHT;
  const frozen = frozenPaneMetrics();
  const scrollX = localX > ROW_HEADER_WIDTH && bodyX >= frozen.width;
  const scrollY = localY > HEADER_HEIGHT && bodyY >= frozen.height;
  if (!scrollX && !scrollY) return null;

  const anchor: XlsxZoomAnchor = {
    rowRatio: 0,
    colRatio: 0,
    clientX,
    clientY,
    scrollX,
    scrollY,
  };
  if (scrollX) {
    const contentX = Math.max(0, bodyX + container.scrollLeft);
    const displayCol = findDisplayIndexAtOffset(colOffsets.value, contentX);
    const col = displayCols.value[displayCol];
    const width = effectiveColWidths.value[displayCol] || 1;
    if (displayCol < 0 || col === undefined) return null;
    anchor.col = col;
    anchor.colRatio = (contentX - getColOffsetSum(displayCol)) / width;
  }
  if (scrollY) {
    const contentY = Math.max(0, bodyY + container.scrollTop);
    const displayRow = findDisplayIndexAtOffset(rowOffsets.value, contentY);
    const row = displayRows.value[displayRow];
    const height = effectiveRowHeights.value[displayRow] || 1;
    if (displayRow < 0 || row === undefined) return null;
    anchor.row = row;
    anchor.rowRatio = (contentY - getRowOffsetSum(displayRow)) / height;
  }
  return anchor;
}

function restoreZoomAnchor(anchor: XlsxZoomAnchor): void {
  const container = containerRef.value;
  if (!container) return;
  const rect = container.getBoundingClientRect();
  if (anchor.scrollX && anchor.col !== undefined) {
    const displayCol = sheetColToDisplayIndex.value.get(anchor.col);
    if (displayCol !== undefined) {
      const contentX = getColOffsetSum(displayCol)
        + (effectiveColWidths.value[displayCol] || 0) * anchor.colRatio;
      container.scrollLeft = contentX - (anchor.clientX - rect.left - ROW_HEADER_WIDTH);
    }
  }
  if (anchor.scrollY && anchor.row !== undefined) {
    const displayRow = sheetRowToDisplayIndex.value.get(anchor.row);
    if (displayRow !== undefined) {
      const contentY = getRowOffsetSum(displayRow)
        + (effectiveRowHeights.value[displayRow] || 0) * anchor.rowRatio;
      container.scrollTop = contentY - (anchor.clientY - rect.top - HEADER_HEIGHT);
    }
  }
  schedulePaint();
}

function workerCellKey(sheetIndex: number, row: number, col: number) {
  // Numeric, for the same reason as `cellKey`: the paint loop reads one worker
  // snapshot per visible cell per frame.
  return sheetCellKey(sheetIndex, row, col);
}

function resetWorkerRows(sheetIndex: number | null) {
  for (const controller of pendingWorkerBatches.values()) controller.abort();
  pendingWorkerBatches.clear();
  loadedWorkerBatches.clear();
  workerCellCache.clear();
  requestedConditionalFormulaCells.clear();
  workerCacheSheetIndex = sheetIndex;
  workerRowsRevision.value += 1;
}

function getGridCellDisplayValue(cell: XlsxCellAddress) {
  workerRowsRevision.value;
  const sheetIndex = activeSheet.value?.workbookSheetIndex;
  if (sheetIndex !== undefined && props.controller.isWorkerBacked) {
    const cached = workerCellCache.get(workerCellKey(sheetIndex, cell.row, cell.col));
    if (cached !== undefined) {
      // Worker rows keep the raw style, so localized formats render the same
      // way they do on the main thread.
      return localizeFormattedNumber(cached.value, cached.style?.numberFormat, props.controller.locale);
    }
  }
  return props.controller.getCellDisplayValue(cell);
}

function getWorkerCellStyle(cell: XlsxCellAddress): XlsxResolvedCellStyle | undefined {
  workerRowsRevision.value;
  const sheetIndex = activeSheet.value?.workbookSheetIndex;
  if (sheetIndex === undefined || !props.controller.isWorkerBacked) return undefined;
  return workerCellCache.get(workerCellKey(sheetIndex, cell.row, cell.col))?.style;
}

function asWorkerCellValue(value: unknown) {
  if (value === null || value === undefined) return "";
  return typeof value === "string" ? value : String(value);
}

function hydrateVisibleWorkerRows() {
  const sheet = activeSheet.value;
  const loadRows = props.controller.getRowsBatchAsync;
  if (!sheet || !props.controller.isWorkerBacked || !loadRows) return;

  const sheetIndex = sheet.workbookSheetIndex;
  if (workerCacheSheetIndex !== sheetIndex) resetWorkerRows(sheetIndex);
  const visibleRows = displayRows.value.slice(
    visibleRowRange.value.start,
    visibleRowRange.value.end + 1,
  );
  if (visibleRows.length === 0) return;
  const firstRow = Math.min(...visibleRows);
  const lastRow = Math.max(...visibleRows);
  const firstBatch = Math.floor(firstRow / WORKER_ROW_BATCH_SIZE) * WORKER_ROW_BATCH_SIZE;
  const lastBatch = Math.floor(lastRow / WORKER_ROW_BATCH_SIZE) * WORKER_ROW_BATCH_SIZE;

  for (let startRow = firstBatch; startRow <= lastBatch; startRow += WORKER_ROW_BATCH_SIZE) {
    const batchKey = `${sheetIndex}:${startRow}`;
    if (loadedWorkerBatches.has(batchKey) || pendingWorkerBatches.has(batchKey)) continue;
    const controller = new AbortController();
    pendingWorkerBatches.set(batchKey, controller);
    void loadRows(sheetIndex, startRow, WORKER_ROW_BATCH_SIZE, controller.signal)
      .then((rows) => {
        if (controller.signal.aborted || workerCacheSheetIndex !== sheetIndex) return;
        if (Array.isArray(rows)) {
          for (const rowEntry of rows) {
            if (!rowEntry || typeof rowEntry !== "object") continue;
            const row = Reflect.get(rowEntry, "index");
            const cells = Reflect.get(rowEntry, "cells");
            if (!Number.isInteger(row) || !Array.isArray(cells)) continue;
            for (const cellEntry of cells) {
              if (!cellEntry || typeof cellEntry !== "object") continue;
              const col = Reflect.get(cellEntry, "col");
              if (!Number.isInteger(col)) continue;
              const style = Reflect.get(cellEntry, "style");
              workerCellCache.set(
                workerCellKey(sheetIndex, row, col),
                {
                  value: asWorkerCellValue(Reflect.get(cellEntry, "value")),
                  ...(style && typeof style === "object" ? { style: style as XlsxResolvedCellStyle } : {}),
                },
              );
            }
          }
        }
        loadedWorkerBatches.add(batchKey);
        workerRowsRevision.value += 1;
        schedulePaint();
      })
      .finally(() => {
        pendingWorkerBatches.delete(batchKey);
      });
  }
}

function columnLabel(col: number): string {
  let label = "";
  let n = col;
  while (n >= 0) { label = String.fromCharCode(65 + (n % 26)) + label; n = Math.floor(n / 26) - 1; }
  return label;
}

function truncateText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string {
  if (ctx.measureText(text).width <= maxWidth) return text;
  let lo = 0, hi = text.length;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    if (ctx.measureText(text.slice(0, mid) + "\u2026").width <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return text.slice(0, lo) + "\u2026";
}

function isCellInRange(cell: XlsxCellAddress, range: XlsxCellRange | null): boolean {
  if (!range) return false;
  return cell.row >= Math.min(range.start.row, range.end.row) &&
    cell.row <= Math.max(range.start.row, range.end.row) &&
    cell.col >= Math.min(range.start.col, range.end.col) &&
    cell.col <= Math.max(range.start.col, range.end.col);
}

function rangesIntersect(left: XlsxCellRange, right: XlsxCellRange): boolean {
  const leftStartRow = Math.min(left.start.row, left.end.row);
  const leftEndRow = Math.max(left.start.row, left.end.row);
  const leftStartCol = Math.min(left.start.col, left.end.col);
  const leftEndCol = Math.max(left.start.col, left.end.col);
  const rightStartRow = Math.min(right.start.row, right.end.row);
  const rightEndRow = Math.max(right.start.row, right.end.row);
  const rightStartCol = Math.min(right.start.col, right.end.col);
  const rightEndCol = Math.max(right.start.col, right.end.col);
  return leftStartRow <= rightEndRow && leftEndRow >= rightStartRow && leftStartCol <= rightEndCol && leftEndCol >= rightStartCol;
}

function drawRangeBorder(
  ctx: CanvasRenderingContext2D,
  range: XlsxCellRange,
  style: {
    color: string;
    lineWidth?: number;
    /** Dash pattern; omitted for a solid outline. */
    dash?: readonly number[];
    /** Phase of the dash pattern in pixels (marching ants). */
    dashOffset?: number;
    /** Inset from the range edge in CSS pixels. */
    inset?: number;
  },
): void {
  const normalized = {
    start: { row: Math.min(range.start.row, range.end.row), col: Math.min(range.start.col, range.end.col) },
    end: { row: Math.max(range.start.row, range.end.row), col: Math.max(range.start.col, range.end.col) },
  };
  const start = getDisplayCell(normalized.start);
  const end = getDisplayCell(normalized.end);
  if (!start || !end) return;
  const x = getColOffsetSum(start.col);
  const y = getRowOffsetSum(start.row);
  const width = getColOffsetSum(end.col + 1) - x;
  const height = getRowOffsetSum(end.row + 1) - y;
  const inset = style.inset ?? 1;
  ctx.save();
  ctx.strokeStyle = style.color;
  ctx.lineWidth = style.lineWidth ?? 2;
  if (style.dash) {
    if (typeof ctx.setLineDash === "function") ctx.setLineDash([...style.dash]);
    ctx.lineDashOffset = style.dashOffset ?? 0;
  }
  ctx.strokeRect(x + inset, y + inset, Math.max(0, width - inset * 2), Math.max(0, height - inset * 2));
  ctx.restore();
}

function drawSelectionBorder(ctx: CanvasRenderingContext2D, range: XlsxCellRange, color: string): void {
  drawRangeBorder(ctx, range, { color });
}

/**
 * Excel's marquee around a cut source: a white halo plus a dashed outline whose
 * dash offset advances every animation tick.
 */
function drawCutAntsBorder(ctx: CanvasRenderingContext2D, range: XlsxCellRange, color: string): void {
  drawRangeBorder(ctx, range, {
    color: props.isDark ? "#18181b" : "#ffffff",
    inset: 0.5,
    lineWidth: 3,
  });
  drawRangeBorder(ctx, range, {
    color,
    dash: CUT_ANTS_DASH,
    dashOffset: -cutAntsStep.value,
    inset: 0.5,
    lineWidth: 1.5,
  });
}

function normalizeMergeRegion(value: unknown): GridMergeRegion | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  const start = record.start && typeof record.start === "object" ? record.start as Record<string, unknown> : undefined;
  const end = record.end && typeof record.end === "object" ? record.end as Record<string, unknown> : undefined;
  const startRow = Number(start?.row ?? record.startRow);
  const startCol = Number(start?.col ?? record.startCol);
  const endRow = Number(end?.row ?? record.endRow);
  const endCol = Number(end?.col ?? record.endCol);
  if (![startRow, startCol, endRow, endCol].every(Number.isFinite)) return null;
  return {
    start: { row: Math.min(startRow, endRow), col: Math.min(startCol, endCol) },
    end: { row: Math.max(startRow, endRow), col: Math.max(startCol, endCol) },
  };
}

function updatePaintedMergeRegions(rawRegions: unknown): void {
  const regions = Array.isArray(rawRegions)
    ? rawRegions.map(normalizeMergeRegion).filter((region): region is GridMergeRegion => region !== null)
    : [];
  const anchors = new Map<number, XlsxCellAddress>();
  const regionsByAnchor = new Map<number, GridMergeRegion>();
  for (const region of regions) {
    regionsByAnchor.set(cellKey(region.start.row, region.start.col), region);
    for (let row = region.start.row; row <= region.end.row; row += 1) {
      for (let col = region.start.col; col <= region.end.col; col += 1) {
        if (row === region.start.row && col === region.start.col) continue;
        anchors.set(cellKey(row, col), region.start);
      }
    }
  }
  paintedMergeRegions = regions;
  paintedMergeAnchors = anchors;
  paintedMergeRegionsByAnchor = regionsByAnchor;
}

function resolveMergeAnchor(cell: XlsxCellAddress): XlsxCellAddress {
  return paintedMergeAnchors.get(cellKey(cell.row, cell.col)) ?? cell;
}

function mergeRegionAtAnchor(cell: XlsxCellAddress): GridMergeRegion | undefined {
  return paintedMergeRegionsByAnchor.get(cellKey(cell.row, cell.col));
}

function hitTestCell(clientX: number, clientY: number): XlsxCellAddress | null {
  const point = resolvePanePoint(clientX, clientY);
  if (!point) return null;
  const displayCol = findDisplayIndexAtOffset(colOffsets.value, point.contentX);
  const displayRow = findDisplayIndexAtOffset(rowOffsets.value, point.contentY);
  if (displayCol < 0 || displayRow < 0) return null;
  const cell = getSheetCell(displayRow, displayCol);
  return cell ? resolveMergeAnchor(cell) : null;
}

function hitTestSheetPoint(clientX: number, clientY: number): { row: number; col: number; xOffset: number; yOffset: number } | null {
  const point = resolvePanePoint(clientX, clientY);
  if (!point) return null;
  const displayCol = findDisplayIndexAtOffset(colOffsets.value, point.contentX);
  const displayRow = findDisplayIndexAtOffset(rowOffsets.value, point.contentY);
  if (displayCol < 0 || displayRow < 0) return null;
  const cell = getSheetCell(displayRow, displayCol);
  if (!cell) return null;
  const left = getColOffsetSum(displayCol);
  const top = getRowOffsetSum(displayRow);
  const width = Math.max(1, effectiveColWidths.value[displayCol] || DEFAULT_COL_WIDTH);
  const height = Math.max(1, effectiveRowHeights.value[displayRow] || DEFAULT_ROW_HEIGHT);
  return {
    row: cell.row,
    col: cell.col,
    xOffset: Math.max(0, Math.min(1, (point.contentX - left) / width)),
    yOffset: Math.max(0, Math.min(1, (point.contentY - top) / height)),
  };
}

// ── Canvas painting ───────────────────────────────────────────────────
function resizeCanvas(canvas: HTMLCanvasElement | null, width: number, height: number) {
  if (!canvas) return;
  const safeWidth = Math.max(0, width);
  const safeHeight = Math.max(0, height);
  // The browser aborts / blanks a canvas whose backing store exceeds its
  // per-canvas area or dimension limit. Lower the effective ratio instead of
  // letting a wide window at high DPR take the grid down.
  const pixelRatio = resolveCanvasPixelRatio(safeWidth, safeHeight, readDevicePixelRatio());
  const pixelWidth = safeWidth > 0 ? Math.max(1, Math.min(MAX_CANVAS_DIMENSION_PX, Math.ceil(safeWidth * pixelRatio))) : 0;
  const pixelHeight = safeHeight > 0 ? Math.max(1, Math.min(MAX_CANVAS_DIMENSION_PX, Math.ceil(safeHeight * pixelRatio))) : 0;
  const cssWidth = `${safeWidth}px`;
  const cssHeight = `${safeHeight}px`;
  if (canvas.width !== pixelWidth) canvas.width = pixelWidth;
  if (canvas.height !== pixelHeight) canvas.height = pixelHeight;
  if (canvas.style.width !== cssWidth) canvas.style.width = cssWidth;
  if (canvas.style.height !== cssHeight) canvas.style.height = cssHeight;
  const ctx = canvas.getContext("2d");
  if (ctx) ctx.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
}

interface CanvasCellStyle {
  alignment: "left" | "center" | "right";
  backgroundColor: string;
  border?: Record<string, Record<string, unknown>>;
  color: string;
  fontFamily: string;
  fontSizePx: number;
  fontStyle: string;
  fontWeight: string;
  strike: boolean;
  underline: boolean;
  vertical: "top" | "middle" | "bottom";
  wrap: boolean;
}

function asStyleRecord(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === "object" ? value as Record<string, unknown> : undefined;
}

function cssText(value: unknown): string | undefined {
  return typeof value === "string" || typeof value === "number" ? String(value) : undefined;
}

function conditionalRuleBounds(scope: ConditionalFormatScope, rule: XlsxConditionalFormatRule): NormalizedRange[] {
  return conditionalRuleRanges.resolve(scope, rule, () => normalizeConditionalFormatRanges(rule.ranges ?? []));
}

// Scratch result for `conditionalRulesForCell`: the paint loop only reads
// index 0 and never keeps the list, so a single reused array replaces one
// `filter()` allocation per visible cell per frame.
const conditionalCellRules: XlsxConditionalFormatRule[] = [];
const NO_CONDITIONAL_RULES: XlsxConditionalFormatRule[] = [];

function conditionalRulesForCell(cell: XlsxCellAddress): XlsxConditionalFormatRule[] {
  const sheet = activeSheet.value;
  const rules = sheet?.conditionalFormatRules ?? [];
  conditionalCellRules.length = 0;
  if (!sheet || !rules.length) return NO_CONDITIONAL_RULES;
  const scope = currentConditionalFormatScope(sheet);
  // Sorted by rule priority once per revision: index 0 is the rule Excel would
  // apply first, and filtering preserves that order.
  const ordered = conditionalRuleOrder.resolve(scope, rules, () => orderConditionalRules(rules));
  for (const rule of ordered) {
    if (conditionalRuleBounds(scope, rule).some((range) => cellInNormalizedRange(cell, range))) {
      conditionalCellRules.push(rule);
    }
  }
  return conditionalCellRules;
}

function conditionalRangeValues(scope: ConditionalFormatScope, rule: XlsxConditionalFormatRule): number[] {
  // Bounded scan: the values feed the cfvo bounds (including percentile), so
  // they are sampled once per revision instead of once per repaint.
  return collectConditionalFormatValues(conditionalRuleBounds(scope, rule), scope, getGridCellDisplayValue);
}

/**
 * Hydrates a cell referenced by a `formula` cfvo that the Worker has not sent
 * yet. The repaint that follows the snapshot rebuilds the rule environment with
 * the resolved value; a failed request is not retried until the sheet changes.
 */
function requestConditionalFormulaCell(sheetIndex: number, cell: XlsxCellAddress): void {
  const load = props.controller.getCellSnapshotAsync;
  if (!load || !props.controller.isWorkerBacked) return;
  const key = workerCellKey(sheetIndex, cell.row, cell.col);
  if (requestedConditionalFormulaCells.has(key)) return;
  requestedConditionalFormulaCells.add(key);
  void load(sheetIndex, cell.row, cell.col)
    .then((snapshot) => {
      if (!snapshot || workerCacheSheetIndex !== sheetIndex) return;
      workerCellCache.set(key, { value: asWorkerCellValue(snapshot.displayValue) });
      workerRowsRevision.value += 1;
      schedulePaint();
    })
    .catch(() => undefined);
}

/**
 * Sync value for a `formula` cfvo. `null` marks an empty cell (Excel reads it as
 * 0); `undefined` marks a value that is not available yet, and the cfvo then
 * falls back to its automatic bound instead of inventing a number.
 */
function readConditionalCellValue(cell: XlsxCellAddress): number | null | undefined {
  const sheet = activeSheet.value;
  if (!sheet) return undefined;
  if (props.controller.isWorkerBacked) {
    const cached = workerCellCache.get(workerCellKey(sheet.workbookSheetIndex, cell.row, cell.col));
    if (cached === undefined) {
      requestConditionalFormulaCell(sheet.workbookSheetIndex, cell);
      return undefined;
    }
    if (!cached.value.trim()) return null;
    return parseDisplayedNumber(localizeFormattedNumber(cached.value, cached.style?.numberFormat, props.controller.locale), props.controller.locale);
  }
  const text = props.controller.getCellDisplayValue(cell);
  if (!text.trim()) return null;
  return parseDisplayedNumber(text, props.controller.locale);
}

function conditionalEnvironmentFor(rule: XlsxConditionalFormatRule): ConditionalEnvironment {
  const sheet = activeSheet.value;
  if (!sheet) return createConditionalEnvironment(rule, [], () => undefined);
  const scope = currentConditionalFormatScope(sheet);
  return conditionalRuleEnvironments.resolve(scope, rule, () => createConditionalEnvironment(
    rule,
    conditionalRangeValues(scope, rule),
    (color) => resolveWorkbookColor(color, sheet.themePalette) ?? undefined,
    readConditionalCellValue,
  ));
}

function conditionalVisual(cell: XlsxCellAddress, valueText: string) {
  const rule = conditionalRulesForCell(cell)[0];
  if (!rule) return null;
  // Excel leaves blank cells unformatted: no bar, no icon, no scale colour.
  if (!valueText.trim()) return null;
  const value = parseDisplayedNumber(valueText, props.controller.locale);
  if (value === null) return null;
  const environment = conditionalEnvironmentFor(rule);
  if (rule.kind === "dataBar") {
    return { environment, kind: rule.kind, showValue: rule.showValue !== false, value } as const;
  }
  if (rule.kind === "colorScale") {
    return { environment, kind: rule.kind, showValue: true, value } as const;
  }
  return { environment, kind: rule.kind, showValue: rule.showValue !== false, value } as const;
}

function drawConditionalDataBar(
  ctx: CanvasRenderingContext2D,
  conditional: NonNullable<ReturnType<typeof conditionalVisual>>,
  colX: number,
  rowY: number,
  colW: number,
  rowHeight: number,
): void {
  const style = conditional.environment.dataBar;
  if (!style) return;
  const width = Math.max(0, colW - 6);
  const geometry = resolveDataBarGeometry(conditional.environment, conditional.value, width);
  const left = colX + 3;
  const top = rowY + 4;
  const height = Math.max(2, rowHeight - 8);
  if (geometry.barWidth > 0) {
    const fillColor = geometry.negative ? style.negativeFillColor : style.color;
    if (style.gradient) {
      const gradient = ctx.createLinearGradient(
        left + geometry.barX,
        top,
        left + geometry.barX + geometry.barWidth,
        top,
      );
      gradient.addColorStop(0, withAlpha(fillColor, "66"));
      gradient.addColorStop(1, withAlpha(fillColor, "dd"));
      ctx.fillStyle = gradient;
    } else {
      ctx.fillStyle = withAlpha(fillColor, "99");
    }
    ctx.fillRect(left + geometry.barX, top, geometry.barWidth, height);
    if (style.border) {
      ctx.strokeStyle = geometry.negative ? style.negativeBorderColor : style.borderColor;
      ctx.lineWidth = 1;
      ctx.strokeRect(left + geometry.barX + 0.5, top + 0.5, Math.max(0, geometry.barWidth - 1), Math.max(0, height - 1));
    }
  }
  if (geometry.showAxis) {
    ctx.fillStyle = style.axisColor;
    ctx.fillRect(left + geometry.axisX - 0.5, top, 1, height);
  }
}

function drawSparkline(ctx: CanvasRenderingContext2D, sparkline: XlsxSparkline, x: number, y: number, width: number, height: number): void {
  const values: number[] = [];
  for (let row = sparkline.range.start.row; row <= sparkline.range.end.row; row++) {
    for (let col = sparkline.range.start.col; col <= sparkline.range.end.col; col++) {
      const value = parseDisplayedNumber(getGridCellDisplayValue({ row, col }), props.controller.locale);
      if (value !== null) values.push(value);
    }
  }
  if (!values.length) return;
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 0);
  const span = Math.max(1e-9, max - min);
  const px = (index: number) => x + 4 + index * Math.max(1, (width - 8) / Math.max(1, values.length - 1));
  const py = (value: number) => y + height - 4 - ((value - min) / span) * Math.max(1, height - 8);
  ctx.save();
  ctx.strokeStyle = sparkline.color || "#2563eb";
  ctx.fillStyle = sparkline.color || "#2563eb";
  ctx.lineWidth = 1.5;
  if (sparkline.type === "column" || sparkline.type === "winLoss") {
    const barWidth = Math.max(2, (width - 8) / values.length - 2);
    values.forEach((value, index) => ctx.fillRect(px(index) - barWidth / 2, Math.min(py(value), py(0)), barWidth, Math.max(1, Math.abs(py(value) - py(0)))));
  } else {
    ctx.beginPath();
    values.forEach((value, index) => index ? ctx.lineTo(px(index), py(value)) : ctx.moveTo(px(index), py(value)));
    ctx.stroke();
    if (sparkline.markers) values.forEach((value, index) => { ctx.beginPath(); ctx.arc(px(index), py(value), 2, 0, Math.PI * 2); ctx.fill(); });
  }
  ctx.restore();
}

/**
 * Fills a caller-owned style object. The paint loop styles every visible cell on
 * every frame, so returning fresh objects (the previous `base` / `resolvedStyle`
 * / `context` trio) was the largest remaining per-cell allocation; the default
 * path, with no `getCellStyle` prop, now allocates nothing.
 */
function resolveCanvasCellStyleInto(
  target: CanvasCellStyle,
  cell: XlsxCellAddress,
  value: string,
  workbookStyle: XlsxResolvedCellStyle | undefined,
  tableHeaderCells: ReadonlySet<number>,
): CanvasCellStyle {
  const sheet = activeSheet.value;
  const font = asStyleRecord(workbookStyle?.font);
  const fill = asStyleRecord(workbookStyle?.fill);
  const alignment = asStyleRecord(workbookStyle?.alignment);
  const fillStyle = resolveWorkbookFillStyle(fill, sheet?.themePalette);
  const fontColor = resolveWorkbookColor(asStyleRecord(font?.color), sheet?.themePalette);
  target.alignment = alignment?.horizontal === "center" || alignment?.horizontal === "right"
    ? alignment.horizontal
    : "left";
  target.backgroundColor = fillStyle.backgroundColor || (props.isDark ? "#18181b" : "#ffffff");
  target.border = asStyleRecord(workbookStyle?.border) as Record<string, Record<string, unknown>> | undefined;
  target.color = fontColor || (props.isDark ? "#e4e4e7" : "#18181b");
  target.fontFamily = typeof font?.name === "string" && font.name.trim() ? font.name : "Segoe UI";
  target.fontSizePx = Math.max(8, (typeof font?.size === "number" ? font.size * 4 / 3 : 12) * gridZoomFactor(zoomScale.value));
  target.fontStyle = font?.italic ? "italic" : "normal";
  target.fontWeight = font?.bold ? "700" : "400";
  target.strike = Boolean(font?.strikethrough);
  target.underline = Boolean(font?.underline && font.underline !== "none");
  target.vertical = alignment?.vertical === "top" || alignment?.vertical === "center"
    ? alignment.vertical === "center" ? "middle" : "top"
    : "bottom";
  target.wrap = Boolean(alignment?.wrapText);
  const getCellStyle = props.getCellStyle;
  if (!getCellStyle) return target;
  // Host styles need the full cell context, so this branch keeps allocating.
  const context: XlsxCellStyleContext = {
    cell,
    hasChartHighlight: false,
    hasConditionalFormat: conditionalRulesForCell(cell).length > 0,
    hasHyperlink: hyperlinksByCell.value.has(cellKey(cell.row, cell.col)),
    hasValidation: false,
    isMerged: paintedMergeRegionsByAnchor.has(cellKey(cell.row, cell.col)),
    isTableHeader: tableHeaderCells.has(cellKey(cell.row, cell.col)),
    resolvedStyle: {
      backgroundColor: target.backgroundColor,
      color: target.color,
      fontFamily: target.fontFamily,
      fontSize: target.fontSizePx,
      fontStyle: target.fontStyle,
      fontWeight: target.fontWeight,
      textAlign: target.alignment,
      whiteSpace: target.wrap ? "pre-wrap" : "nowrap",
    },
    sheetName: sheet?.name ?? "",
    value,
    workbookSheetIndex: sheet?.workbookSheetIndex ?? 0,
  };
  const override = getCellStyle(cell, context) ?? undefined;
  if (!override) return target;
  const background = cssText(override.backgroundColor ?? override.background);
  const color = cssText(override.color);
  const fontFamily = cssText(override.fontFamily);
  const fontSize = Number.parseFloat(cssText(override.fontSize) ?? "");
  const textAlign = cssText(override.textAlign);
  const verticalAlign = cssText(override.verticalAlign);
  if (background) target.backgroundColor = background;
  if (color) target.color = color;
  if (fontFamily) target.fontFamily = fontFamily;
  if (Number.isFinite(fontSize)) target.fontSizePx = fontSize;
  if (override.fontStyle) target.fontStyle = String(override.fontStyle);
  if (override.fontWeight) target.fontWeight = String(override.fontWeight);
  if (textAlign === "center" || textAlign === "right" || textAlign === "left") target.alignment = textAlign;
  if (verticalAlign === "top" || verticalAlign === "middle" || verticalAlign === "bottom") target.vertical = verticalAlign;
  if (override.whiteSpace === "pre-wrap" || override.whiteSpace === "normal") target.wrap = true;
  return target;
}

/** Style object reused for every cell painted in one frame. */
function createCanvasCellStyleScratch(): CanvasCellStyle {
  return {
    alignment: "left",
    backgroundColor: "#ffffff",
    color: "#18181b",
    fontFamily: "Segoe UI",
    fontSizePx: 12,
    fontStyle: "normal",
    fontWeight: "400",
    strike: false,
    underline: false,
    vertical: "bottom",
    wrap: false,
  };
}

/** Table header cells of the active sheet, keyed by `cellKey`. */
function collectTableHeaderCells(): ReadonlySet<number> {
  const tables = props.controller.tables ?? [];
  if (!tables.length) return EMPTY_CELL_KEYS;
  const headers = new Set<number>();
  for (const table of tables) {
    for (let col = table.start.col; col <= table.end.col; col += 1) {
      headers.add(cellKey(table.start.row, col));
    }
  }
  return headers;
}

const EMPTY_CELL_KEYS: ReadonlySet<number> = new Set<number>();

function borderWidth(style: unknown): number {
  if (style === "medium" || style === "mediumDashed" || style === "mediumDashDot" || style === "mediumDashDotDot") return 2;
  if (style === "thick" || style === "double") return 3;
  return 1;
}

function drawCellBorders(
  ctx: CanvasRenderingContext2D,
  style: CanvasCellStyle,
  x: number,
  y: number,
  width: number,
  height: number,
): void {
  if (!style.border) return;
  const edges = [
    ["top", x, y, x + width, y],
    ["right", x + width, y, x + width, y + height],
    ["bottom", x, y + height, x + width, y + height],
    ["left", x, y, x, y + height],
  ] as const;
  for (const [name, x1, y1, x2, y2] of edges) {
    const edge = style.border[name];
    if (!edge?.style || edge.style === "none") continue;
    ctx.save();
    ctx.strokeStyle = resolveWorkbookColor(asStyleRecord(edge.color), activeSheet.value?.themePalette) || (props.isDark ? "#71717a" : "#52525b");
    ctx.lineWidth = borderWidth(edge.style);
    ctx.setLineDash(edge.style === "dashed" || String(edge.style).includes("Dash") ? [5, 3] : edge.style === "dotted" ? [1, 2] : []);
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    ctx.restore();
  }
}

function wrapCellText(ctx: CanvasRenderingContext2D, value: string, width: number, maxLines: number): string[] {
  // Alt+Enter stores a literal line break in the cell. Excel only shows the
  // first line while wrap text is off, and breaks at every hard newline once
  // it is on; soft wrapping then applies inside each hard line.
  const hardLines = value.replace(/\r\n?/g, "\n").split("\n");
  if (maxLines <= 1) return [truncateText(ctx, hardLines[0] ?? "", width)];
  const lines: string[] = [];
  for (const hardLine of hardLines) {
    if (lines.length >= maxLines) break;
    let line = "";
    for (const character of hardLine) {
      const next = line + character;
      if (line && ctx.measureText(next).width > width) {
        lines.push(line);
        line = character;
        if (lines.length >= maxLines - 1) break;
      } else {
        line = next;
      }
    }
    if (lines.length >= maxLines) break;
    if (line || hardLines.length > 1) lines.push(line);
  }
  return lines;
}

interface CellPaintState {
  style: CanvasCellStyle;
  conditional: ReturnType<typeof conditionalVisual>;
  comment?: XlsxCellComment;
  hyperlink?: XlsxCellHyperlink;
  sparkline?: XlsxSparkline;
  value: string;
  /**
   * Stroke the cell's box with the grid colour. The scrolling pass outlines
   * merged cells only (its gridlines cover the rest); the frozen strips have
   * no scrolled gridlines underneath, so they outline every cell.
   */
  gridLines: boolean;
  selected: boolean;
  selectionFill: string;
  gridColor: string;
  textColor: string;
  /** Width available for text when Excel-style overflow into empty neighbours is allowed. */
  overflowWidth?: number;
  /** Defer text/icon drawing so the caller can repaint it after all cell backgrounds. */
  skipText?: boolean;
}

/**
 * Draws a cell's text and conditional icon. Shared by `drawCellContent` and by
 * the overflow pass, which repaints text after every neighbour background has
 * been filled (otherwise an empty right-hand cell would erase the spill).
 */
function drawCellText(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  state: CellPaintState,
  hyperlinkColor: string,
): void {
  const { style, conditional, hyperlink, sparkline, value } = state;
  const conditionalIcon = conditional?.kind === "iconSet"
    ? resolveIconGeometry(conditional.environment, conditional.value, width, height)
    : null;
  const textAreaWidth = conditionalIcon?.textWidth ?? state.overflowWidth ?? width;
  // A data-bar or icon-set rule may be configured to hide the cell value while
  // still drawing its shape. Respect `showValue` for both conditional kinds.
  const conditionalHidesValue = (conditional?.kind === "dataBar" || conditional?.kind === "iconSet")
    && conditional.showValue === false;
  if (value && !sparkline && !conditionalHidesValue) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(x + 2, y + 1, Math.max(0, textAreaWidth - 4), Math.max(0, height - 2));
    ctx.clip();
    ctx.fillStyle = hyperlink ? hyperlinkColor : style.color || state.textColor;
    ctx.font = `${style.fontStyle} ${style.fontWeight} ${Math.round(style.fontSizePx)}px ${JSON.stringify(style.fontFamily)}, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif`;
    ctx.textAlign = style.alignment;
    const padding = 5;
    const textX = style.alignment === "right"
      ? x + textAreaWidth - padding
      : style.alignment === "center" ? x + textAreaWidth / 2 : x + padding;
    const lineHeight = Math.max(12, style.fontSizePx * 1.25);
    const maxLines = style.wrap ? Math.max(1, Math.floor((height - 4) / lineHeight)) : 1;
    const lines = wrapCellText(ctx, value, Math.max(0, textAreaWidth - padding * 2), maxLines);
    const blockHeight = lines.length * lineHeight;
    const firstBaseline = style.vertical === "top"
      ? y + 2 + lineHeight * 0.8
      : style.vertical === "middle"
        ? y + (height - blockHeight) / 2 + lineHeight * 0.8
        : y + height - blockHeight + lineHeight * 0.8 - 2;
    lines.forEach((line, index) => ctx.fillText(line, textX, firstBaseline + index * lineHeight));
    if (style.underline || style.strike || hyperlink) {
      const measured = Math.min(ctx.measureText(lines[0] ?? "").width, Math.max(0, textAreaWidth - padding * 2));
      const startX = style.alignment === "right" ? textX - measured : style.alignment === "center" ? textX - measured / 2 : textX;
      ctx.strokeStyle = style.color;
      ctx.lineWidth = 1;
      for (const offset of [style.underline || hyperlink ? 2 : null, style.strike ? -style.fontSizePx * 0.3 : null]) {
        if (offset === null) continue;
        ctx.beginPath();
        ctx.moveTo(startX, firstBaseline + offset);
        ctx.lineTo(startX + measured, firstBaseline + offset);
        ctx.stroke();
      }
    }
    ctx.restore();
  }
  if (conditionalIcon) {
    ctx.fillStyle = conditionalIcon.color;
    ctx.font = `${conditionalIcon.size}px "Segoe UI Symbol", "Apple Symbols", sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(
      conditionalIcon.glyph,
      x + conditionalIcon.iconX + conditionalIcon.size / 2,
      y + height / 2,
    );
  }
}

/**
 * Draws a cell's complete content — background, conditional formatting,
 * sparkline, selection tint, text, comment marker and borders — into a
 * viewport rectangle. The scrolling pass and the frozen-pane overdraw share
 * it so a pinned cell keeps the formatting it has in the live pane.
 */
function drawCellContent(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  state: CellPaintState,
): void {
  const { style, conditional, comment, sparkline } = state;
  const conditionalBackground = conditional?.kind === "colorScale"
    ? resolveColorScaleFill(conditional.environment, conditional.value)
    : undefined;
  ctx.fillStyle = conditionalBackground || style.backgroundColor;
  ctx.fillRect(x + 0.5, y + 0.5, Math.max(0, width - 1), Math.max(0, height - 1));
  if (conditional?.kind === "dataBar") {
    drawConditionalDataBar(ctx, conditional, x, y, width, height);
  }
  if (sparkline) drawSparkline(ctx, sparkline, x, y, width, height);
  if (state.gridLines) {
    ctx.strokeStyle = state.gridColor;
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, Math.max(0, width - 1), Math.max(0, height - 1));
  }

  if (state.selected) {
    ctx.fillStyle = state.selectionFill;
    ctx.fillRect(x + 1, y + 1, Math.max(0, width - 2), Math.max(0, height - 2));
  }

  if (!state.skipText) drawCellText(ctx, x, y, width, height, state, "#2563eb");
  if (comment) {
    ctx.fillStyle = "#dc2626";
    ctx.beginPath(); ctx.moveTo(x + width - 9, y + 1); ctx.lineTo(x + width - 1, y + 1); ctx.lineTo(x + width - 1, y + 9); ctx.closePath(); ctx.fill();
  }

  drawCellBorders(ctx, style, x, y, width, height);
}

type ActiveWorksheet = NonNullable<ReturnType<XlsxViewerController["getActiveWorksheet"]>>;

/** Workbook style for a cell: worker cache, then the loaded worksheet, then row/column styles. */
function resolveCellWorkbookStyle(
  cell: XlsxCellAddress,
  worksheet: ActiveWorksheet | null,
  sheet: XlsxSheetData,
): XlsxResolvedCellStyle | undefined {
  let workbookStyle = getWorkerCellStyle(cell);
  if (!workbookStyle && worksheet) {
    try {
      workbookStyle = worksheet.getCellStyleAt(cell.row, cell.col) as XlsxResolvedCellStyle | undefined;
    } catch {
      workbookStyle = undefined;
    }
  }
  if (!workbookStyle) {
    const rowStyleId = sheet.rowStyleIds[cell.row];
    const colStyleId = sheet.colStyleIds[cell.col];
    workbookStyle = rowStyleId !== undefined
      ? sheet.styleById[rowStyleId]
      : colStyleId !== undefined ? sheet.styleById[colStyleId] : undefined;
  }
  return workbookStyle;
}

/**
 * Selection, active-cell and pending-cut outlines, drawn once per pane: each
 * pane clips to its own box and translates content by the scroll offset only on
 * its scrolling axes, so an outline inside the frozen strips lands on the
 * pinned cells instead of their scrolled positions.
 */
function paintSelectionBorders(
  ctx: CanvasRenderingContext2D,
  bodyWidth: number,
  bodyHeight: number,
  color: string,
): void {
  const sel = selection.value;
  const active = activeCell.value;
  const cut = cutRange.value;
  const ranges: XlsxCellRange[] = [];
  if (sel) ranges.push(sel);
  if (active && (!sel || (sel.start.row === sel.end.row && sel.start.col === sel.end.col))) {
    ranges.push({ start: active, end: active });
  }
  if (!ranges.length && !cut) return;

  const frozen = frozenPaneMetrics();
  const panes = [
    { x: 0, y: 0, width: frozen.width, height: frozen.height, scrollX: false, scrollY: false },
    { x: 0, y: frozen.height, width: frozen.width, height: Math.max(0, bodyHeight - frozen.height), scrollX: false, scrollY: true },
    { x: frozen.width, y: 0, width: Math.max(0, bodyWidth - frozen.width), height: frozen.height, scrollX: true, scrollY: false },
    { x: frozen.width, y: frozen.height, width: Math.max(0, bodyWidth - frozen.width), height: Math.max(0, bodyHeight - frozen.height), scrollX: true, scrollY: true },
  ];
  for (const pane of panes) {
    if (pane.width <= 0 || pane.height <= 0) continue;
    ctx.save();
    ctx.beginPath();
    ctx.rect(pane.x, pane.y, pane.width, pane.height);
    ctx.clip();
    ctx.translate(pane.scrollX ? -scrollLeft.value : 0, pane.scrollY ? -scrollTop.value : 0);
    for (const range of ranges) drawSelectionBorder(ctx, range, color);
    if (cut) drawCutAntsBorder(ctx, cut, color);
    ctx.restore();
  }
}

// One worksheet handle per paint, shared by the body and its frozen strips and
// released once the whole frame is drawn. The public accessor creates a fresh
// WASM handle on every call (unlike the controller's internal cached one), so
// allocating it per section would churn handles at frame rate.
function paintGrid() {
  const worksheet: ActiveWorksheet | null = props.controller.isWorkerBacked ? null : props.controller.getActiveWorksheet();
  try {
    paintBody(worksheet);
    paintColHeaders();
    paintRowHeaders();
    paintCorner();
  } finally {
    worksheet?.free();
  }
}

interface OverflowTextTask {
  height: number;
  state: CellPaintState;
  width: number;
  x: number;
  y: number;
}

/** Reusable per-frame paint state; deferred overflow tasks snapshot it. */
function createCellPaintStateScratch(): CellPaintState {
  return {
    style: createCanvasCellStyleScratch(),
    conditional: null,
    comment: undefined,
    hyperlink: undefined,
    sparkline: undefined,
    value: "",
    gridLines: false,
    selected: false,
    selectionFill: "",
    gridColor: "",
    textColor: "",
    overflowWidth: undefined,
    skipText: false,
  };
}

function snapshotConditionalVisual(
  conditional: ReturnType<typeof conditionalVisual>,
): ReturnType<typeof conditionalVisual> {
  return conditional ? { ...conditional } as ReturnType<typeof conditionalVisual> : null;
}

function paintBody(worksheet: ActiveWorksheet | null) {
  const canvas = bodyCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const bw = Math.max(0, containerWidth.value - ROW_HEADER_WIDTH);
  const bh = Math.max(0, containerHeight.value - HEADER_HEIGHT);
  resizeCanvas(canvas, bw, bh);

  ctx.clearRect(0, 0, bw, bh);
  const rowRange = visibleRowRange.value;
  const colRange = visibleColRange.value;
  const sheet = activeSheet.value;
  if (!sheet) return;

  const widths = effectiveColWidths.value;
  const heights = effectiveRowHeights.value;
  updatePaintedMergeRegions(props.controller.isWorkerBacked ? sheet.mergedRegions : worksheet?.mergedRegions);

  const bgColor = props.isDark ? "#18181b" : "#ffffff";
  const gridColor = props.isDark ? "#3f3f46" : "#d9d9d9";
  const textColor = props.isDark ? "#e4e4e7" : "#18181b";
  const selFill = props.selectionFillColor || (props.isDark ? "rgba(37, 99, 235, 0.15)" : "rgba(37, 99, 235, 0.08)");
  const selBorder = props.selectionColor || "#2563eb";
  const viewport = { bottom: scrollTop.value + bh, left: scrollLeft.value, right: scrollLeft.value + bw, top: scrollTop.value };
  const paintedBodyCells = new Set<number>();
  const overflowTextTasks: OverflowTextTask[] = [];

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, bw, bh);
  const hasVisibleRange = rowRange.end >= rowRange.start && colRange.end >= colRange.start;
  const frozen = frozenPaneMetrics();
  const hasFrozenPanes = frozen.width > 0 || frozen.height > 0;

  ctx.save();
  if (hasFrozenPanes) {
    // The frozen strips are painted last by paintFrozenPanes, at pinned
    // coordinates. Excluding them from the scrolling pass keeps every pixel
    // painted once: a translucent selection tint or a soft fill would
    // otherwise double up wherever the two passes overlap.
    ctx.beginPath();
    ctx.rect(0, frozen.height, bw, Math.max(0, bh - frozen.height));
    ctx.rect(frozen.width, 0, Math.max(0, bw - frozen.width), frozen.height);
    ctx.clip();
  }
  ctx.translate(-scrollLeft.value, -scrollTop.value);

  // Gridlines
  if (hasVisibleRange && sheet.showGridLines !== false) {
    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 1;
    const gridTop = getRowOffsetSum(rowRange.start);
    const gridBottom = getRowOffsetSum(rowRange.end + 1);
    const gridLeft = getColOffsetSum(colRange.start);
    const gridRight = getColOffsetSum(colRange.end + 1);
    for (let c = colRange.start; c <= colRange.end + 1 && c <= widths.length; c += 1) {
      const x = getColOffsetSum(c);
      ctx.beginPath(); ctx.moveTo(x, gridTop); ctx.lineTo(x, gridBottom); ctx.stroke();
    }
    for (let r = rowRange.start; r <= rowRange.end + 1 && r <= heights.length; r += 1) {
      const y = getRowOffsetSum(r);
      ctx.beginPath(); ctx.moveTo(gridLeft, y); ctx.lineTo(gridRight, y); ctx.stroke();
    }
  }

  const sel = selection.value;
  const tableHeaderCells = props.getCellStyle ? collectTableHeaderCells() : EMPTY_CELL_KEYS;
  // One state object and one style object for the whole frame. Rebuilding a
  // 14-field `CellPaintState` plus the style trio per visible cell was the bulk
  // of the allocation churn while scrolling; deferred overflow tasks snapshot
  // the two objects they keep a reference to.
  const paintState = createCellPaintStateScratch();
  const styleScratch = createCanvasCellStyleScratch();
  const overflowNeighbourWidths: number[] = [];

  // Excel lets left-aligned text spill into adjacent empty cells. Only the
  // scrolling pass allows it: frozen strips keep their pinned-cell clipping.
  const resolveBodyOverflowWidth = (
    displayRow: number,
    displayCol: number,
    value: string,
    merge: GridMergeRegion | undefined,
    style: CanvasCellStyle,
    conditional: ReturnType<typeof conditionalVisual>,
    sparkline: XlsxSparkline | undefined,
  ): number => {
    const cellWidth = widths[displayCol] || 0;
    if (merge || style.wrap || style.alignment !== "left" || conditional?.kind === "iconSet") return cellWidth;
    if (!value || sparkline || (conditional?.kind === "dataBar" && !conditional.showValue)) return cellWidth;
    overflowNeighbourWidths.length = 0;
    for (let nextCol = displayCol + 1; nextCol < widths.length; nextCol += 1) {
      const neighbour = getSheetCell(displayRow, nextCol);
      if (!neighbour) break;
      const neighbourKey = cellKey(neighbour.row, neighbour.col);
      if (paintedMergeAnchors.has(neighbourKey) || paintedMergeRegionsByAnchor.has(neighbourKey)) break;
      if (commentsByCell.value.has(neighbourKey) || sparklinesByCell.value.has(neighbourKey)) break;
      if (getGridCellDisplayValue(neighbour)) break;
      overflowNeighbourWidths.push(widths[nextCol] || 0);
    }
    return resolveOverflowTextWidth(cellWidth, overflowNeighbourWidths);
  };

  const paintBodyCell = (displayRow: number, displayCol: number, knownCell?: XlsxCellAddress | null): void => {
    const paintKey = cellKey(displayRow, displayCol);
    if (paintedBodyCells.has(paintKey)) return;
    const cell = knownCell === undefined ? getSheetCell(displayRow, displayCol) : knownCell;
    if (!cell || paintedMergeAnchors.has(cellKey(cell.row, cell.col))) return;
    paintedBodyCells.add(paintKey);

    const merge = mergeRegionAtAnchor(cell);
    const mergeEnd = merge ? getDisplayCell(merge.end) : null;
    const colW = mergeEnd
      ? getColOffsetSum(mergeEnd.col + 1) - getColOffsetSum(displayCol)
      : widths[displayCol] || 0;
    const rowHResolved = mergeEnd
      ? getRowOffsetSum(mergeEnd.row + 1) - getRowOffsetSum(displayRow)
      : heights[displayRow] || 0;
    const colX = getColOffsetSum(displayCol);
    const rowY = getRowOffsetSum(displayRow);
    const value = getGridCellDisplayValue(cell);
    const style = resolveCanvasCellStyleInto(
      styleScratch,
      cell,
      value,
      resolveCellWorkbookStyle(cell, worksheet, sheet),
      tableHeaderCells,
    );
    const conditional = conditionalVisual(cell, value);
    const cellKeyValue = cellKey(cell.row, cell.col);
    const sparkline = sparklinesByCell.value.get(cellKeyValue);
    const state = paintState;
    state.style = style;
    state.conditional = conditional;
    state.comment = commentsByCell.value.get(cellKeyValue);
    state.hyperlink = hyperlinksByCell.value.get(cellKeyValue);
    state.sparkline = sparkline;
    state.value = value;
    state.gridLines = Boolean(merge) && sheet.showGridLines !== false;
    state.selected = isCellInRange(cell, sel) || Boolean(merge && sel && rangesIntersect(merge, sel));
    state.selectionFill = selFill;
    state.gridColor = gridColor;
    state.textColor = textColor;
    state.overflowWidth = undefined;
    state.skipText = false;
    const overflowWidth = resolveBodyOverflowWidth(displayRow, displayCol, value, merge, style, conditional, sparkline);
    if (overflowWidth > colW) {
      // Paint the background now and defer the text until every neighbour has
      // been filled; otherwise the empty neighbour background erases the spill.
      state.overflowWidth = overflowWidth;
      state.skipText = true;
      drawCellContent(ctx, colX, rowY, colW, rowHResolved, state);
      // The task outlives this cell, so it owns a snapshot of the two objects
      // the shared state points at.
      overflowTextTasks.push({
        height: rowHResolved,
        state: { ...state, conditional: snapshotConditionalVisual(conditional), style: { ...style } },
        width: colW,
        x: colX,
        y: rowY,
      });
    } else {
      drawCellContent(ctx, colX, rowY, colW, rowHResolved, state);
    }
  };

  try {
    if (hasVisibleRange) {
      for (let r = rowRange.start; r <= rowRange.end && r < heights.length; r++) {
        for (let c = colRange.start; c <= colRange.end && c < widths.length; c++) {
          paintBodyCell(r, c, getSheetCell(r, c));
        }
      }
    }

    // A merged region is painted from its anchor rect even when that anchor is
    // outside the visible range: the canvas clips the region to the viewport,
    // so the still-visible part keeps its fill, borders and selection.
    for (const region of paintedMergeRegions) {
      const start = getDisplayCell(region.start);
      const end = getDisplayCell(region.end);
      if (!start || !end) continue;
      const regionX = getColOffsetSum(start.col);
      const regionY = getRowOffsetSum(start.row);
      const regionWidth = getColOffsetSum(end.col + 1) - regionX;
      const regionHeight = getRowOffsetSum(end.row + 1) - regionY;
      if (!intersectsViewport(
        { height: regionHeight, left: regionX, top: regionY, width: regionWidth },
        viewport,
      )) continue;
      paintBodyCell(start.row, start.col);
    }

    for (const task of overflowTextTasks) {
      drawCellText(ctx, task.x, task.y, task.width, task.height, task.state, "#2563eb");
    }
  } finally {
    ctx.restore();
  }
  paintFrozenPanes(ctx, bw, bh, gridColor, selFill, textColor, worksheet);
  paintSelectionBorders(ctx, bw, bh, selBorder);
}

function paintFrozenPanes(
  ctx: CanvasRenderingContext2D,
  bodyWidth: number,
  bodyHeight: number,
  gridColor: string,
  selectionFill: string,
  textColor: string,
  worksheet: ActiveWorksheet | null,
): void {
  const sheet = activeSheet.value;
  const freezeRow = sheet?.freezePanes?.row ?? 0;
  const freezeCol = sheet?.freezePanes?.col ?? 0;
  if (!sheet || (freezeRow <= 0 && freezeCol <= 0)) return;
  const frozenRows = displayRows.value
    .map((actualIndex, displayIndex) => ({ actualIndex, displayIndex }))
    .filter((entry) => entry.actualIndex < freezeRow);
  const frozenCols = displayCols.value
    .map((actualIndex, displayIndex) => ({ actualIndex, displayIndex }))
    .filter((entry) => entry.actualIndex < freezeCol);
  const frozenHeight = frozenRows.reduce((sum, entry) => sum + (effectiveRowHeights.value[entry.displayIndex] || 0), 0);
  const frozenWidth = frozenCols.reduce((sum, entry) => sum + (effectiveColWidths.value[entry.displayIndex] || 0), 0);
  const visibleRows = Array.from(
    { length: Math.max(0, visibleRowRange.value.end - visibleRowRange.value.start + 1) },
    (_, offset) => visibleRowRange.value.start + offset,
  ).filter((index) => (displayRows.value[index] ?? 0) >= freezeRow);
  const visibleCols = Array.from(
    { length: Math.max(0, visibleColRange.value.end - visibleColRange.value.start + 1) },
    (_, offset) => visibleColRange.value.start + offset,
  ).filter((index) => (displayCols.value[index] ?? 0) >= freezeCol);

  const tableHeaderCells = props.getCellStyle ? collectTableHeaderCells() : EMPTY_CELL_KEYS;
  const paintState = createCellPaintStateScratch();
  const styleScratch = createCanvasCellStyleScratch();

  const paintCell = (displayRow: number, displayCol: number, frozenY: boolean, frozenX: boolean) => {
    const cell = getSheetCell(displayRow, displayCol);
    if (!cell || paintedMergeAnchors.has(cellKey(cell.row, cell.col))) return;
    const merge = mergeRegionAtAnchor(cell);
    const mergeEnd = merge ? getDisplayCell(merge.end) : null;
    const width = mergeEnd
      ? getColOffsetSum(mergeEnd.col + 1) - getColOffsetSum(displayCol)
      : effectiveColWidths.value[displayCol] || 0;
    const height = mergeEnd
      ? getRowOffsetSum(mergeEnd.row + 1) - getRowOffsetSum(displayRow)
      : effectiveRowHeights.value[displayRow] || 0;
    const x = paneViewOffset(getColOffsetSum(displayCol), scrollLeft.value, frozenX);
    const y = paneViewOffset(getRowOffsetSum(displayRow), scrollTop.value, frozenY);
    const value = getGridCellDisplayValue(cell);
    const cellKeyValue = cellKey(cell.row, cell.col);
    const style = resolveCanvasCellStyleInto(
      styleScratch,
      cell,
      value,
      resolveCellWorkbookStyle(cell, worksheet, sheet),
      tableHeaderCells,
    );
    const state = paintState;
    state.style = style;
    state.conditional = conditionalVisual(cell, value);
    state.comment = commentsByCell.value.get(cellKeyValue);
    state.hyperlink = hyperlinksByCell.value.get(cellKeyValue);
    state.sparkline = sparklinesByCell.value.get(cellKeyValue);
    state.value = value;
    state.gridLines = sheet.showGridLines !== false;
    state.selected = isCellInRange(cell, selection.value)
      || Boolean(merge && selection.value && rangesIntersect(merge, selection.value));
    state.selectionFill = selectionFill;
    state.gridColor = gridColor;
    state.textColor = textColor;
    state.overflowWidth = undefined;
    state.skipText = false;
    drawCellContent(ctx, x, y, width, height, state);
  };

  if (frozenRows.length && visibleCols.length) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(frozenWidth, 0, Math.max(0, bodyWidth - frozenWidth), frozenHeight);
    ctx.clip();
    for (const row of frozenRows) for (const col of visibleCols) paintCell(row.displayIndex, col, true, false);
    ctx.restore();
  }
  if (frozenCols.length && visibleRows.length) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, frozenHeight, frozenWidth, Math.max(0, bodyHeight - frozenHeight));
    ctx.clip();
    for (const row of visibleRows) for (const col of frozenCols) paintCell(row, col.displayIndex, false, true);
    ctx.restore();
  }
  if (frozenRows.length && frozenCols.length) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, frozenWidth, frozenHeight);
    ctx.clip();
    for (const row of frozenRows) for (const col of frozenCols) paintCell(row.displayIndex, col.displayIndex, true, true);
    ctx.restore();
  }
  ctx.save();
  ctx.strokeStyle = props.isDark ? "#71717a" : "#a1a1aa";
  ctx.lineWidth = 2;
  if (frozenWidth > 0) { ctx.beginPath(); ctx.moveTo(frozenWidth, 0); ctx.lineTo(frozenWidth, bodyHeight); ctx.stroke(); }
  if (frozenHeight > 0) { ctx.beginPath(); ctx.moveTo(0, frozenHeight); ctx.lineTo(bodyWidth, frozenHeight); ctx.stroke(); }
  ctx.restore();
}

function paintColHeaders() {
  const canvas = colHeaderCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const cw = Math.max(0, containerWidth.value - ROW_HEADER_WIDTH);
  resizeCanvas(canvas, cw, HEADER_HEIGHT);
  ctx.clearRect(0, 0, cw, HEADER_HEIGHT);

  const bgColor = props.isDark ? "#27272a" : "#f4f4f5";
  const textColor = props.isDark ? "#a1a1aa" : "#71717a";
  const borderColor = props.isDark ? "#3f3f46" : "#d4d4d8";
  const dividerColor = props.isDark ? "#71717a" : "#a1a1aa";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, cw, HEADER_HEIGHT);
  ctx.font = "11px -apple-system, BlinkMacSystemFont, sans-serif";
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";
  ctx.fillStyle = textColor;

  const widths = effectiveColWidths.value;
  const colRange = visibleColRange.value;
  const selected = selection.value;
  const frozen = frozenPaneMetrics();

  const paintColumn = (displayIndex: number) => {
    const colW = widths[displayIndex] || 0;
    // The frozen strip pins its labels above the pinned cells; only the
    // scrolling strip follows scrollLeft.
    const x = paneViewOffset(getColOffsetSum(displayIndex), scrollLeft.value, displayIndex < frozen.cols);
    const sheetCol = displayCols.value[displayIndex] ?? displayIndex;
    const selectedCol = activeCell.value?.col === sheetCol || Boolean(selected &&
      sheetCol >= Math.min(selected.start.col, selected.end.col) &&
      sheetCol <= Math.max(selected.start.col, selected.end.col));
    if (selectedCol) {
      ctx.fillStyle = props.isDark ? "#1e3a5f" : "#dbeafe";
      ctx.fillRect(x, 0, colW, HEADER_HEIGHT);
    }
    ctx.fillStyle = selectedCol ? (props.isDark ? "#bfdbfe" : "#1d4ed8") : textColor;
    ctx.fillText(columnLabel(sheetCol), x + colW / 2, HEADER_HEIGHT / 2);
    ctx.strokeStyle = borderColor; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(x + colW, 0); ctx.lineTo(x + colW, HEADER_HEIGHT); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(x, HEADER_HEIGHT - 0.5); ctx.lineTo(x + colW, HEADER_HEIGHT - 0.5); ctx.stroke();
  };

  ctx.save();
  ctx.beginPath();
  ctx.rect(frozen.width, 0, Math.max(0, cw - frozen.width), HEADER_HEIGHT);
  ctx.clip();
  for (let c = Math.max(colRange.start, frozen.cols); c <= colRange.end && c < widths.length; c++) {
    paintColumn(c);
  }
  ctx.restore();

  if (frozen.width > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, frozen.width, HEADER_HEIGHT);
    ctx.clip();
    for (let c = 0; c < frozen.cols && c < widths.length; c++) paintColumn(c);
    ctx.restore();
    ctx.strokeStyle = dividerColor;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(frozen.width - 0.5, 0);
    ctx.lineTo(frozen.width - 0.5, HEADER_HEIGHT);
    ctx.stroke();
  }
}

function paintRowHeaders() {
  const canvas = rowHeaderCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  const rh = Math.max(0, containerHeight.value - HEADER_HEIGHT);
  resizeCanvas(canvas, ROW_HEADER_WIDTH, rh);
  ctx.clearRect(0, 0, ROW_HEADER_WIDTH, rh);

  const bgColor = props.isDark ? "#27272a" : "#f4f4f5";
  const textColor = props.isDark ? "#a1a1aa" : "#71717a";
  const borderColor = props.isDark ? "#3f3f46" : "#d4d4d8";
  const dividerColor = props.isDark ? "#71717a" : "#a1a1aa";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, ROW_HEADER_WIDTH, rh);
  ctx.font = "11px -apple-system, BlinkMacSystemFont, sans-serif";
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";
  ctx.fillStyle = textColor;

  const heights = effectiveRowHeights.value;
  const rowRange = visibleRowRange.value;
  const selected = selection.value;
  const frozen = frozenPaneMetrics();

  const paintRow = (displayIndex: number) => {
    const rowH = heights[displayIndex] || 0;
    // The frozen strip pins its labels beside the pinned cells; only the
    // scrolling strip follows scrollTop.
    const y = paneViewOffset(getRowOffsetSum(displayIndex), scrollTop.value, displayIndex < frozen.rows);
    const sheetRow = displayRows.value[displayIndex] ?? displayIndex;
    const selectedRow = activeCell.value?.row === sheetRow || Boolean(selected &&
      sheetRow >= Math.min(selected.start.row, selected.end.row) &&
      sheetRow <= Math.max(selected.start.row, selected.end.row));
    if (selectedRow) {
      ctx.fillStyle = props.isDark ? "#1e3a5f" : "#dbeafe";
      ctx.fillRect(0, y, ROW_HEADER_WIDTH, rowH);
    }
    ctx.fillStyle = selectedRow ? (props.isDark ? "#bfdbfe" : "#1d4ed8") : textColor;
    ctx.fillText(String(sheetRow + 1), ROW_HEADER_WIDTH / 2, y + rowH / 2);
    ctx.strokeStyle = borderColor; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(0, y + rowH); ctx.lineTo(ROW_HEADER_WIDTH, y + rowH); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(ROW_HEADER_WIDTH - 0.5, y); ctx.lineTo(ROW_HEADER_WIDTH - 0.5, y + rowH); ctx.stroke();
  };

  ctx.save();
  ctx.beginPath();
  ctx.rect(0, frozen.height, ROW_HEADER_WIDTH, Math.max(0, rh - frozen.height));
  ctx.clip();
  for (let r = Math.max(rowRange.start, frozen.rows); r <= rowRange.end && r < heights.length; r++) {
    paintRow(r);
  }
  ctx.restore();

  if (frozen.height > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, ROW_HEADER_WIDTH, frozen.height);
    ctx.clip();
    for (let r = 0; r < frozen.rows && r < heights.length; r++) paintRow(r);
    ctx.restore();
    ctx.strokeStyle = dividerColor;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0, frozen.height - 0.5);
    ctx.lineTo(ROW_HEADER_WIDTH, frozen.height - 0.5);
    ctx.stroke();
  }
}

function paintCorner() {
  const canvas = cornerCanvasRef.value;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  resizeCanvas(canvas, ROW_HEADER_WIDTH, HEADER_HEIGHT);
  const bgColor = props.isDark ? "#27272a" : "#f4f4f5";
  const borderColor = props.isDark ? "#3f3f46" : "#d4d4d8";

  ctx.fillStyle = bgColor;
  ctx.fillRect(0, 0, ROW_HEADER_WIDTH, HEADER_HEIGHT);
  ctx.strokeStyle = borderColor; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(ROW_HEADER_WIDTH - 0.5, 0); ctx.lineTo(ROW_HEADER_WIDTH - 0.5, HEADER_HEIGHT); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(0, HEADER_HEIGHT - 0.5); ctx.lineTo(ROW_HEADER_WIDTH, HEADER_HEIGHT - 0.5); ctx.stroke();
}

// ── Event handlers ────────────────────────────────────────────────────
let paintFrame: number | null = null;

function schedulePaint() {
  if (paintFrame !== null) return;
  paintFrame = window.requestAnimationFrame(() => {
    paintFrame = null;
    const container = containerRef.value;
    if (!container) return;
    scrollTop.value = container.scrollTop;
    scrollLeft.value = container.scrollLeft;
    emit("viewportChange", {
      height: Math.max(0, container.clientHeight - HEADER_HEIGHT),
      scrollLeft: scrollLeft.value,
      scrollTop: scrollTop.value,
      width: Math.max(0, container.clientWidth - ROW_HEADER_WIDTH),
    });
    paintGrid();
    hydrateVisibleWorkerRows();
  });
}

function onScroll() {
  schedulePaint();
}

/**
 * Advances the pending-cut marquee a few times per second.
 *
 * A cut stays "armed" until it is pasted or cancelled, so the dashes have to
 * move while the user looks for the paste target. Each step only advances the
 * dash phase and reuses `schedulePaint()`, which coalesces the repaint with any
 * scroll that happens in the same frame.
 */
function scheduleCutAnts() {
  if (cutAntsFrame !== null) return;
  if (!cutRange.value) return;
  if (typeof window === "undefined" || typeof window.requestAnimationFrame !== "function") return;
  cutAntsFrame = window.requestAnimationFrame((timestamp: number) => {
    cutAntsFrame = null;
    if (!cutRange.value) return;
    if (timestamp - cutAntsLastStepAt >= CUT_ANTS_STEP_MS) {
      cutAntsLastStepAt = timestamp;
      cutAntsStep.value = (cutAntsStep.value + 1) % CUT_ANTS_DASH_PERIOD;
      schedulePaint();
    }
    scheduleCutAnts();
  });
}

function stopCutAnts() {
  if (cutAntsFrame === null) return;
  if (typeof window !== "undefined" && typeof window.cancelAnimationFrame === "function") {
    window.cancelAnimationFrame(cutAntsFrame);
  }
  cutAntsFrame = null;
}

type AxisHit = {
  actualIndex: number;
  displayIndex: number;
  kind: "column" | "row" | "corner";
  resize: boolean;
};

function hitTestAxis(clientX: number, clientY: number): AxisHit | null {
  const container = containerRef.value;
  if (!container) return null;
  const rect = container.getBoundingClientRect();
  const localX = clientX - rect.left;
  const localY = clientY - rect.top;
  if (localX < 0 || localY < 0 || localX > rect.width || localY > rect.height) return null;
  if (localX <= ROW_HEADER_WIDTH && localY <= HEADER_HEIGHT) {
    return { actualIndex: 0, displayIndex: 0, kind: "corner", resize: false };
  }
  if (localY <= HEADER_HEIGHT && localX > ROW_HEADER_WIDTH) {
    const frozen = frozenPaneMetrics();
    const contentX = paneAxisContentOffset(localX - ROW_HEADER_WIDTH, container.scrollLeft, frozen.width);
    const hitDisplayIndex = findDisplayIndexAtOffset(colOffsets.value, contentX);
    if (hitDisplayIndex < 0) return null;
    const startBoundary = getColOffsetSum(hitDisplayIndex);
    const usePreviousBoundary = hitDisplayIndex > 0 && Math.abs(contentX - startBoundary) <= 5;
    const displayIndex = usePreviousBoundary ? hitDisplayIndex - 1 : hitDisplayIndex;
    const actualIndex = displayCols.value[displayIndex];
    if (actualIndex === undefined) return null;
    const boundary = usePreviousBoundary ? startBoundary : getColOffsetSum(displayIndex + 1);
    return {
      actualIndex,
      displayIndex,
      kind: "column",
      resize: Math.abs(contentX - boundary) <= 5,
    };
  }
  if (localX <= ROW_HEADER_WIDTH && localY > HEADER_HEIGHT) {
    const frozen = frozenPaneMetrics();
    const contentY = paneAxisContentOffset(localY - HEADER_HEIGHT, container.scrollTop, frozen.height);
    const hitDisplayIndex = findDisplayIndexAtOffset(rowOffsets.value, contentY);
    if (hitDisplayIndex < 0) return null;
    const startBoundary = getRowOffsetSum(hitDisplayIndex);
    const usePreviousBoundary = hitDisplayIndex > 0 && Math.abs(contentY - startBoundary) <= 5;
    const displayIndex = usePreviousBoundary ? hitDisplayIndex - 1 : hitDisplayIndex;
    const actualIndex = displayRows.value[displayIndex];
    if (actualIndex === undefined) return null;
    const boundary = usePreviousBoundary ? startBoundary : getRowOffsetSum(displayIndex + 1);
    return {
      actualIndex,
      displayIndex,
      kind: "row",
      resize: Math.abs(contentY - boundary) <= 5,
    };
  }
  return null;
}

function selectAxisRange(axisHit: Pick<AxisHit, "actualIndex" | "kind">) {
  const firstRow = displayRows.value[0] ?? 0;
  const lastRow = displayRows.value[displayRows.value.length - 1] ?? firstRow;
  const firstCol = displayCols.value[0] ?? 0;
  const lastCol = displayCols.value[displayCols.value.length - 1] ?? firstCol;
  if (axisHit.kind === "column") {
    props.controller.selectRange({
      start: { row: firstRow, col: axisHit.actualIndex },
      end: { row: lastRow, col: axisHit.actualIndex },
    });
  } else if (axisHit.kind === "row") {
    props.controller.selectRange({
      start: { row: axisHit.actualIndex, col: firstCol },
      end: { row: axisHit.actualIndex, col: lastCol },
    });
  } else {
    props.controller.selectRange({
      start: { row: firstRow, col: firstCol },
      end: { row: lastRow, col: lastCol },
    });
  }
}

function onPointerDown(event: PointerEvent) {
  // Only left-click (button 0) should change selection.
  // Right-click (button 2) must preserve the existing selection
  // so the context menu operates on the intended range.
  if (event.button !== 0) return
  const axisHit = hitTestAxis(event.clientX, event.clientY);
  if (axisHit?.resize && axisHit.kind !== "corner") {
    const z = gridZoomFactor(zoomScale.value);
    const renderedSize = axisHit.kind === "column"
      ? effectiveColWidths.value[axisHit.displayIndex] ?? DEFAULT_COL_WIDTH * z
      : effectiveRowHeights.value[axisHit.displayIndex] ?? DEFAULT_ROW_HEIGHT * z;
    resizePreview.value = {
      actualIndex: axisHit.actualIndex,
      displayIndex: axisHit.displayIndex,
      kind: axisHit.kind,
      moved: false,
      pointerId: event.pointerId,
      size: renderedSize / z,
      startPosition: axisHit.kind === "column" ? event.clientX : event.clientY,
      startSize: renderedSize / z,
    };
    containerRef.value?.focus();
    containerRef.value?.setPointerCapture(event.pointerId);
    event.preventDefault();
    return;
  }
  if (axisHit) {
    selectAxisRange(axisHit);
    containerRef.value?.focus();
    event.preventDefault();
    return;
  }
  const cell = hitTestCell(event.clientX, event.clientY);
  if (!cell) return;
  containerRef.value?.focus();

  if (event.shiftKey && activeCell.value) {
    props.controller.selectRange({ start: activeCell.value, end: cell });
  } else {
    props.controller.selectCell(cell, { extend: event.shiftKey });
  }
  isSelecting.value = true;
  selectionAnchor.value = cell;
  containerRef.value?.setPointerCapture(event.pointerId);
}

function onPointerMove(event: PointerEvent) {
  const resizing = resizePreview.value;
  if (resizing) {
    const position = resizing.kind === "column" ? event.clientX : event.clientY;
    const delta = position - resizing.startPosition;
    if (!resizing.moved && Math.abs(delta) < AXIS_RESIZE_DRAG_THRESHOLD) return;
    const minimum = resizing.kind === "column" ? 32 : 18;
    resizing.moved = true;
    resizing.size = resolveAxisResizeSize(resizing.startSize, delta, zoomScale.value, minimum);
    resizePreview.value = { ...resizing };
    schedulePaint();
    return;
  }
  const axisHit = hitTestAxis(event.clientX, event.clientY);
  hoveredCell.value = axisHit ? null : hitTestCell(event.clientX, event.clientY);
  hoveredAxis.value = axisHit && !axisHit.resize && axisHit.kind !== "corner"
    ? {
        actualIndex: axisHit.actualIndex,
        displayIndex: axisHit.displayIndex,
        kind: axisHit.kind,
      }
    : null;
  if (containerRef.value) {
    containerRef.value.style.cursor = axisHit?.resize
      ? axisHit.kind === "column" ? "col-resize" : "row-resize"
      : axisHit ? "default" : hoveredCell.value && hyperlinksByCell.value.has(cellKey(hoveredCell.value.row, hoveredCell.value.col)) ? "pointer" : "cell";
  }
  if (!isSelecting.value || !selectionAnchor.value) return;
  const cell = hitTestCell(event.clientX, event.clientY);
  if (!cell) return;
  props.controller.selectRange({ start: selectionAnchor.value, end: cell });
}

function resolveAxisResizeRelease(
  moved: boolean,
  eventType?: string,
): "cancel" | "resize" | "select" {
  if (eventType === "pointercancel") return "cancel";
  return moved ? "resize" : "select";
}

function onPointerUp(event?: PointerEvent) {
  const clickedCell = event ? hitTestCell(event.clientX, event.clientY) : null;
  const pressCell = selectionAnchor.value;
  const resizing = resizePreview.value;
  if (resizing) {
    const releaseAction = resolveAxisResizeRelease(resizing.moved, event?.type);
    if (releaseAction === "resize") {
      if (resizing.kind === "column") props.controller.resizeColumn(resizing.actualIndex, resizing.size);
      else props.controller.resizeRow(resizing.actualIndex, resizing.size);
    } else if (releaseAction === "select") {
      selectAxisRange(resizing);
    }
    resizePreview.value = null;
  }
  isSelecting.value = false;
  selectionAnchor.value = null;
  if (event && containerRef.value?.hasPointerCapture(event.pointerId)) {
    containerRef.value.releasePointerCapture(event.pointerId);
  }
  schedulePaint();
  if (clickedCell && pressCell && clickedCell.row === pressCell.row && clickedCell.col === pressCell.col) {
    activateHyperlink(clickedCell);
  }
}

function onPointerLeave(): void {
  hoveredCell.value = null;
  hoveredAxis.value = null;
  if (!resizePreview.value && !isSelecting.value && containerRef.value) {
    containerRef.value.style.cursor = "cell";
  }
}

function activateHyperlink(cell: XlsxCellAddress): void {
  const hyperlink = hyperlinksByCell.value.get(cellKey(cell.row, cell.col)) as XlsxCellHyperlink | undefined;
  if (!hyperlink) return;
  if (hyperlink.target.startsWith("#")) {
    const reference = hyperlink.target.slice(1);
    const [sheetPart, cellPart = sheetPart] = reference.split("!");
    const sheetName = cellPart === sheetPart ? null : sheetPart.replace(/^'|'$/g, "").replace(/''/g, "'");
    const match = /\$?([A-Z]+)\$?(\d+)/i.exec(cellPart);
    if (!match) return;
    if (sheetName) {
      const sheetIndex = props.controller.sheets.findIndex((sheet) => sheet.name === sheetName);
      if (sheetIndex >= 0) props.controller.setActiveSheetIndex(sheetIndex);
    }
    let col = 0;
    for (const char of match[1].toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64;
    props.controller.selectCell({ row: Number(match[2]) - 1, col: col - 1 });
    return;
  }
  const target = sanitizeOfficeLink(hyperlink.target);
  if (!target) return;
  try {
    const url = new URL(target, globalThis.location?.href);
    if (!["http:", "https:", "mailto:"].includes(url.protocol)) return;
    globalThis.open?.(url.href, "_blank", "noopener,noreferrer");
  } catch { /* invalid links stay inert */ }
}

function getCellInputValue(cell: XlsxCellAddress): string {
  // Seed the inline editor with the formula so editing a formula cell and
  // committing it again cannot replace the formula with its displayed value.
  return props.controller.getCellFormula(cell) || props.controller.getCellDisplayValue(cell);
}

function openValidationDropdown() {
  if (editingLocked.value || !activeListValidation.value || validationDropdownOptions.value.length === 0) return;
  validationDropdownOpen.value = true;
}

function closeValidationDropdown() {
  validationDropdownOpen.value = false;
}

function chooseValidationOption(option: string) {
  const cell = activeCell.value;
  validationDropdownOpen.value = false;
  if (!cell || editingLocked.value) return;
  props.controller.setCellValue(cell, option);
  void nextTick(() => containerRef.value?.focus());
}

function onDblClick(event: MouseEvent) {
  const cell = hitTestCell(event.clientX, event.clientY);
  if (!cell || props.readOnly || props.controller?.isEditingLocked === true) return;
  const validation = props.controller?.getCellDataValidation?.(cell) ?? null;
  if (validation && isDataValidationListDropdownVisible(validation)) {
    props.controller?.selectCell?.(cell);
    emit("cellDoubleClick", cell);
    nextTick(openValidationDropdown);
    return;
  }
  editingCell.value = cell;
  editingValue.value = getCellInputValue(cell);
  // Bring the cell fully into its pane first: an editor opened over a cell
  // still half-hidden under a frozen strip would sit on the freeze line.
  scrollToCell(cell);
  emit("cellDoubleClick", cell);
  nextTick(() => { editInputRef.value?.focus(); editInputRef.value?.select(); });
}

function commitEdit() {
  if (!editingCell.value) return;
  props.controller.commitCellInput(editingCell.value, editingValue.value);
  editingCell.value = null;
  editingValue.value = "";
  isEditComposing.value = false;
  commitEditOnCompositionEnd = false;
  void nextTick(() => containerRef.value?.focus());
}

function cancelEdit() {
  editingCell.value = null;
  editingValue.value = "";
  isEditComposing.value = false;
  commitEditOnCompositionEnd = false;
  void nextTick(() => containerRef.value?.focus());
}

function isComposingKeyboardEvent(event: KeyboardEvent) {
  // keyCode 229 is the legacy signal browsers emit while an IME owns the key.
  return event.isComposing || event.keyCode === 229;
}

/** Hard line breaks in the draft; soft wraps are measured by the mirror. */
function countEditorLines(value: string): number {
  return value.replace(/\r\n?/g, "\n").split("\n").length;
}

const editorLineCount = computed(() => countEditorLines(editingValue.value));

function editorLineHeightPx(): number {
  return Math.max(EDITOR_MIN_LINE_HEIGHT_PX, 12 * gridZoomFactor(zoomScale.value) * EDITOR_LINE_HEIGHT_RATIO);
}

function readEditorSelection(): { start: number; end: number } {
  const value = editingValue.value;
  const textarea = editInputRef.value;
  const rawStart = textarea?.selectionStart;
  const rawEnd = textarea?.selectionEnd;
  const start = typeof rawStart === "number" && Number.isFinite(rawStart) ? rawStart : value.length;
  const end = typeof rawEnd === "number" && Number.isFinite(rawEnd) ? rawEnd : start;
  return {
    start: Math.max(0, Math.min(start, value.length)),
    end: Math.max(0, Math.min(end, value.length)),
  };
}

/** Inserts text at the textarea caret, keeping the draft in one commit. */
function insertEditText(text: string) {
  const normalized = text.replace(/\r\n?/g, "\n");
  const value = editingValue.value;
  const { start, end } = readEditorSelection();
  const insertionStart = Math.min(start, end);
  const insertionEnd = Math.max(start, end);
  editingValue.value = value.slice(0, insertionStart) + normalized + value.slice(insertionEnd);
  const caret = insertionStart + normalized.length;
  void nextTick(() => {
    const textarea = editInputRef.value;
    textarea?.focus();
    if (textarea && typeof textarea.setSelectionRange === "function") {
      textarea.setSelectionRange(caret, caret);
    }
    syncEditorSize();
  });
}

/**
 * Multi-line clipboard text keeps its line breaks. Single-line pastes stay on
 * the native path so the browser keeps its caret and undo stack.
 */
function onEditPaste(event: ClipboardEvent) {
  if (isEditComposing.value) return;
  const clipboard = event.clipboardData;
  if (!clipboard) return;
  const text = clipboard.getData("text/plain") || clipboard.getData("text") || "";
  if (!/[\r\n]/.test(text)) return;
  event.preventDefault();
  event.stopPropagation();
  insertEditText(text);
}

/**
 * Measures the hidden mirror, which carries the draft with the textarea's
 * typography and width. Its height is how many wrapped lines the editor has to
 * show; hard breaks are a deterministic fallback before the mirror has layout.
 */
function measureEditorMirrorHeight(): number {
  const mirror = editMirrorRef.value;
  if (!mirror) return 0;
  const offset = Number(mirror.offsetHeight);
  if (Number.isFinite(offset) && offset > 0) return offset;
  const scroll = Number(mirror.scrollHeight);
  if (!Number.isFinite(scroll) || scroll <= 0) return 0;
  // scrollHeight includes padding but not the border; the editor is border-box.
  return scroll + EDITOR_BORDER_PX * 2;
}

function syncEditorSize() {
  if (!editingCell.value) {
    editorMeasuredHeight.value = 0;
    return;
  }
  const rect = paneOverlayRect(editingCell.value);
  const cellHeight = rect?.height || DEFAULT_ROW_HEIGHT;
  const hardBreakHeight = editorLineCount.value > 1
    ? Math.ceil(editorLineCount.value * editorLineHeightPx() + EDITOR_BOX_CHROME_PX)
    : 0;
  editorMeasuredHeight.value = Math.max(cellHeight, hardBreakHeight, measureEditorMirrorHeight());
}

function onEditKeydown(event: KeyboardEvent) {
  if (isEditComposing.value || isComposingKeyboardEvent(event)) {
    // Let the IME consume the key. A plain Enter confirms the composition;
    // remember it so the cell is committed once the final composition text is
    // in place. Alt+Enter stays a composition key, not a commit.
    if (event.key === "Enter" && !event.altKey) commitEditOnCompositionEnd = true;
    return;
  }

  if (event.key === "Enter" && event.altKey) {
    // Alt+Enter inserts a hard line break without closing the editor, so the
    // cell can be authored with multiple lines (Excel's shortcut).
    event.stopPropagation();
    event.preventDefault();
    insertEditText("\n");
  } else if (event.key === "Enter") {
    event.stopPropagation();
    event.preventDefault();
    commitEdit();
  } else if (event.key === "Escape") {
    event.stopPropagation();
    event.preventDefault();
    cancelEdit();
  } else if (event.key === "Tab") {
    event.stopPropagation();
    event.preventDefault();
    commitAndTab();
  } else if (
    (event.key === "ArrowUp" || event.key === "ArrowDown")
    && !event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey
    && editorLineCount.value === 1
  ) {
    // A single-line draft has no text for ↑/↓ to move through, so Excel
    // commits it and moves the selection out of the cell. A multi-line draft
    // keeps the arrows for caret movement between its lines.
    event.stopPropagation();
    event.preventDefault();
    commitAndMoveCell(event.key === "ArrowUp" ? -1 : 1, 0);
  }
}

function onEditCompositionStart() {
  isEditComposing.value = true;
  commitEditOnCompositionEnd = false;
}

function onEditCompositionEnd() {
  isEditComposing.value = false;
  if (!commitEditOnCompositionEnd) return;
  commitEditOnCompositionEnd = false;
  // v-model writes the final composition text during compositionend listeners,
  // so wait one tick before reading editingValue.
  void nextTick(commitEdit);
}

function commitAndTab() {
  if (!editingCell.value) return;
  props.controller.commitCellInput(editingCell.value, editingValue.value);
  const nextCell = getAdjacentCell(editingCell.value, 0, 1);
  if (!nextCell) return;
  props.controller.selectCell(nextCell);
  editingCell.value = nextCell;
  editingValue.value = getCellInputValue(nextCell);
  scrollToCell(nextCell);
}

/** Commits the draft and leaves the editor, moving the selection by one cell. */
function commitAndMoveCell(rowDelta: number, colDelta: number) {
  const cell = editingCell.value;
  if (!cell) return;
  props.controller.commitCellInput(cell, editingValue.value);
  editingCell.value = null;
  editingValue.value = "";
  isEditComposing.value = false;
  commitEditOnCompositionEnd = false;
  const nextCell = getAdjacentCell(cell, rowDelta, colDelta);
  if (nextCell) {
    props.controller.selectCell(nextCell);
    scrollToCell(nextCell);
  }
  void nextTick(() => containerRef.value?.focus());
}

function getAdjacentCell(cell: XlsxCellAddress, rowDelta: number, colDelta: number): XlsxCellAddress | null {
  return resolveAdjacentSheetCell(
    cell,
    rowDelta,
    colDelta,
    displayRows.value,
    displayCols.value,
    sheetRowToDisplayIndex.value,
    sheetColToDisplayIndex.value,
  );
}

function moveCell(rowDelta: number, colDelta: number, extend: boolean) {
  const current = activeCell.value;
  if (!current) return;
  const cell = getAdjacentCell(current, rowDelta, colDelta);
  if (!cell) return;
  if (extend) {
    props.controller.selectRange({ start: current, end: cell });
  } else {
    props.controller.selectCell(cell);
  }
  scrollToCell(cell);
}

function startEdit(cell: XlsxCellAddress) {
  editingCell.value = cell;
  editingValue.value = getCellInputValue(cell);
  scrollToCell(cell);
  nextTick(() => editInputRef.value?.focus());
}

// The context menu, ribbon and hosts can request an inline edit without
// reaching into the grid; the request arrives as data and is consumed here.
watch(
  () => props.controller.editRequest,
  (request) => {
    if (!request || editingLocked.value) return;
    props.controller.selectCell(request.cell);
    startEdit(request.cell);
  },
);

function scrollToCell(cell: XlsxCellAddress) {
  const container = containerRef.value;
  if (!container) return;
  const displayCell = getDisplayCell(cell);
  if (!displayCell) return;
  const frozen = frozenPaneMetrics();
  const colX = getColOffsetSum(displayCell.col);
  const rowY = getRowOffsetSum(displayCell.row);
  const colW = effectiveColWidths.value[displayCell.col] || DEFAULT_COL_WIDTH;
  const rowH = effectiveRowHeights.value[displayCell.row] || DEFAULT_ROW_HEIGHT;
  const bodyWidth = Math.max(0, container.clientWidth - ROW_HEADER_WIDTH);
  const bodyHeight = Math.max(0, container.clientHeight - HEADER_HEIGHT);

  // Pinned cells are always on screen — scrolling them into view would jump
  // the sheet for no reason. In the scrolling pane the freeze line hides the
  // first frozenWidth pixels, so a cell is only fully visible once its left
  // edge clears the strip and its right edge clears the body's right edge.
  if (displayCell.col >= frozen.cols) {
    if (colX - container.scrollLeft < frozen.width) container.scrollLeft = Math.max(0, colX - frozen.width);
    else if (colX + colW - container.scrollLeft > bodyWidth) container.scrollLeft = colX + colW - bodyWidth;
  }
  if (displayCell.row >= frozen.rows) {
    if (rowY - container.scrollTop < frozen.height) container.scrollTop = Math.max(0, rowY - frozen.height);
    else if (rowY + rowH - container.scrollTop > bodyHeight) container.scrollTop = rowY + rowH - bodyHeight;
  }
  schedulePaint();
}

function onGridKeydown(event: KeyboardEvent) {
  // When the inline edit input is active, let it consume all keys —
  // do not re-enter edit mode and overwrite the editing value.
  if (editingCell.value) return
  // An IME composition owns the keyboard; never treat its keys as grid commands.
  if (isComposingKeyboardEvent(event)) return

  // A pending cut is a grid mode, not a selection: Escape leaves the selection
  // alone and only drops the marquee (Excel's cancel-cut). The check is on the
  // controller, not on this sheet's marquee, so a cut made on another sheet is
  // still cancellable from here; drawings carry their own pending cut with no
  // cell range, so the return value decides whether Escape was consumed.
  if (event.key === "Escape") {
    if (props.controller.cancelCut?.()) {
      event.preventDefault();
      return;
    }
  }

  const active = activeCell.value;
  if (!active) return;

  if ((event.ctrlKey || event.metaKey) && event.key === "c") { event.preventDefault(); void copySelection(); return; }
  if ((event.ctrlKey || event.metaKey) && event.key === "x") { event.preventDefault(); if (!editingLocked.value) props.controller.cutSelection?.(); return; }
  if ((event.ctrlKey || event.metaKey) && event.key === "v") { event.preventDefault(); props.controller.pasteFromClipboard(); return; }

  const handlers: Record<string, () => void> = {
    ArrowUp: () => moveCell(-1, 0, event.shiftKey),
    ArrowDown: () => moveCell(1, 0, event.shiftKey),
    ArrowLeft: () => moveCell(0, -1, event.shiftKey),
    ArrowRight: () => moveCell(0, 1, event.shiftKey),
    Tab: () => { event.preventDefault(); moveCell(0, 1, false); },
    Enter: () => { if (!editingLocked.value) startEdit(active); },
    Delete: () => { if (!editingLocked.value) props.controller.clearSelectedCells(); },
    Backspace: () => { if (!editingLocked.value) props.controller.setCellValue(active, ""); },
    F2: () => { if (!editingLocked.value) startEdit(active); },
  };

  const handler = handlers[event.key];
  if (handler) { event.preventDefault(); handler(); return; }

  if (!editingLocked.value && event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    editingCell.value = active;
    editingValue.value = event.key;
    nextTick(() => editInputRef.value?.focus());
  }
}

// ── Styles ────────────────────────────────────────────────────────────
const containerStyle: CSSProperties = {
  flex: "1",
  minHeight: "0px",
  minWidth: "0px",
  outline: "none",
  overflow: "auto",
  position: "relative",
};
const scrollSpaceStyle = computed<CSSProperties>(() => ({
  height: `${Math.max(containerHeight.value, HEADER_HEIGHT + (rowOffsets.value.at(-1) ?? 0))}px`,
  width: `${Math.max(containerWidth.value, ROW_HEADER_WIDTH + (colOffsets.value.at(-1) ?? 0))}px`,
}));
const viewportStyle = computed<CSSProperties>(() => ({
  height: `${containerHeight.value}px`,
  width: `${containerWidth.value}px`,
}));
const bodyCanvasStyle: CSSProperties = { left: `${ROW_HEADER_WIDTH}px`, position: "absolute", top: `${HEADER_HEIGHT}px` };
const colHeaderCanvasStyle: CSSProperties = { left: `${ROW_HEADER_WIDTH}px`, position: "absolute", top: "0px" };
const rowHeaderCanvasStyle: CSSProperties = { left: "0px", position: "absolute", top: `${HEADER_HEIGHT}px` };
const cornerCanvasStyle: CSSProperties = { left: "0px", position: "absolute", top: "0px" };

const editInputStyle = computed<CSSProperties>(() => {
  if (!editingCell.value) return { display: "none" };
  const z = gridZoomFactor(zoomScale.value);
  const base = paneOverlayRect(editingCell.value);
  if (!base) return { display: "none" };
  // Grow over the rows below for wrapped or hard-broken lines. The pane clip
  // still wins, so a cell at the bottom edge never paints over the headers.
  const rect = paneOverlayRect(editingCell.value, {
    height: Math.max(base.height, editorMeasuredHeight.value),
  });
  if (!rect) return { display: "none" };
  return {
    left: `${ROW_HEADER_WIDTH + rect.left}px`,
    top: `${HEADER_HEIGHT + rect.top}px`,
    width: `${rect.width}px`,
    height: `${rect.height}px`,
    position: "absolute", zIndex: "10",
    border: `${EDITOR_BORDER_PX}px solid #2563eb`,
    padding: `${EDITOR_PADDING_Y_PX}px ${EDITOR_PADDING_X_PX}px`,
    fontSize: `${12 * z}px`, fontFamily: "inherit", lineHeight: `${editorLineHeightPx()}px`,
    outline: "none", resize: "none", overflow: "hidden",
    whiteSpace: "pre-wrap", overflowWrap: "break-word", wordBreak: "break-word",
    background: props.isDark ? "#18181b" : "#ffffff",
    color: props.isDark ? "#e4e4e7" : "#18181b",
    boxSizing: "border-box",
  };
});

/**
 * Hidden measurement twin of the editor. It carries the same text, width,
 * typography, padding and border, so its natural height is exactly the editor
 * box height the draft needs.
 */
const editMirrorStyle = computed<CSSProperties>(() => {
  if (!editingCell.value) return { display: "none" };
  const z = gridZoomFactor(zoomScale.value);
  const rect = paneOverlayRect(editingCell.value);
  if (!rect) return { display: "none" };
  return {
    left: `${ROW_HEADER_WIDTH + rect.left}px`,
    top: `${HEADER_HEIGHT + rect.top}px`,
    width: `${rect.width}px`,
    position: "absolute", zIndex: "9",
    border: `${EDITOR_BORDER_PX}px solid transparent`,
    padding: `${EDITOR_PADDING_Y_PX}px ${EDITOR_PADDING_X_PX}px`,
    fontSize: `${12 * z}px`, fontFamily: "inherit", lineHeight: `${editorLineHeightPx()}px`,
    whiteSpace: "pre-wrap", overflowWrap: "break-word", wordBreak: "break-word",
    visibility: "hidden", pointerEvents: "none", boxSizing: "border-box",
  };
});

defineExpose({
  captureZoomAnchor,
  restoreZoomAnchor,
  scrollToCell,
  hitTestCell,
  hitTestAxis,
  hitTestSheetPoint,
  isEditingCell,
  get scrollContainer() {
    return containerRef.value;
  },
});

// ── Watchers ──────────────────────────────────────────────────────────
watch(
  () => [
    revision.value,
    workerRowsRevision.value,
    devicePixelRatioRevision.value,
    zoomScale.value,
    containerWidth.value,
    containerHeight.value,
    activeSheet.value,
    activeCell.value?.row,
    activeCell.value?.col,
    selection.value?.start.row,
    selection.value?.start.col,
    selection.value?.end.row,
    selection.value?.end.col,
    cutRange.value?.start.row,
    cutRange.value?.start.col,
    cutRange.value?.end.row,
    cutRange.value?.end.col,
    props.isDark,
    props.selectionColor,
    props.selectionFillColor,
  ],
  schedulePaint,
  { flush: "post" },
);

watch(
  cutRange,
  (range) => {
    cutAntsStep.value = 0;
    cutAntsLastStepAt = 0;
    if (range) scheduleCutAnts();
    else stopCutAnts();
  },
  { flush: "post" },
);

watch(
  () => [activeSheet.value?.name, activeCell.value?.row, activeCell.value?.col],
  () => { validationDropdownOpen.value = false; },
  { flush: "post" },
);

// The mirror is the editor's sizing source: measure it after every draft or
// zoom change and let the textarea grow instead of scrolling its first line.
watch(
  [editingCell, editingValue, zoomScale, containerWidth, containerHeight],
  syncEditorSize,
  { flush: "post" },
);

let resizeObs: ResizeObserver | null = null;
let dprMediaQuery: MediaQueryList | null = null;
let dprMediaListener: (() => void) | null = null;

function unbindDevicePixelRatioWatch(): void {
  if (!dprMediaQuery) return;
  const listener = dprMediaListener;
  if (listener) {
    if (typeof dprMediaQuery.removeEventListener === "function") dprMediaQuery.removeEventListener("change", listener);
    else if (typeof dprMediaQuery.removeListener === "function") dprMediaQuery.removeListener(listener);
  }
  dprMediaQuery = null;
  dprMediaListener = null;
}

// `devicePixelRatio` changes when the window moves between monitors with
// different scaling or the browser zoom changes. Re-arm the media query for the
// new ratio and repaint, otherwise the backing store stays at the old
// resolution until some unrelated resize happens.
function bindDevicePixelRatioWatch(): void {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") return;
  unbindDevicePixelRatioWatch();
  try {
    dprMediaQuery = window.matchMedia(`(resolution: ${readDevicePixelRatio()}dppx)`);
  } catch {
    dprMediaQuery = null;
    return;
  }
  dprMediaListener = () => {
    devicePixelRatioRevision.value += 1;
    bindDevicePixelRatioWatch();
    schedulePaint();
  };
  if (typeof dprMediaQuery.addEventListener === "function") dprMediaQuery.addEventListener("change", dprMediaListener);
  else if (typeof dprMediaQuery.addListener === "function") dprMediaQuery.addListener(dprMediaListener);
}

onMounted(() => {
  const container = containerRef.value;
  if (!container) return;
  containerWidth.value = container.clientWidth;
  containerHeight.value = container.clientHeight;
  bindDevicePixelRatioWatch();
  resizeObs = new ResizeObserver((entries) => {
    for (const entry of entries) { containerWidth.value = entry.contentRect.width; containerHeight.value = entry.contentRect.height; }
  });
  resizeObs.observe(container);
  schedulePaint();
  if (cutRange.value) scheduleCutAnts();
});

onUnmounted(() => {
  resizeObs?.disconnect();
  unbindDevicePixelRatioWatch();
  stopCutAnts();
  resetWorkerRows(null);
  if (paintFrame !== null) window.cancelAnimationFrame(paintFrame);
  paintFrame = null;
});
</script>

<style scoped>
.xlsx-grid { outline: none; touch-action: pan-x pan-y; }
.xlsx-grid__a11y, .xlsx-grid__a11y-cell {
  block-size: 1px;
  clip-path: inset(50%);
  inline-size: 1px;
  overflow: hidden;
  pointer-events: none;
  position: absolute;
  white-space: nowrap;
}
.xlsx-grid__scroll-space { position: relative; }
.xlsx-grid__page-breaks { inset: 0; pointer-events: none; position: absolute; z-index: 6; }
.xlsx-grid__page-break-line { border: 0; display: block; margin: 0; padding: 0; position: absolute; }
.xlsx-grid__page-break-line--column { border-left: 1px dashed #2563eb; width: 0; }
.xlsx-grid__page-break-line--row { border-top: 1px dashed #2563eb; height: 0; }
.xlsx-grid__page-break-label { background: rgba(37, 99, 235, .92); border-radius: 3px; color: #ffffff; font-family: inherit; font-size: 10px; font-style: normal; font-weight: 600; line-height: 1; padding: 2px 4px; position: absolute; }
.xlsx-grid__viewport {
  left: 0;
  overflow: hidden;
  pointer-events: none;
  position: sticky;
  top: 0;
}
.xlsx-grid__col-header,
.xlsx-grid__row-header,
.xlsx-grid__corner,
.xlsx-grid__body { pointer-events: none; }
.xlsx-grid__edit-input { box-sizing: border-box; pointer-events: auto; }
.xlsx-grid__edit-mirror { pointer-events: none; position: absolute; visibility: hidden; }
.xlsx-grid__edit-mirror::after { content: "\200b"; }
.xlsx-grid__comment-popover { background: #fffbe6; border: 1px solid #eab308; border-radius: 6px; box-shadow: 0 8px 24px rgba(0,0,0,.18); color: #422006; display: flex; flex-direction: column; gap: 3px; max-width: 260px; min-width: 150px; padding: 8px 10px; pointer-events: none; position: absolute; white-space: pre-wrap; z-index: 30; }
.xlsx-grid__comment-popover strong { font-size: 11px; }
.xlsx-grid__comment-popover span { font-size: 12px; line-height: 1.4; }
.xlsx-grid__validation-hint { background: #eff6ff; border: 1px solid #93c5fd; border-radius: 6px; box-shadow: 0 8px 24px rgba(0,0,0,.12); color: #1e3a8a; display: flex; flex-direction: column; gap: 3px; max-width: 260px; min-width: 140px; padding: 6px 9px; pointer-events: none; position: absolute; white-space: pre-wrap; z-index: 28; }
.xlsx-grid__validation-hint strong { font-size: 11px; }
.xlsx-grid__validation-hint span { font-size: 12px; line-height: 1.4; }
.xlsx-grid__validation-trigger { align-items: center; background: #ffffff; border: 1px solid #94a3b8; border-radius: 3px; color: #334155; cursor: pointer; display: flex; font-size: 10px; height: 18px; justify-content: center; padding: 0; pointer-events: auto; position: absolute; width: 16px; z-index: 29; }
.xlsx-grid__validation-trigger:hover { background: #e2e8f0; }
.xlsx-grid__validation-dropdown { background: #ffffff; border: 1px solid #94a3b8; border-radius: 4px; box-shadow: 0 8px 24px rgba(0,0,0,.18); display: flex; flex-direction: column; max-height: 220px; overflow-y: auto; padding: 3px; pointer-events: auto; position: absolute; z-index: 32; }
.xlsx-grid__validation-option { background: transparent; border: 0; border-radius: 3px; color: #18181b; cursor: pointer; font: inherit; font-size: 12px; padding: 5px 8px; text-align: left; }
.xlsx-grid__validation-option:hover { background: #dbeafe; }
.xlsx-grid__axis-menu-trigger {
  align-items: center;
  background: transparent;
  border: 0;
  color: #71717a;
  cursor: pointer;
  display: flex;
  font-size: 10px;
  height: 16px;
  justify-content: center;
  padding: 0;
  pointer-events: auto;
  position: absolute;
  width: 14px;
  z-index: 20;
}
.xlsx-grid__axis-menu-trigger:hover { background: rgba(113, 113, 122, .2); border-radius: 3px; color: #18181b; }
.xlsx-grid__validation-alert { align-items: flex-start; border: 1px solid #f87171; border-radius: 6px; box-shadow: 0 8px 24px rgba(0,0,0,.18); color: #7f1d1d; display: flex; flex-direction: column; gap: 3px; max-width: 300px; min-width: 170px; padding: 7px 24px 7px 10px; pointer-events: auto; position: absolute; white-space: pre-wrap; z-index: 34; }
.xlsx-grid__validation-alert strong { font-size: 11px; }
.xlsx-grid__validation-alert span { font-size: 12px; line-height: 1.45; }
.xlsx-grid__validation-alert--error { background: #fef2f2; }
.xlsx-grid__validation-alert--warning { background: #fffbeb; border-color: #fbbf24; color: #78350f; }
.xlsx-grid__validation-alert--information { background: #eff6ff; border-color: #60a5fa; color: #1e3a8a; }
.xlsx-grid__validation-dismiss { background: transparent; border: 0; color: inherit; cursor: pointer; font-size: 16px; line-height: 1; padding: 0 2px; position: absolute; right: 4px; top: 3px; }
</style>
