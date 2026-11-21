import { computed } from "vue";
import {
  addWorkbookConditionalFormatRule,
  addWorkbookTable,
  createConditionalFormatRuleFromDraft,
  deleteWorkbookConditionalFormatRule,
  deleteWorkbookTable,
  insertWorkbookChart,
  loadWorkbookChartAssets,
  normalizeWorkbookTableMetadata,
  resolvePageSetupInput,
  setWorkbookCellComment,
  setWorkbookCellHyperlink,
  setWorkbookDataValidation,
  setWorkbookFreezePanes,
  setWorkbookPageSetup,
  updateWorkbookConditionalFormatRule,
  type XlsxCellAddress,
  type XlsxCellComment,
  type XlsxCellCommentInput,
  type XlsxCellHyperlink,
  type XlsxCellHyperlinkInput,
  type XlsxCellRange,
  type XlsxChart,
  type XlsxChartInsertInput,
  type XlsxConditionalFormatDraft,
  type XlsxConditionalFormatRule,
  type XlsxDataValidation,
  type XlsxDataValidationInput,
  type XlsxFreezePanes,
  type XlsxPageMargins,
  type XlsxPageSetup,
  type XlsxPageSetupInput,
  type XlsxPrintOptions,
  type XlsxSheetData,
  type XlsxTable,
  type XlsxTableCreateInput
} from "@arcships/xlsx-core";
import type { XlsxControllerContext } from "./internal";
import { mapWorksheetTables } from "./workbook-state";

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  const startRow = Math.min(range.start.row, range.end.row);
  const startCol = Math.min(range.start.col, range.end.col);
  const endRow = Math.max(range.start.row, range.end.row);
  const endCol = Math.max(range.start.col, range.end.col);
  return {
    start: { col: startCol, row: startRow },
    end: { col: endCol, row: endRow }
  };
}

function normalizeFreezePanes(freeze: XlsxFreezePanes | null): XlsxFreezePanes | null {
  if (!freeze) {
    return null;
  }
  const row = Math.max(0, Math.floor(Number(freeze.row) || 0));
  const col = Math.max(0, Math.floor(Number(freeze.col) || 0));
  return row > 0 || col > 0 ? { row, col } : null;
}

function sameFreezePanes(left: XlsxFreezePanes | null, right: XlsxFreezePanes | null) {
  if (!left || !right) {
    return left === right;
  }
  return left.row === right.row && left.col === right.col;
}

/** Shallow equality over the modelled keys of one page-setup record. */
function sameRecord<T extends object>(left: T, right: T, keys: readonly (keyof T)[]) {
  return keys.every((key) => left[key] === right[key]);
}

const PAGE_SETUP_KEYS: readonly (keyof XlsxPageSetup)[] = [
  "paperSize",
  "orientation",
  "scale",
  "fitToWidth",
  "fitToHeight",
  "fitToPage",
  "pageOrder",
  "firstPageNumber",
  "useFirstPageNumber",
  "blackAndWhite",
  "draft",
  "cellComments",
  "errors",
];
const PAGE_MARGIN_KEYS: readonly (keyof XlsxPageMargins)[] = [
  "left",
  "right",
  "top",
  "bottom",
  "header",
  "footer",
];
const PRINT_OPTION_KEYS: readonly (keyof XlsxPrintOptions)[] = [
  "gridLines",
  "headings",
  "horizontalCentered",
  "verticalCentered",
];

function sameRangeList(left: readonly XlsxCellRange[], right: readonly XlsxCellRange[]) {
  if (left.length !== right.length) {
    return false;
  }
  return left.every((range, index) => {
    const other = right[index];
    return Boolean(other)
      && range.start.row === other.start.row
      && range.start.col === other.start.col
      && range.end.row === other.end.row
      && range.end.col === other.end.col;
  });
}

/** Turn an editor's list input into the OOXML-style formula string. */
export function resolveDataValidationListSource(source: string | undefined) {
  const trimmed = source?.trim() ?? "";
  if (!trimmed) {
    return "";
  }
  if (trimmed.startsWith("=")) {
    return trimmed.slice(1).trim();
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed;
  }
  return trimmed.includes(",") ? `"${trimmed.replace(/"/g, '""')}"` : trimmed;
}

export function createAuthoringDomain(ctx: XlsxControllerContext) {
  function getAssets() {
    return ctx.imageAssetsRef.value;
  }

  function activeWorkbookSheetIndex() {
    return ctx.activeSheet.value?.workbookSheetIndex ?? -1;
  }

  /**
   * Authoring writes worksheet parts directly, so it needs the unzipped source
   * archive (main-thread loads only) and an unlocked sheet.
   */
  function canAuthor() {
    const assets = getAssets();
    return Boolean(
      assets
        && assets.archive
        && assets.sheetPathsByWorkbookSheetIndex.length > 0
        && !ctx.isWorkerBacked.value
    ) && !ctx.isEditingLocked.value;
  }

  function patchActiveSheet(patch: Partial<XlsxSheetData>) {
    const workbookSheetIndex = activeWorkbookSheetIndex();
    if (workbookSheetIndex < 0) {
      return;
    }
    // Keep the parsed sheet state in sync with the archive. Later refreshes
    // (cell edits, recalculation, sorting, resizing) rebuild the sheet model
    // from this state, so an authored object would otherwise disappear from the
    // view even though it is still in the workbook.
    const assets = getAssets();
    const sheetStates = assets?.sheetStatesByWorkbookSheetIndex;
    const currentState = sheetStates?.[workbookSheetIndex];
    if (sheetStates && currentState) {
      sheetStates[workbookSheetIndex] = { ...currentState, ...patch };
    }
    ctx.sheets.value = ctx.sheets.value.map((sheet) =>
      sheet.workbookSheetIndex === workbookSheetIndex ? { ...sheet, ...patch } : sheet
    );
    ctx.revision.value = ctx.revision.value + 1;
  }

  // ── Freeze panes ───────────────────────────────────────────────────
  const freezePanes = computed<XlsxFreezePanes | null>(() => ctx.activeSheet.value?.freezePanes ?? null);

  function setFreezePanes(freeze: XlsxFreezePanes | null) {
    if (!canAuthor()) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const next = normalizeFreezePanes(freeze);
    if (workbookSheetIndex < 0 || sameFreezePanes(ctx.activeSheet.value?.freezePanes ?? null, next)) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    if (!setWorkbookFreezePanes(getAssets()!, workbookSheetIndex, next)) {
      return false;
    }
    patchActiveSheet({ freezePanes: next });
    return true;
  }

  function toggleFreezePanesAtActiveCell() {
    const cell = ctx.activeCell.value;
    if (!cell) {
      return false;
    }
    const current = ctx.activeSheet.value?.freezePanes ?? null;
    if (current && current.row === cell.row && current.col === cell.col) {
      return setFreezePanes(null);
    }
    return setFreezePanes({ row: cell.row, col: cell.col });
  }

  // ── Page set-up ─────────────────────────────────────────────────────
  /**
   * Writes the page setup dialog's patch into the worksheet part.
   *
   * `resolvePageSetupInput` normalizes the patch once, so the model fields the
   * viewer renders and the XML an export carries cannot disagree.
   */
  function setPageSetup(input: XlsxPageSetupInput) {
    if (!canAuthor()) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const sheet = ctx.activeSheet.value;
    const assets = getAssets();
    if (!sheet || workbookSheetIndex < 0 || !assets?.sheetPathsByWorkbookSheetIndex[workbookSheetIndex]) {
      return false;
    }
    const resolved = resolvePageSetupInput(sheet, input);
    const printAreas = resolved.printAreas ?? sheet.printAreas;
    const unchanged = sameRecord(resolved.pageSetup, sheet.pageSetup, PAGE_SETUP_KEYS)
      && sameRecord(resolved.pageMargins, sheet.pageMargins, PAGE_MARGIN_KEYS)
      && sameRecord(resolved.printOptions, sheet.printOptions, PRINT_OPTION_KEYS)
      && sameRangeList(printAreas, sheet.printAreas);
    if (unchanged) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    if (!setWorkbookPageSetup(assets, workbookSheetIndex, input)) {
      return false;
    }
    patchActiveSheet({
      pageMargins: resolved.pageMargins,
      pageSetup: resolved.pageSetup,
      printAreas,
      printOptions: resolved.printOptions,
    });
    return true;
  }

  // ── Comments ───────────────────────────────────────────────────────
  function getCellComment(cell?: XlsxCellAddress | null) {
    const target = cell ?? ctx.activeCell.value;
    if (!target) {
      return null;
    }
    return ctx.activeSheet.value?.comments.find(
      (comment) => comment.cell.row === target.row && comment.cell.col === target.col
    ) ?? null;
  }

  function setCellComment(cell: XlsxCellAddress, input: XlsxCellCommentInput | null) {
    if (!canAuthor()) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const text = input?.text?.trim() ?? "";
    const author = input?.author?.trim() || undefined;
    const existing = getCellComment(cell);
    if (input && !text) {
      return false;
    }
    if (!input && !existing) {
      return false;
    }
    if (input && existing && existing.text === text && (existing.author ?? "") === (author ?? "")) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    const comments = setWorkbookCellComment(
      getAssets()!,
      workbookSheetIndex,
      cell,
      input ? { author, text } : null
    );
    if (!comments) {
      return false;
    }
    patchActiveSheet({ comments });
    return true;
  }

  // ── Hyperlinks ─────────────────────────────────────────────────────
  function getCellHyperlink(cell?: XlsxCellAddress | null) {
    const target = cell ?? ctx.activeCell.value;
    if (!target) {
      return null;
    }
    return ctx.activeSheet.value?.hyperlinks.find(
      (hyperlink) => hyperlink.cell.row === target.row && hyperlink.cell.col === target.col
    ) ?? null;
  }

  function setCellHyperlink(cell: XlsxCellAddress, input: XlsxCellHyperlinkInput | null) {
    if (!canAuthor()) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const target = input?.target.trim() ?? "";
    const tooltip = input?.tooltip?.trim() || undefined;
    const existing = getCellHyperlink(cell);
    if (input && !target) {
      return false;
    }
    if (!input && !existing) {
      return false;
    }
    if (
      input
      && existing
      && existing.target === target
      && (existing.tooltip ?? "") === (tooltip ?? "")
    ) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    const hyperlinks = setWorkbookCellHyperlink(
      getAssets()!,
      workbookSheetIndex,
      cell,
      input ? { target, tooltip } : null
    );
    if (!hyperlinks) {
      return false;
    }
    patchActiveSheet({ hyperlinks });
    return true;
  }

  // ── Conditional formatting ─────────────────────────────────────────
  function getConditionalFormatRules() {
    return ctx.activeSheet.value?.conditionalFormatRules ?? [];
  }

  function resolveDraftRanges(draft: XlsxConditionalFormatDraft, fallback?: XlsxCellRange[]) {
    if (draft.ranges?.length) {
      return draft.ranges.map(normalizeRange);
    }
    if (fallback?.length) {
      return fallback.map(normalizeRange);
    }
    return ctx.selection.value ? [normalizeRange(ctx.selection.value)] : [];
  }

  function addConditionalFormatRule(draft: XlsxConditionalFormatDraft) {
    if (!canAuthor()) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const ranges = resolveDraftRanges(draft);
    const rules = getConditionalFormatRules();
    const rule = createConditionalFormatRuleFromDraft(draft, ranges, rules.length + 1);
    if (!rule || workbookSheetIndex < 0) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    const nextRules = addWorkbookConditionalFormatRule(getAssets()!, workbookSheetIndex, rule);
    if (!nextRules) {
      return false;
    }
    patchActiveSheet({ conditionalFormatRules: nextRules });
    return true;
  }

  function updateConditionalFormatRule(index: number, draft: XlsxConditionalFormatDraft) {
    if (!canAuthor()) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const existing = getConditionalFormatRules()[index];
    if (!existing || workbookSheetIndex < 0) {
      return false;
    }
    const ranges = resolveDraftRanges(draft, existing.ranges);
    const rule = createConditionalFormatRuleFromDraft(draft, ranges, existing.priority);
    if (!rule) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    const nextRules = updateWorkbookConditionalFormatRule(
      getAssets()!,
      workbookSheetIndex,
      index,
      rule
    );
    if (!nextRules) {
      return false;
    }
    patchActiveSheet({ conditionalFormatRules: nextRules });
    return true;
  }

  function deleteConditionalFormatRule(index: number) {
    if (!canAuthor()) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    if (!getConditionalFormatRules()[index] || workbookSheetIndex < 0) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    const nextRules = deleteWorkbookConditionalFormatRule(getAssets()!, workbookSheetIndex, index);
    if (!nextRules) {
      return false;
    }
    patchActiveSheet({ conditionalFormatRules: nextRules });
    return true;
  }

  // ── Tables ─────────────────────────────────────────────────────────
  function worksheetHeaderNames(range: XlsxCellRange) {
    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet) {
      return [];
    }
    // The controller owns this worksheet handle; do not free it here.
    return Array.from(
      { length: range.end.col - range.start.col + 1 },
      (_, index) => {
        const column = range.start.col + index;
        const formatted = worksheet.getFormattedValueAt(range.start.row, column);
        if (formatted && formatted.trim()) {
          return formatted.trim();
        }
        const cell = worksheet.getCellAt(range.start.row, column);
        try {
          const value = cell.toJs();
          return value === null || value === undefined ? "" : String(value).trim();
        } finally {
          cell.free();
        }
      }
    );
  }

  function createTable(input?: XlsxTableCreateInput | null) {
    if (!canAuthor()) {
      return null;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const source = input?.range ?? ctx.selection.value;
    if (!source || workbookSheetIndex < 0) {
      return null;
    }
    const range = normalizeRange(source);
    const hasHeaders = input?.hasHeaders !== false;
    const columnNames = hasHeaders ? worksheetHeaderNames(range) : [];
    ctx.recordHistoryBeforeMutation();
    const result = addWorkbookTable(getAssets()!, workbookSheetIndex, {
      columnNames,
      hasHeaders,
      name: input?.name,
      range,
      styleName: input?.styleName
    });
    if (!result) {
      return null;
    }
    ctx.revision.value = ctx.revision.value + 1;
    return result.table;
  }

  function getSheetTables(): XlsxTable[] {
    const workbookSheetIndex = activeWorkbookSheetIndex();
    if (workbookSheetIndex < 0) {
      return [];
    }
    const parsed = getAssets()?.tableMetadataByWorkbookSheetIndex?.[workbookSheetIndex];
    if (parsed?.length) {
      return normalizeWorkbookTableMetadata(parsed);
    }
    return mapWorksheetTables(ctx.getActiveWorksheet());
  }

  function removeTable(name: string) {
    if (!canAuthor() || !name) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    if (workbookSheetIndex < 0) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    const tables = deleteWorkbookTable(getAssets()!, workbookSheetIndex, name);
    if (!tables) {
      return false;
    }
    ctx.revision.value = ctx.revision.value + 1;
    return true;
  }

  function getTableAtCell(cell?: XlsxCellAddress | null) {
    const target = cell ?? ctx.activeCell.value;
    if (!target) {
      return null;
    }
    return getSheetTables().find((table) => {
      const startRow = Math.min(table.start.row, table.end.row);
      const endRow = Math.max(table.start.row, table.end.row);
      const startCol = Math.min(table.start.col, table.end.col);
      const endCol = Math.max(table.start.col, table.end.col);
      return target.row >= startRow
        && target.row <= endRow
        && target.col >= startCol
        && target.col <= endCol;
    }) ?? null;
  }

  // ── Data validation ────────────────────────────────────────────────
  function buildDataValidation(input: XlsxDataValidationInput): XlsxDataValidation {
    const listSource = input.validationType === "list"
      ? resolveDataValidationListSource(input.listSource)
      : input.listSource?.trim();
    return {
      allowBlank: input.allowBlank ?? true,
      errorMessage: input.errorMessage?.trim() || undefined,
      errorStyle: "stop",
      inputMessage: input.inputMessage?.trim() || undefined,
      listSource: listSource || undefined,
      operator: input.operator,
      ranges: (input.ranges ?? []).map(normalizeRange),
      showDropdown: input.validationType === "list" ? input.showDropdown ?? true : undefined,
      showErrorAlert: true,
      showInputMessage: Boolean(input.inputMessage?.trim()),
      validationType: input.validationType,
      value1: input.validationType === "list" ? undefined : input.value1?.trim() || undefined,
      value2: input.validationType === "list" ? undefined : input.value2?.trim() || undefined
    };
  }

  function setSelectionDataValidation(input: XlsxDataValidationInput | null) {
    if (!canAuthor()) {
      return false;
    }
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const ranges = (input?.ranges?.length ? input.ranges : ctx.selection.value ? [ctx.selection.value] : [])
      .map(normalizeRange);
    if (workbookSheetIndex < 0 || ranges.length === 0) {
      return false;
    }
    const validation = input ? buildDataValidation({ ...input, ranges }) : null;
    if (validation && !validation.listSource && !validation.value1 && !validation.value2) {
      return false;
    }
    ctx.recordHistoryBeforeMutation();
    const dataValidations = setWorkbookDataValidation(
      getAssets()!,
      workbookSheetIndex,
      ranges,
      validation
    );
    if (!dataValidations) {
      return false;
    }
    patchActiveSheet({ dataValidations });
    return true;
  }

  // ── Charts ─────────────────────────────────────────────────────────
  function resolveChartRange(input?: XlsxChartInsertInput | null) {
    if (input?.range) {
      return normalizeRange(input.range);
    }
    const selection = ctx.selection.value ? normalizeRange(ctx.selection.value) : null;
    if (
      selection
      && (selection.start.row !== selection.end.row || selection.start.col !== selection.end.col)
    ) {
      return selection;
    }
    const sheet = ctx.activeSheet.value;
    if (!sheet || sheet.maxUsedRow < sheet.minUsedRow || sheet.maxUsedCol < sheet.minUsedCol) {
      return null;
    }
    return {
      start: { col: sheet.minUsedCol, row: sheet.minUsedRow },
      end: { col: Math.min(sheet.maxUsedCol, sheet.minUsedCol + 12), row: Math.min(sheet.maxUsedRow, sheet.minUsedRow + 50) }
    };
  }

  function insertChart(input?: XlsxChartInsertInput | null) {
    if (!canAuthor() || !ctx.workbook.value || !ctx.activeSheet.value) {
      return null;
    }
    const assets = getAssets();
    const workbookSheetIndex = activeWorkbookSheetIndex();
    const range = resolveChartRange(input);
    if (!assets || workbookSheetIndex < 0 || !range) {
      return null;
    }
    ctx.recordHistoryBeforeMutation();
    const result = insertWorkbookChart(assets, workbookSheetIndex, {
      family: input?.family ?? "bar",
      hasHeaderColumn: input?.hasHeaderColumn ?? true,
      hasHeaderRow: input?.hasHeaderRow ?? true,
      range,
      sheetName: ctx.activeSheet.value.name,
      title: input?.title
    });
    if (!result) {
      return null;
    }
    ctx.sheetOriginsRef.value = assets.sheetOrigins.slice();
    const nextChartAssets = loadWorkbookChartAssets(
      ctx.workbook.value,
      assets,
      ctx.visibleSheetIndexByWorkbookSheetIndex.value,
      ctx.showHiddenSheets
    );
    ctx.setChartAssets(nextChartAssets);
    const inserted: XlsxChart | null = nextChartAssets.chartsByWorkbookSheetIndex[workbookSheetIndex]
      ?.find((chart) => chart.chartPath === result.chartPath) ?? null;
    ctx.revision.value = ctx.revision.value + 1;
    return inserted;
  }

  return {
    addConditionalFormatRule,
    canAuthor,
    createTable,
    deleteConditionalFormatRule,
    freezePanes,
    getCellComment,
    getCellHyperlink,
    getConditionalFormatRules,
    getSheetTables,
    getTableAtCell,
    insertChart,
    removeTable,
    setCellComment,
    setCellHyperlink,
    setFreezePanes,
    setPageSetup,
    setSelectionDataValidation,
    toggleFreezePanesAtActiveCell,
    updateConditionalFormatRule
  };
}

export type XlsxAuthoringDomain = ReturnType<typeof createAuthoringDomain>;
