import { shallowRef, watch } from "vue";
import { XLSX_MIME_TYPE, CSV_MIME_TYPE, INTERNAL_CLIPBOARD_MIME } from "./internal";
import type { ClipboardPayload, ClipboardMatrixCell, ClipboardMerge } from "./internal";
import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxClipboardCutState,
  XlsxClipboardData,
  XlsxPasteSpecialOptions
} from "@arcships/xlsx-core";
import { resolveWorkbookColor, resolveWorkbookFillStyle } from "@arcships/xlsx-core";
import { cellAddressToA1, normalizeRange, rangeContainsCell, rangeToA1 } from "./selection";
import { isWithinSheetBounds, SHEET_MAX_COLUMNS, SHEET_MAX_ROWS, shiftFormulaReferences } from "./formula-references";
import { decodeHtmlEntities, escapeHtml, mapBorder, resolveEffectiveCellNumberFormat, resolveInheritedCellStyle } from "./formatting";
import { DEFAULT_CELL_STYLE_PATCH, localizeFormattedNumber, resolveCellInput } from "./cell-input";
import type { XlsxControllerContext } from "./internal";
import { XlsxSourceError } from "@arcships/xlsx-core";
import { officeMessagesEn } from "@arcships/office-interaction";
import {
  parseClipboardText,
  serializeClipboardText,
  writeClipboardPayload
} from "./clipboard-text";
import type { ClipboardWriteTargets } from "./clipboard-text";

export { INTERNAL_CLIPBOARD_MIME, escapeHtml, parseClipboardText };
export type { ClipboardPayload, ClipboardMatrixCell, ClipboardMerge };

/** Resolves the clipboard handles of the current browser; absent ones disable that flavour. */
function resolveClipboardWriteTargets(mimeType: string): ClipboardWriteTargets {
  return {
    mimeType,
    clipboard: typeof navigator === "undefined" ? undefined : navigator.clipboard,
    ClipboardItem: typeof ClipboardItem === "function" ? ClipboardItem : undefined
  };
}

function downloadBlob(blob: Blob, fileName: string) {
  const ownerDocument = globalThis.document;
  if (!ownerDocument?.body || typeof ownerDocument.createElement !== "function") {
    throw new Error(officeMessagesEn.xlsx.clipboard.noDocument);
  }
  const anchor = ownerDocument.createElement("a");
  let objectUrl: string | undefined;
  let clicked = false;
  try {
    objectUrl = URL.createObjectURL(blob);
    anchor.href = objectUrl;
    anchor.download = fileName;
    ownerDocument.body.append(anchor);
    anchor.click();
    clicked = true;
  } finally {
    try { anchor.remove(); } catch { /* URL cleanup must continue */ }
    if (objectUrl) {
      if (clicked) {
        try {
          setTimeout(() => URL.revokeObjectURL(objectUrl as string), 0);
        } catch {
          URL.revokeObjectURL(objectUrl);
        }
      } else {
        URL.revokeObjectURL(objectUrl);
      }
    }
  }
}

export function downloadArrayBuffer(file: ArrayBuffer, fileName: string) {
  downloadBlob(new Blob([file], { type: XLSX_MIME_TYPE }), fileName);
}

export function downloadBytes(bytes: Uint8Array, fileName: string, mimeType: string) {
  const normalizedBytes = new Uint8Array(bytes.byteLength);
  normalizedBytes.set(bytes);
  downloadBlob(new Blob([normalizedBytes], { type: mimeType }), fileName);
}

export function downloadText(text: string, fileName: string, mimeType: string) {
  downloadBlob(new Blob([text], { type: mimeType }), fileName);
}

export function downloadUrl(src: string, fileName: string) {
  let url: URL;
  try {
    url = new URL(src);
  } catch {
    throw new XlsxSourceError({
      code: "SOURCE_NOT_ALLOWED",
      message: officeMessagesEn.xlsx.clipboard.invalidUrl,
      sourceKind: "url",
    });
  }
  if (url.protocol !== "blob:") {
    throw new XlsxSourceError({
      code: "SOURCE_NOT_ALLOWED",
      message: officeMessagesEn.xlsx.clipboard.unverifiedBlob,
      sourceKind: "url",
      url: url.href,
    });
  }
  const ownerDocument = globalThis.document;
  if (!ownerDocument?.body || typeof ownerDocument.createElement !== "function") {
    throw new Error(officeMessagesEn.xlsx.clipboard.noDocument);
  }
  const anchor = ownerDocument.createElement("a");
  try {
    anchor.href = url.href;
    anchor.download = fileName;
    anchor.rel = "noreferrer";
    ownerDocument.body.append(anchor);
    anchor.click();
  } finally {
    anchor.remove();
  }
}

export function createClipboardDomain(ctx: XlsxControllerContext) {
  // Only re-render values that can be a localized number; dates, text and
  // errors skip the extra style lookup entirely.
  const localizableDisplayValue = /^[\s($€£¥₹₽₪¢₫₴₦₱₸₺฿]*[+-]?\d/;
  // The last copy/cut made inside the viewer. Paste Special uses it so the
  // options apply even where the system clipboard cannot be read.
  let storedClipboard: {
    cutRange: XlsxCellRange | null;
    payload: ClipboardPayload;
    sheetIndex: number;
  } | null = null;
  // The cut that is still waiting for its paste. It outlives the payload store
  // so hosts can render a marquee, and it is the state Escape cancels.
  const cutState = shallowRef<XlsxClipboardCutState | null>(null);

  function resetClipboardState() {
    storedClipboard = null;
    cutState.value = null;
  }

  // A restarted workbook (new file, sheet rebuild) invalidates both the payload
  // and the source range: pasting it would write into another document and
  // "clear" cells that were never cut.
  watch(() => ctx.workbook.value, () => { resetClipboardState(); });

  function localizeDisplayValue(value: string, resolveFormat: () => unknown) {
    if (!localizableDisplayValue.test(value)) {
      return value;
    }
    return localizeFormattedNumber(value, resolveFormat(), ctx.locale.value);
  }

  function getCellDisplayValue(cell?: XlsxCellAddress | null) {
    if (cell && ctx.activeSheet.value) {
      const workerSnapshot = ctx.workerCellSnapshotCacheRef.value.get(`${ctx.activeSheet.value.workbookSheetIndex}:${cell.row}:${cell.col}`);
      if (workerSnapshot) {
        // Worker snapshots carry the resolved number format so the formula bar
        // and a11y labels can localize exactly like the main thread does.
        return localizeDisplayValue(workerSnapshot.displayValue, () => workerSnapshot.numberFormat);
      }
    }

    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    const formula = worksheet.getFormulaAt(cell.row, cell.col);
    const cachedFormulaValue = formula ? ctx.activeSheet.value?.cachedFormulaValues?.[cellAddressToA1(cell)] : undefined;
    const formatted = worksheet.getFormattedValueAt(cell.row, cell.col);
    if (formatted && !(formula && cachedFormulaValue !== undefined && formatted.startsWith("#"))) {
      return localizeDisplayValue(
        decodeHtmlEntities(formatted),
        () => resolveEffectiveCellNumberFormat(
          worksheet.getCellStyleAt(cell.row, cell.col),
          ctx.activeSheet.value,
          cell.row,
          cell.col
        )
      );
    }

    const calculated = worksheet.getCalculatedValueAt(cell.row, cell.col);
    try {
      if (formula && cachedFormulaValue !== undefined && calculated.is_error) {
        return cachedFormulaValue;
      }
      if (calculated.is_error) {
        return calculated.asError() ?? "";
      }
      if (calculated.is_empty) {
        return "";
      }

      return localizeDisplayValue(
        calculated.toString(),
        () => resolveEffectiveCellNumberFormat(
          worksheet.getCellStyleAt(cell.row, cell.col),
          ctx.activeSheet.value,
          cell.row,
          cell.col
        )
      );
    } finally {
      calculated.free();
    }
  }

  function getCellFormula(cell?: XlsxCellAddress | null) {
    if (cell && ctx.activeSheet.value) {
      const workerSnapshot = ctx.workerCellSnapshotCacheRef.value.get(`${ctx.activeSheet.value.workbookSheetIndex}:${cell.row}:${cell.col}`);
      if (workerSnapshot) {
        return workerSnapshot.formula;
      }
    }

    const worksheet = ctx.getActiveWorksheet();
    if (!worksheet || !cell) {
      return "";
    }

    return worksheet.getFormulaAt(cell.row, cell.col) ?? "";
  }

  function getClipboardData(): XlsxClipboardData | null {
    const worksheet = ctx.getActiveWorksheet();
    const targetRange = ctx.selection.value ?? (ctx.activeCell.value ? { start: ctx.activeCell.value, end: ctx.activeCell.value } : null);
    if (!worksheet || !targetRange) {
      return null;
    }

    const normalized = normalizeRange(targetRange);
    const rows: string[][] = [];
    const htmlRows: string[] = [];
    const payload: ClipboardPayload = {
      cells: [],
      cols: normalized.end.col - normalized.start.col + 1,
      merges: [],
      origin: { col: normalized.start.col, row: normalized.start.row },
      rows: normalized.end.row - normalized.start.row + 1
    };

    for (let row = normalized.start.row; row <= normalized.end.row; row += 1) {
      const textCells: string[] = [];
      const htmlCells: string[] = [];

      for (let col = normalized.start.col; col <= normalized.end.col; col += 1) {
        if (worksheet.isMergedSecondary(row, col)) {
          textCells.push("");
          continue;
        }

        const formula = worksheet.getFormulaAt(row, col) ?? null;
        const value = getCellDisplayValue({ row, col });
        const rawCellValue = worksheet.getCellAt(row, col);
        let rawValue: unknown = "";
        try {
          rawValue = rawCellValue.toJs();
        } finally {
          rawCellValue.free();
        }
        const merge = worksheet.getMergeSpan(row, col) as
          | { colSpan?: number; rowSpan?: number }
          | null
          | undefined;
        const explicitStyle = worksheet.getCellStyleAt(row, col);
        const inheritedStyle = explicitStyle
          ? null
          : resolveInheritedCellStyle(ctx.activeSheet.value, row, col);
        // Carry the style so an internal paste keeps the number format. Cells
        // without an explicit style inherit row/column formatting, so only the
        // inherited number format is materialised instead of freezing the
        // whole inherited style onto the pasted cell.
        const carriedStyle = explicitStyle
          ?? (inheritedStyle?.numberFormat ? { numberFormat: inheritedStyle.numberFormat } : null);
        const rawStyle = (explicitStyle as Record<string, unknown> | null | undefined) ?? inheritedStyle;
        const cellStyles: string[] = [
          "padding:2px 4px",
          "white-space:pre-wrap",
          "vertical-align:top"
        ];

        const fill = rawStyle?.fill as Record<string, unknown> | undefined;
        if (fill) {
          const fillStyle = resolveWorkbookFillStyle(fill, ctx.activeSheet.value?.themePalette);
          if (fillStyle.backgroundColor && fillStyle.backgroundColor.toLowerCase() !== "#ffffff") {
            cellStyles.push(`background-color:${fillStyle.backgroundColor}`);
          }
          if (fillStyle.backgroundImage) {
            cellStyles.push(`background-image:${fillStyle.backgroundImage}`);
          }
        }

        const font = rawStyle?.font as Record<string, unknown> | undefined;
        if (font) {
          if (font.bold) {
            cellStyles.push("font-weight:700");
          }
          if (font.italic) {
            cellStyles.push("font-style:italic");
          }
          if (font.underline && font.underline !== "none") {
            cellStyles.push("text-decoration:underline");
          }
          if (font.strikethrough) {
            cellStyles.push("text-decoration:line-through");
          }
          const fontColor = resolveWorkbookColor(font.color as Record<string, unknown> | undefined, ctx.activeSheet.value?.themePalette);
          if (fontColor) {
            cellStyles.push(`color:${fontColor}`);
          }
          if (typeof font.size === "number") {
            cellStyles.push(`font-size:${font.size}pt`);
          }
        }

        const alignment = rawStyle?.alignment as Record<string, unknown> | undefined;
        if (alignment?.horizontal && alignment.horizontal !== "general") {
          cellStyles.push(`text-align:${String(alignment.horizontal)}`);
        }
        if (alignment?.wrapText) {
          cellStyles.push("white-space:pre-wrap");
          cellStyles.push("word-break:break-word");
        }

        const border = rawStyle?.border as Record<string, Record<string, unknown>> | undefined;
        if (border?.top?.style && border.top.style !== "none") {
          cellStyles.push(`border-top:${mapBorder(border.top as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.right?.style && border.right.style !== "none") {
          cellStyles.push(`border-right:${mapBorder(border.right as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.bottom?.style && border.bottom.style !== "none") {
          cellStyles.push(`border-bottom:${mapBorder(border.bottom as { color?: { hex?: string }; style: string })}`);
        }
        if (border?.left?.style && border.left.style !== "none") {
          cellStyles.push(`border-left:${mapBorder(border.left as { color?: { hex?: string }; style: string })}`);
        }

        const rowSpan = Math.min(merge?.rowSpan ?? 1, normalized.end.row - row + 1);
        const colSpan = Math.min(merge?.colSpan ?? 1, normalized.end.col - col + 1);

        payload.cells.push({
          colOffset: col - normalized.start.col,
          formula,
          rawValue: rawValue === undefined ? "" : rawValue,
          rowOffset: row - normalized.start.row,
          style: carriedStyle,
          value
        });

        if (rowSpan > 1 || colSpan > 1) {
          payload.merges.push({
            colOffset: col - normalized.start.col,
            colSpan,
            rowOffset: row - normalized.start.row,
            rowSpan
          });
        }

        textCells.push(value);
        htmlCells.push(
          `<td${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ""}${colSpan > 1 ? ` colspan="${colSpan}"` : ""} style="${escapeHtml(cellStyles.join(";"))}">${escapeHtml(value)}</td>`
        );
      }

      rows.push(textCells);
      htmlRows.push(`<tr>${htmlCells.join("")}</tr>`);
    }

    return {
      html: `<table style="border-collapse:collapse">${htmlRows.join("")}</table>`,
      structured: JSON.stringify(payload),
      text: serializeClipboardText(rows)
    };
  }

  // Set while a cut-paste is applying its own snapshot, so the inner paste
  // does not add a second history entry that undo would stop on.
  let pasteHistorySuppressed = false;

  function storeSelectionForPaste(cut: boolean) {
    const data = getClipboardData();
    if (!data) {
      return null;
    }
    let payload: ClipboardPayload;
    try {
      payload = JSON.parse(data.structured) as ClipboardPayload;
    } catch {
      return null;
    }
    const range = ctx.selection.value
      ?? (ctx.activeCell.value ? { end: ctx.activeCell.value, start: ctx.activeCell.value } : null);
    const sheetIndex = ctx.activeSheet.value?.workbookSheetIndex ?? -1;
    const cutRange = cut && range && sheetIndex >= 0 ? normalizeRange(range) : null;
    storedClipboard = {
      cutRange,
      payload,
      sheetIndex
    };
    // A plain copy always ends a pending cut; a cut replaces it and starts a
    // new marquee on the current sheet.
    cutState.value = cutRange && sheetIndex >= 0
      ? { range: cutRange, sheetIndex }
      : null;
    return data;
  }

  function transformPayloadForPaste(
    payload: ClipboardPayload,
    options: XlsxPasteSpecialOptions,
  ): ClipboardPayload {
    const transpose = options.transpose === true;
    const mode = options.mode ?? "all";
    const cells = payload.cells.map((cell) => {
      const positioned = transpose
        ? { ...cell, colOffset: cell.rowOffset, rowOffset: cell.colOffset }
        : { ...cell };
      if (mode === "all") {
        return positioned;
      }
      if (mode === "values") {
        // Values-only writes the formula result, never the formula itself or
        // the source formatting.
        const rawValue = cell.formula
          ? (cell.rawValue !== undefined && cell.rawValue !== "" ? cell.rawValue : cell.value)
          : cell.rawValue;
        return {
          ...positioned,
          formula: null,
          rawValue,
          style: undefined
        };
      }
      return positioned;
    });
    const merges = mode !== "all"
      ? []
      : transpose
        ? (payload.merges ?? []).map((merge) => ({
            colOffset: merge.rowOffset,
            colSpan: merge.rowSpan,
            rowOffset: merge.colOffset,
            rowSpan: merge.colSpan
          }))
        : payload.merges;
    return {
      ...payload,
      cells,
      cols: transpose ? Math.max(1, payload.rows) : Math.max(1, payload.cols),
      merges,
      rows: transpose ? Math.max(1, payload.cols) : Math.max(1, payload.rows)
    };
  }

  function applyFormatsOnly(payload: ClipboardPayload, options: XlsxPasteSpecialOptions) {
    const worksheet = ctx.getActiveWorksheet();
    const targetCell = ctx.activeCell.value ?? ctx.selection.value?.start ?? null;
    if (ctx.isEditingLocked.value || !worksheet || !ctx.workbook.value || !targetCell) {
      return false;
    }
    const transpose = options.transpose === true;
    const beforeSnapshot = {
      activeCell: ctx.activeCell.value,
      selection: ctx.selection.value
    };
    const mutations: import("./internal").RangeCellMutation[] = [];
    for (const cell of payload.cells) {
      if (cell.style === undefined) {
        continue;
      }
      const nextCell = transpose
        ? { col: targetCell.col + cell.rowOffset, row: targetCell.row + cell.colOffset }
        : { col: targetCell.col + cell.colOffset, row: targetCell.row + cell.rowOffset };
      if (!isWithinSheetBounds(nextCell) || worksheet.isMergedSecondary(nextCell.row, nextCell.col)) {
        continue;
      }
      const before = ctx.captureCellMutationState(nextCell);
      if (!before) {
        continue;
      }
      worksheet.setCellStyleAt(
        nextCell.row,
        nextCell.col,
        cell.style && typeof cell.style === "object" ? cell.style : DEFAULT_CELL_STYLE_PATCH
      );
      const after = ctx.captureCellMutationState(nextCell);
      if (!after) {
        continue;
      }
      mutations.push({ after, before, cell: nextCell });
    }
    if (mutations.length === 0) {
      return false;
    }
    ctx.refreshWorkbookState(ctx.workbook.value);
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: Math.min(targetCell.col + Math.max((transpose ? payload.rows : payload.cols) - 1, 0), SHEET_MAX_COLUMNS - 1),
        row: Math.min(targetCell.row + Math.max((transpose ? payload.cols : payload.rows) - 1, 0), SHEET_MAX_ROWS - 1)
      }
    });
    ctx.activeCell.value = targetCell;
    ctx.selection.value = nextRange;
    ctx.selectionAnchorRef.value = targetCell;
    ctx.recordRangeEditHistory(mutations, nextRange, targetCell, beforeSnapshot);
    return true;
  }

  function clearCutSource(targetRange: XlsxCellRange | null) {
    const stored = storedClipboard;
    if (!stored?.cutRange) {
      return;
    }
    try {
      const workbook = ctx.workbook.value;
      if (!workbook || stored.sheetIndex < 0) {
        return;
      }
      const worksheet = workbook.getSheet(stored.sheetIndex);
      try {
        for (let row = stored.cutRange.start.row; row <= stored.cutRange.end.row; row += 1) {
          for (let col = stored.cutRange.start.col; col <= stored.cutRange.end.col; col += 1) {
            if (targetRange && rangeContainsCell(targetRange, { row, col })) {
              continue;
            }
            if (worksheet.isMergedSecondary(row, col)) {
              continue;
            }
            worksheet.setCell(cellAddressToA1({ row, col }), "");
            if (worksheet.getCellStyleAt(row, col)) {
              worksheet.setCellStyleAt(row, col, DEFAULT_CELL_STYLE_PATCH);
            }
          }
        }
      } finally {
        worksheet.free();
      }
      if (ctx.workbook.value) {
        ctx.refreshWorkbookState(ctx.workbook.value);
      }
    } finally {
      // The move happened: neither the payload nor the marquee may survive it.
      resetClipboardState();
    }
  }

  function sameClipboardPayload(left: ClipboardPayload, right: ClipboardPayload) {
    if (left.cols !== right.cols || left.rows !== right.rows || left.cells.length !== right.cells.length) {
      return false;
    }
    if (!left.origin || !right.origin) {
      return left.origin === right.origin;
    }
    return left.origin.row === right.origin.row && left.origin.col === right.origin.col;
  }

  /**
   * Applies Excel's paste options to a structured clipboard payload.
   *
   * @returns `true` when the workbook changed.
   */
  function pasteSpecialStructuredData(
    serializedPayload: string,
    options: XlsxPasteSpecialOptions = {},
  ) {
    let payload: ClipboardPayload;
    try {
      payload = JSON.parse(serializedPayload) as ClipboardPayload;
    } catch {
      return false;
    }
    if (!Array.isArray(payload.cells) || payload.cells.length === 0) {
      return false;
    }
    if ((options.mode ?? "all") === "formats") {
      const applied = applyFormatsOnly(payload, options);
      // A formats-only paste consumes the pending cut like a copy would; the
      // source keeps its values, so the move never happens.
      if (applied && storedClipboard) {
        storedClipboard = { ...storedClipboard, cutRange: null };
        cutState.value = null;
      }
      return applied;
    }
    const stored = storedClipboard;
    const isCut = Boolean(
      stored?.cutRange && sameClipboardPayload(stored.payload, payload)
    );
    if (isCut) {
      ctx.recordHistoryBeforeMutation();
      pasteHistorySuppressed = true;
    }
    try {
      const applied = pasteStructuredClipboardData(
        JSON.stringify(transformPayloadForPaste(payload, options))
      );
      if (applied && isCut) {
        clearCutSource(ctx.selection.value);
      }
      return applied;
    } finally {
      pasteHistorySuppressed = false;
    }
  }

  async function pasteSpecialFromClipboard(options: XlsxPasteSpecialOptions = {}) {
    if (ctx.isEditingLocked.value) {
      return false;
    }
    if (storedClipboard) {
      return pasteSpecialStructuredData(JSON.stringify(storedClipboard.payload), options);
    }
    if (typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
    }
    if (navigator.clipboard.read) {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes(INTERNAL_CLIPBOARD_MIME)) {
          const blob = await item.getType(INTERNAL_CLIPBOARD_MIME);
          return pasteSpecialStructuredData(await blob.text(), options);
        }
      }
    }
    const text = await navigator.clipboard.readText();
    // Plain text has no formatting or formulas, so only a values/all paste can
    // consume it; a formats-only paste stays a no-op.
    if ((options.mode ?? "all") === "formats") {
      return false;
    }
    return pasteText(text);
  }

  function pasteText(text: string) {
    const worksheet = ctx.getActiveWorksheet();
    const targetCell = ctx.activeCell.value ?? ctx.selection.value?.start ?? null;
    if (ctx.isEditingLocked.value || !worksheet || !ctx.workbook.value || !targetCell || !text) {
      return false;
    }

    const grid = parseClipboardText(text);
    if (grid.length === 0 || grid.every((row) => row.length === 0)) {
      return false;
    }

    const beforeSnapshot = {
      activeCell: ctx.activeCell.value,
      selection: ctx.selection.value
    };
    ctx.setValidationFeedback(null);
    const mutations: import("./internal").RangeCellMutation[] = [];
    for (let rowIndex = 0; rowIndex < grid.length; rowIndex += 1) {
      const row = grid[rowIndex] ?? [];
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        const rawValue = row[colIndex] ?? "";
        const nextCell = {
          col: targetCell.col + colIndex,
          row: targetCell.row + rowIndex
        };
        // Overflow past the sheet and merged secondary cells are not
        // writable: skip them instead of corrupting the worksheet.
        if (!isWithinSheetBounds(nextCell) || worksheet.isMergedSecondary(nextCell.row, nextCell.col)) {
          continue;
        }
        const before = ctx.captureCellMutationState(nextCell);
        if (!before) {
          continue;
        }
        const input = resolveCellInput(rawValue, {
          date1904: worksheet.date1904 === true,
          locale: ctx.locale.value,
          numberFormat: resolveEffectiveCellNumberFormat(
            worksheet.getCellStyleAt(nextCell.row, nextCell.col),
            ctx.activeSheet.value,
            nextCell.row,
            nextCell.col
          )
        });
        if (input.kind !== "formula") {
          const validationCheck = ctx.checkCellValue(nextCell, rawValue);
          if (validationCheck.feedback) {
            ctx.setValidationFeedback(validationCheck.feedback);
          }
          if (!validationCheck.accepted) {
            continue;
          }
        }
        const address = cellAddressToA1(nextCell);
        if (input.kind === "formula") {
          worksheet.setFormula(address, input.formula);
        } else {
          worksheet.setCell(address, input.value);
        }
        if (input.numberFormat) {
          worksheet.setCellStyleAt(nextCell.row, nextCell.col, { numberFormat: input.numberFormat });
        }
        const after = ctx.captureCellMutationState(nextCell);
        if (!after) {
          continue;
        }
        mutations.push({
          after,
          before,
          cell: nextCell
        });
      }
    }

    if (ctx.maybeRecalculateWorkbook(ctx.workbook.value)) {
      ctx.refreshWorkbookState(ctx.workbook.value);
    }
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: Math.min(targetCell.col + Math.max(0, Math.max(...grid.map((row) => row.length), 1) - 1), SHEET_MAX_COLUMNS - 1),
        row: Math.min(targetCell.row + grid.length - 1, SHEET_MAX_ROWS - 1)
      }
    });
    ctx.activeCell.value = targetCell;
    ctx.selection.value = nextRange;
    ctx.selectionAnchorRef.value = targetCell;
    ctx.recordRangeEditHistory(mutations, nextRange, targetCell, beforeSnapshot);
    return true;
  }

  function pasteStructuredClipboardData(serializedPayload: string) {
    const worksheet = ctx.getActiveWorksheet();
    const targetCell = ctx.activeCell.value ?? ctx.selection.value?.start ?? null;
    if (ctx.isEditingLocked.value || !worksheet || !ctx.workbook.value || !targetCell || !serializedPayload) {
      return false;
    }

    let payload: ClipboardPayload;
    try {
      payload = JSON.parse(serializedPayload) as ClipboardPayload;
    } catch {
      return false;
    }

    if (!Array.isArray(payload.cells) || payload.cells.length === 0) {
      return false;
    }

    const hasMergeOperations = Array.isArray(payload.merges) && payload.merges.some((merge) => (merge.rowSpan ?? 1) > 1 || (merge.colSpan ?? 1) > 1);
    // Formulas copied from another range shift by the distance between the
    // copied range's origin and the paste target, like Excel. Payloads from
    // other hosts may omit the origin; those paste verbatim.
    const rowDelta = payload.origin ? targetCell.row - payload.origin.row : 0;
    const colDelta = payload.origin ? targetCell.col - payload.origin.col : 0;
    const beforeSnapshot = {
      activeCell: ctx.activeCell.value,
      selection: ctx.selection.value
    };
    ctx.setValidationFeedback(null);
    const mutations: import("./internal").RangeCellMutation[] = [];
    if (hasMergeOperations) {
      if (!pasteHistorySuppressed) {
        ctx.recordHistoryBeforeMutation();
      }
    }
    for (const cell of payload.cells) {
      const nextCell = {
        col: targetCell.col + cell.colOffset,
        row: targetCell.row + cell.rowOffset
      };
      // Overflow past the sheet and merged secondary cells are not writable:
      // skip them instead of corrupting the merge.
      if (!isWithinSheetBounds(nextCell) || worksheet.isMergedSecondary(nextCell.row, nextCell.col)) {
        continue;
      }
      if (!cell.formula) {
        const validationValue = cell.rawValue !== undefined
          ? String(cell.rawValue ?? "")
          : typeof cell.value === "string" ? cell.value : "";
        const validationCheck = ctx.checkCellValue(nextCell, validationValue);
        if (validationCheck.feedback) {
          ctx.setValidationFeedback(validationCheck.feedback);
        }
        if (!validationCheck.accepted) {
          continue;
        }
      }
      const before = hasMergeOperations ? null : ctx.captureCellMutationState(nextCell);
      const address = cellAddressToA1(nextCell);

      if (cell.formula) {
        worksheet.setFormula(address, shiftFormulaReferences(cell.formula, rowDelta, colDelta));
      } else if (cell.rawValue !== undefined) {
        worksheet.setCell(address, cell.rawValue);
      } else {
        // Payloads written before the raw value was carried only hold the
        // display string; coerce it like typed input instead of pasting text.
        const input = resolveCellInput(typeof cell.value === "string" ? cell.value : "", {
          date1904: worksheet.date1904 === true,
          locale: ctx.locale.value,
          numberFormat: resolveEffectiveCellNumberFormat(
            worksheet.getCellStyleAt(nextCell.row, nextCell.col),
            ctx.activeSheet.value,
            nextCell.row,
            nextCell.col
          )
        });
        if (input.kind === "formula") {
          worksheet.setFormula(address, shiftFormulaReferences(input.formula, rowDelta, colDelta));
        } else {
          worksheet.setCell(address, input.value);
        }
        if (input.numberFormat) {
          worksheet.setCellStyleAt(nextCell.row, nextCell.col, { numberFormat: input.numberFormat });
        }
      }

      if (cell.style !== undefined) {
        worksheet.setCellStyleAt(
          nextCell.row,
          nextCell.col,
          cell.style && typeof cell.style === "object" ? cell.style : DEFAULT_CELL_STYLE_PATCH
        );
      }

      if (before) {
        const after = ctx.captureCellMutationState(nextCell);
        if (!after) {
          continue;
        }
        mutations.push({
          after,
          before,
          cell: nextCell
        });
      }
    }

    if (Array.isArray(payload.merges)) {
      for (const merge of payload.merges) {
        if ((merge.rowSpan ?? 1) <= 1 && (merge.colSpan ?? 1) <= 1) {
          continue;
        }

        const mergeRange = normalizeRange({
          start: {
            col: targetCell.col + merge.colOffset,
            row: targetCell.row + merge.rowOffset
          },
          end: {
            col: targetCell.col + merge.colOffset + merge.colSpan - 1,
            row: targetCell.row + merge.rowOffset + merge.rowSpan - 1
          }
        });
        if (!isWithinSheetBounds(mergeRange.start) || !isWithinSheetBounds(mergeRange.end)) {
          continue;
        }
        worksheet.mergeCells(rangeToA1(mergeRange));
      }
    }

    if (ctx.maybeRecalculateWorkbook(ctx.workbook.value)) {
      ctx.refreshWorkbookState(ctx.workbook.value);
    }
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: Math.min(targetCell.col + Math.max((payload.cols ?? 1) - 1, 0), SHEET_MAX_COLUMNS - 1),
        row: Math.min(targetCell.row + Math.max((payload.rows ?? 1) - 1, 0), SHEET_MAX_ROWS - 1)
      }
    });
    ctx.activeCell.value = targetCell;
    ctx.selection.value = nextRange;
    ctx.selectionAnchorRef.value = targetCell;
    if (!hasMergeOperations && !pasteHistorySuppressed) {
      ctx.recordRangeEditHistory(mutations, nextRange, targetCell, beforeSnapshot);
    }
    return true;
  }

  async function copySelectionToClipboard() {
    const clipboardData = storeSelectionForPaste(false);
    if (!clipboardData) {
      return false;
    }

    return writeClipboardPayload(
      clipboardData,
      resolveClipboardWriteTargets(INTERNAL_CLIPBOARD_MIME)
    );
  }

  function cutSelection() {
    const clipboardData = ctx.isEditingLocked.value ? null : storeSelectionForPaste(true);
    if (!clipboardData) {
      return false;
    }
    // The internal store already holds the move payload; writing the system
    // clipboard is best-effort because not every host exposes one.
    void writeClipboardPayload(
      clipboardData,
      resolveClipboardWriteTargets(INTERNAL_CLIPBOARD_MIME)
    ).catch(() => false);
    return true;
  }

  /**
   * Cancels a pending cut without touching the workbook.
   *
   * The source cells were never cleared (the clear happens on paste), so only
   * the marquee and the internal payload have to go; a later paste falls back
   * to the system clipboard and behaves like an ordinary copy.
   */
  function cancelCut() {
    if (!cutState.value) {
      return false;
    }
    resetClipboardState();
    return true;
  }

  async function pasteFromClipboard() {
    if (ctx.isEditingLocked.value) {
      return false;
    }
    // A cut stays pending in the internal store: the next paste is the move,
    // even when the host clipboard cannot be read back.
    if (storedClipboard?.cutRange) {
      return pasteSpecialStructuredData(JSON.stringify(storedClipboard.payload), { mode: "all" });
    }
    if (typeof navigator === "undefined" || !navigator.clipboard) {
      return storedClipboard
        ? pasteSpecialStructuredData(JSON.stringify(storedClipboard.payload), { mode: "all" })
        : false;
    }

    if (navigator.clipboard.read) {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes(INTERNAL_CLIPBOARD_MIME)) {
          const blob = await item.getType(INTERNAL_CLIPBOARD_MIME);
          return pasteStructuredClipboardData(await blob.text());
        }
      }

      for (const item of items) {
        if (item.types.includes("text/plain")) {
          const blob = await item.getType("text/plain");
          return pasteText(await blob.text());
        }
      }
    }

    return pasteText(await navigator.clipboard.readText());
  }

  return {
    cancelCut,
    copySelectionToClipboard,
    cutSelection,
    cutState,
    getCellDisplayValue,
    getCellFormula,
    getClipboardData,
    pasteFromClipboard,
    pasteSpecialFromClipboard,
    pasteSpecialStructuredData,
    pasteStructuredClipboardData,
    pasteText
  };
}
