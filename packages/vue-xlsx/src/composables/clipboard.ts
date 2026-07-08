import { XLSX_MIME_TYPE, CSV_MIME_TYPE, INTERNAL_CLIPBOARD_MIME } from "./internal";
import type { ClipboardPayload, ClipboardMatrixCell, ClipboardMerge } from "./internal";
import type { Workbook } from "@dukelib/sheets-wasm";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxClipboardData
} from "@extend-ai/xlsx-core";
import { resolveWorkbookColor, resolveWorkbookFillStyle } from "@extend-ai/xlsx-core";
import { cellAddressToA1, normalizeRange, rangeToA1 } from "./selection";
import { decodeHtmlEntities, escapeHtml, mapBorder, resolveInheritedCellStyle } from "./formatting";
import { coerceUserEnteredValue } from "./internal";
import type { XlsxControllerContext } from "./internal";

export { INTERNAL_CLIPBOARD_MIME, escapeHtml };
export type { ClipboardPayload, ClipboardMatrixCell, ClipboardMerge };

export function parseClipboardText(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const rows = normalized.split("\n");

  if (rows.length > 1 && rows[rows.length - 1] === "") {
    rows.pop();
  }

  return rows.map((row) => row.split("\t"));
}

export function downloadArrayBuffer(file: ArrayBuffer, fileName: string) {
  const blob = new Blob([file], { type: XLSX_MIME_TYPE });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

export function downloadBytes(bytes: Uint8Array, fileName: string, mimeType: string) {
  const normalizedBytes = new Uint8Array(bytes.byteLength);
  normalizedBytes.set(bytes);
  const blob = new Blob([normalizedBytes], { type: mimeType });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

export function downloadText(text: string, fileName: string, mimeType: string) {
  const blob = new Blob([text], { type: mimeType });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

export function downloadUrl(src: string, fileName: string) {
  const anchor = document.createElement("a");
  anchor.href = src;
  anchor.download = fileName;
  anchor.rel = "noreferrer";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
}

export function createClipboardDomain(ctx: XlsxControllerContext) {
  function getCellDisplayValue(cell?: XlsxCellAddress | null) {
    if (cell && ctx.activeSheet.value) {
      const workerSnapshot = ctx.workerCellSnapshotCacheRef.value.get(`${ctx.activeSheet.value.workbookSheetIndex}:${cell.row}:${cell.col}`);
      if (workerSnapshot) {
        return workerSnapshot.displayValue;
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
      return decodeHtmlEntities(formatted);
    }

    const calculated = worksheet.getCalculatedValueAt(cell.row, cell.col);
    if (formula && cachedFormulaValue !== undefined && calculated.is_error) {
      return cachedFormulaValue;
    }
    if (calculated.is_error) {
      return calculated.asError() ?? "";
    }
    if (calculated.is_empty) {
      return "";
    }

    return calculated.toString();
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
    const rows: string[] = [];
    const htmlRows: string[] = [];
    const payload: ClipboardPayload = {
      cells: [],
      cols: normalized.end.col - normalized.start.col + 1,
      merges: [],
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
        const merge = worksheet.getMergeSpan(row, col) as
          | { colSpan?: number; rowSpan?: number }
          | null
          | undefined;
        const rawStyle = (
          worksheet.getCellStyleAt(row, col) as Record<string, unknown> | null | undefined
        ) ?? resolveInheritedCellStyle(ctx.activeSheet.value, row, col);
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
          rowOffset: row - normalized.start.row,
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

      rows.push(textCells.join("\t"));
      htmlRows.push(`<tr>${htmlCells.join("")}</tr>`);
    }

    return {
      html: `<table style="border-collapse:collapse">${htmlRows.join("")}</table>`,
      structured: JSON.stringify(payload),
      text: rows.join("\n")
    };
  }

  function pasteText(text: string) {
    const worksheet = ctx.getActiveWorksheet();
    const targetCell = ctx.activeCell.value ?? ctx.selection.value?.start ?? null;
    if (ctx.readOnly.value || !worksheet || !ctx.workbook.value || !targetCell || !text) {
      return false;
    }

    const grid = parseClipboardText(text);
    if (grid.length === 0 || grid.every((row) => row.length === 0)) {
      return false;
    }

    const mutations: import("./internal").RangeCellMutation[] = [];
    for (let rowIndex = 0; rowIndex < grid.length; rowIndex += 1) {
      const row = grid[rowIndex] ?? [];
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        const rawValue = row[colIndex] ?? "";
        const nextCell = {
          col: targetCell.col + colIndex,
          row: targetCell.row + rowIndex
        };
        const before = ctx.captureCellMutationState(nextCell);
        if (!before) {
          continue;
        }
        if (rawValue.startsWith("=") && rawValue.length > 1) {
          worksheet.setFormula(cellAddressToA1(nextCell), rawValue);
          const after = ctx.captureCellMutationState(nextCell);
          if (!after) {
            continue;
          }
          mutations.push({
            after,
            before,
            cell: nextCell
          });
        } else {
          const nextValue = coerceUserEnteredValue(rawValue);
          worksheet.setCell(cellAddressToA1(nextCell), nextValue);
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
    }

    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: targetCell.col + Math.max(0, Math.max(...grid.map((row) => row.length), 1) - 1),
        row: targetCell.row + grid.length - 1
      }
    });
    ctx.activeCell.value = targetCell;
    ctx.selection.value = nextRange;
    ctx.selectionAnchorRef.value = targetCell;
    ctx.recordRangeEditHistory(mutations, nextRange, targetCell);
    return true;
  }

  function pasteStructuredClipboardData(serializedPayload: string) {
    const worksheet = ctx.getActiveWorksheet();
    const targetCell = ctx.activeCell.value ?? ctx.selection.value?.start ?? null;
    if (ctx.readOnly.value || !worksheet || !ctx.workbook.value || !targetCell || !serializedPayload) {
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
    const mutations: import("./internal").RangeCellMutation[] = [];
    if (hasMergeOperations) {
      ctx.recordHistoryBeforeMutation();
    }
    for (const cell of payload.cells) {
      const nextCell = {
        col: targetCell.col + cell.colOffset,
        row: targetCell.row + cell.rowOffset
      };
      const before = hasMergeOperations ? null : ctx.captureCellMutationState(nextCell);

      if (cell.formula) {
        worksheet.setFormula(cellAddressToA1(nextCell), cell.formula);
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
      } else {
        worksheet.setCell(cellAddressToA1(nextCell), cell.value);
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
        worksheet.mergeCells(rangeToA1(mergeRange));
      }
    }

    ctx.maybeRecalculateWorkbook(ctx.workbook.value);
    ctx.refreshWorkbookState(ctx.workbook.value);
    const nextRange = normalizeRange({
      start: targetCell,
      end: {
        col: targetCell.col + Math.max((payload.cols ?? 1) - 1, 0),
        row: targetCell.row + Math.max((payload.rows ?? 1) - 1, 0)
      }
    });
    ctx.activeCell.value = targetCell;
    ctx.selection.value = nextRange;
    ctx.selectionAnchorRef.value = targetCell;
    if (!hasMergeOperations) {
      ctx.recordRangeEditHistory(mutations, nextRange, targetCell);
    }
    return true;
  }

  async function copySelectionToClipboard() {
    const clipboardData = getClipboardData();
    if (!clipboardData || typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
    }

    if (typeof ClipboardItem === "function" && navigator.clipboard.write) {
      const item = new ClipboardItem({
        [INTERNAL_CLIPBOARD_MIME]: new Blob([clipboardData.structured], { type: INTERNAL_CLIPBOARD_MIME }),
        "text/html": new Blob([clipboardData.html], { type: "text/html" }),
        "text/plain": new Blob([clipboardData.text], { type: "text/plain" })
      });
      await navigator.clipboard.write([item]);
      return true;
    }

    await navigator.clipboard.writeText(clipboardData.text);
    return true;
  }

  async function pasteFromClipboard() {
    if (ctx.readOnly.value || typeof navigator === "undefined" || !navigator.clipboard) {
      return false;
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
    getCellDisplayValue,
    getCellFormula,
    getClipboardData,
    pasteText,
    pasteStructuredClipboardData,
    copySelectionToClipboard,
    pasteFromClipboard
  };
}
