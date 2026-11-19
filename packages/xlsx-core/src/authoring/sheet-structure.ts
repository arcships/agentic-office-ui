import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxConditionalFormatRule,
  XlsxDataValidation,
  XlsxFreezePanes,
  XlsxImageAnchor,
  XlsxNamedRange,
  XlsxSparkline,
  XlsxTable,
} from "../types";
import { columnLabel } from "../core";
import {
  SPREADSHEET_NS,
  getLocalElements,
  normalizeArchivePath,
  parseXml,
  readArchiveText,
  serializeXml,
  type ArchiveEntries,
} from "../images";
import {
  deleteWorkbookTable,
  readWorkbookSheetAuthoringState,
  replaceWorkbookConditionalFormatRules,
  replaceWorkbookDataValidations,
  setWorkbookCellComment,
  setWorkbookCellHyperlink,
  setWorkbookFreezePanes,
  updateWorkbookConditionalFormatSqrefs,
  updateWorkbookTableRange,
  type WorkbookAuthoringAssets,
} from "./sheet-authoring";

// ── Structure primitives ─────────────────────────────────────────────────────

export type XlsxStructureAxis = "row" | "column";

/**
 * A row/column insertion or deletion expressed in zero-based sheet indices.
 *
 * `count` is clamped to at least one; `index` is clamped to the sheet bounds.
 */
export type XlsxStructureChange = {
  axis: XlsxStructureAxis;
  /** Zero-based first affected row or column. */
  index: number;
  /** Number of rows/columns to insert or delete. */
  count: number;
  kind: "insert" | "delete";
};

export const XLSX_MAX_ROWS = 1_048_576;
export const XLSX_MAX_COLUMNS = 16_384;

function normalizeChange(change: XlsxStructureChange): XlsxStructureChange {
  const limit = change.axis === "row" ? XLSX_MAX_ROWS : XLSX_MAX_COLUMNS;
  const index = Math.max(0, Math.min(limit - 1, Math.floor(Number(change.index) || 0)));
  const count = Math.max(1, Math.min(limit - index, Math.floor(Number(change.count) || 1)));
  return { axis: change.axis, count, index, kind: change.kind };
}

function axisValue(value: XlsxCellAddress, axis: XlsxStructureAxis) {
  return axis === "row" ? value.row : value.col;
}

function withAxisValue(value: XlsxCellAddress, axis: XlsxStructureAxis, next: number): XlsxCellAddress {
  return axis === "row" ? { ...value, row: next } : { ...value, col: next };
}

/**
 * Maps one zero-based row/column index through an insertion or deletion.
 *
 * @returns The shifted index, or `null` when a deletion removes it.
 */
export function remapAxisIndexForStructure(
  value: number,
  change: XlsxStructureChange,
): number | null {
  const normalized = normalizeChange(change);
  if (normalized.kind === "insert") {
    return value >= normalized.index ? value + normalized.count : value;
  }
  if (value < normalized.index) {
    return value;
  }
  if (value >= normalized.index + normalized.count) {
    return value - normalized.count;
  }
  return null;
}

/** Maps a cell through an insertion or deletion; `null` means it was removed. */
export function remapCellForStructure(
  cell: XlsxCellAddress,
  change: XlsxStructureChange,
): XlsxCellAddress | null {
  const nextRow = change.axis === "row" ? remapAxisIndexForStructure(cell.row, change) : cell.row;
  const nextCol = change.axis === "column" ? remapAxisIndexForStructure(cell.col, change) : cell.col;
  if (nextRow === null || nextCol === null) {
    return null;
  }
  return { col: nextCol, row: nextRow };
}

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  return {
    end: {
      col: Math.max(range.start.col, range.end.col),
      row: Math.max(range.start.row, range.end.row),
    },
    start: {
      col: Math.min(range.start.col, range.end.col),
      row: Math.min(range.start.row, range.end.row),
    },
  };
}

/**
 * Maps a range through an insertion or deletion.
 *
 * Inserting inside a range expands it, deleting a leading/trailing slice shrinks
 * it, and deleting everything the range covers removes it (`null`).
 */
export function remapRangeForStructure(
  range: XlsxCellRange,
  change: XlsxStructureChange,
): XlsxCellRange | null {
  const normalized = normalizeChange(change);
  const target = normalizeRange(range);
  const start = axisValue(target.start, normalized.axis);
  const end = axisValue(target.end, normalized.axis);
  if (normalized.kind === "insert") {
    const nextStart = start >= normalized.index ? start + normalized.count : start;
    const nextEnd = end >= normalized.index ? end + normalized.count : end;
    return {
      end: withAxisValue(target.end, normalized.axis, nextEnd),
      start: withAxisValue(target.start, normalized.axis, nextStart),
    };
  }
  if (end < normalized.index) {
    return target;
  }
  if (start >= normalized.index + normalized.count) {
    return {
      end: withAxisValue(target.end, normalized.axis, end - normalized.count),
      start: withAxisValue(target.start, normalized.axis, start - normalized.count),
    };
  }
  const nextStart = start < normalized.index ? start : normalized.index;
  const nextEnd = end >= normalized.index + normalized.count ? end - normalized.count : normalized.index - 1;
  if (nextEnd < nextStart) {
    return null;
  }
  return {
    end: withAxisValue(target.end, normalized.axis, nextEnd),
    start: withAxisValue(target.start, normalized.axis, nextStart),
  };
}

/** Maps every range, dropping the ones a deletion removed. */
export function remapRangesForStructure(
  ranges: readonly XlsxCellRange[],
  change: XlsxStructureChange,
): XlsxCellRange[] {
  return ranges.flatMap((range) => {
    const next = remapRangeForStructure(range, change);
    return next ? [next] : [];
  });
}

/**
 * Maps a defined name's A1 reference (`Sheet1!$A$1:$C$3`, `SUM(...)`) through
 * an insertion or deletion on `sheetName`.
 *
 * References to other sheets are preserved verbatim.
 */
export function remapDefinedNameForStructure(
  refersTo: string,
  change: XlsxStructureChange,
  sheetName?: string,
): string {
  return remapA1References(refersTo, change, { sheetName });
}

/** Maps frozen row/column counts through an insertion or deletion. */
export function remapFreezePanesForStructure(
  freeze: XlsxFreezePanes | null | undefined,
  change: XlsxStructureChange,
): XlsxFreezePanes | null {
  if (!freeze) {
    return null;
  }
  const normalized = normalizeChange(change);
  const current = normalized.axis === "row" ? freeze.row : freeze.col;
  let next = current;
  if (normalized.kind === "insert") {
    if (current > normalized.index) {
      next = current + normalized.count;
    }
  } else if (current > normalized.index + normalized.count) {
    next = current - normalized.count;
  } else if (current > normalized.index) {
    next = normalized.index;
  }
  const row = normalized.axis === "row" ? next : freeze.row;
  const col = normalized.axis === "column" ? next : freeze.col;
  return row > 0 || col > 0 ? { col, row } : null;
}

/**
 * Maps drawing markers (image/shaper/chart anchors) through a structure edit.
 *
 * Markers inside a deleted range collapse onto the deletion point, matching how
 * Excel keeps an anchored object alive when its anchor row disappears.
 */
export function remapImageAnchorForStructure(
  anchor: XlsxImageAnchor,
  change: XlsxStructureChange,
): XlsxImageAnchor {
  if (anchor.kind === "absolute") {
    return anchor;
  }
  const normalizeMarker = <T extends { col: number; row: number }>(marker: T): T => {
    const nextRow = change.axis === "row" ? remapMarker(marker.row, change) : marker.row;
    const nextCol = change.axis === "column" ? remapMarker(marker.col, change) : marker.col;
    return { ...marker, col: nextCol, row: nextRow };
  };
  return anchor.kind === "one-cell"
    ? { ...anchor, from: normalizeMarker(anchor.from) }
    : { ...anchor, from: normalizeMarker(anchor.from), to: normalizeMarker(anchor.to) };
}

function remapMarker(value: number, change: XlsxStructureChange) {
  const normalized = normalizeChange(change);
  const next = remapAxisIndexForStructure(value, normalized);
  return next === null ? normalized.index : next;
}

/** Maps sparkline anchors and source ranges, dropping deleted targets. */
export function remapSparklineForStructure(
  sparkline: XlsxSparkline,
  change: XlsxStructureChange,
): XlsxSparkline | null {
  const target = remapCellForStructure(sparkline.target, change);
  const range = remapRangeForStructure(sparkline.range, change);
  if (!target || !range) {
    return null;
  }
  return { ...sparkline, range, target };
}

/** Maps a conditional-format rule's ranges and formula thresholds. */
export function remapConditionalFormatRuleForStructure(
  rule: XlsxConditionalFormatRule,
  change: XlsxStructureChange,
): XlsxConditionalFormatRule | null {
  const ranges = remapRangesForStructure(rule.ranges, change);
  if (ranges.length === 0) {
    return null;
  }
  const cfvos = rule.cfvos.map((cfvo) => {
    if (cfvo.type !== "formula" || typeof cfvo.value !== "string") {
      return cfvo;
    }
    const value = remapA1References(cfvo.value, change);
    return value === cfvo.value ? cfvo : { ...cfvo, value };
  });
  return { ...rule, cfvos, ranges };
}

/** Maps a data-validation rule's ranges and reference formulas. */
export function remapDataValidationForStructure(
  validation: XlsxDataValidation,
  change: XlsxStructureChange,
): XlsxDataValidation | null {
  const ranges = remapRangesForStructure(validation.ranges, change);
  if (ranges.length === 0) {
    return null;
  }
  const remapFormula = (value: string | undefined) => {
    if (!value || value.startsWith("\"")) {
      return value;
    }
    return remapA1References(value, change);
  };
  return {
    ...validation,
    listSource: remapFormula(validation.listSource),
    ranges,
    value1: remapFormula(validation.value1),
    value2: remapFormula(validation.value2),
  };
}

/** Moves table range bookkeeping after a structural edit. */
export function remapTableForStructure(
  table: Pick<XlsxTable, "end" | "start">,
  change: XlsxStructureChange,
): { end: XlsxCellAddress; start: XlsxCellAddress } | null {
  const start = remapCellForStructure(table.start, change);
  const end = remapCellForStructure(table.end, change);
  if (!start || !end || end.row < start.row || end.col < start.col) {
    return null;
  }
  return { end, start };
}

// ── A1 reference remapping ───────────────────────────────────────────────────

const CELL_REFERENCE = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y;
const SHEET_PREFIX = /(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!/y;

function readCellReference(text: string, index: number) {
  CELL_REFERENCE.lastIndex = index;
  const match = CELL_REFERENCE.exec(text);
  if (!match) {
    return null;
  }
  const before = index > 0 ? text[index - 1] : "";
  const after = text[index + match[0].length] ?? "";
  if (before && /[A-Za-z0-9_.$]/.test(before)) {
    return null;
  }
  if (after && /[A-Za-z0-9_(!]/.test(after)) {
    return null;
  }
  const column = columnIndexFromLabel(match[2]);
  const row = Number(match[4]);
  if (column < 1 || column > XLSX_MAX_COLUMNS || row < 1 || row > XLSX_MAX_ROWS) {
    return null;
  }
  return {
    column,
    columnAnchor: match[1] === "$",
    row,
    rowAnchor: match[3] === "$",
    token: match[0],
  };
}

function columnIndexFromLabel(label: string) {
  let index = 0;
  for (const char of label.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index;
}

function readSheetPrefix(text: string, index: number) {
  SHEET_PREFIX.lastIndex = index;
  const match = SHEET_PREFIX.exec(text);
  if (!match) {
    return null;
  }
  const raw = match[0].slice(0, -1);
  const name = raw.startsWith("'") ? raw.slice(1, -1).replace(/''/g, "'") : raw;
  return { length: match[0].length, name };
}

function sameSheetName(left: string, right: string) {
  return left.trim().toLowerCase() === right.trim().toLowerCase();
}

type ReferenceRemapOptions = {
  /** Only references to this sheet are adjusted. */
  sheetName?: string;
  /**
   * Leave unqualified references alone. Formulas on another sheet use this so
   * only their `EditedSheet!A1` references move; their bare `A1` references
   * belong to their own sheet.
   */
  skipUnqualified?: boolean;
};

type ParsedReference = {
  column: number;
  columnAnchor: boolean;
  row: number;
  rowAnchor: boolean;
  token: string;
};

function formatReference(
  reference: ParsedReference,
  column: number | null,
  row: number | null,
) {
  if (column === null || row === null) {
    return "#REF!";
  }
  return `${reference.columnAnchor ? "$" : ""}${columnLabel(column - 1)}${reference.rowAnchor ? "$" : ""}${row}`;
}

function remapSingleReference(
  reference: ParsedReference,
  change: XlsxStructureChange,
) {
  const normalized = normalizeChange(change);
  const value = normalized.axis === "row" ? reference.row : reference.column;
  const next = remapAxisIndexForStructure(value - 1, normalized);
  if (next === null) {
    return "#REF!";
  }
  const row = normalized.axis === "row" ? next + 1 : reference.row;
  const column = normalized.axis === "column" ? next + 1 : reference.column;
  return formatReference(reference, column, row);
}

function remapRangeReferences(
  start: ParsedReference,
  end: ParsedReference,
  change: XlsxStructureChange,
) {
  const normalized = normalizeChange(change);
  const startValue = normalized.axis === "row" ? start.row - 1 : start.column - 1;
  const endValue = normalized.axis === "row" ? end.row - 1 : end.column - 1;
  const nextStart = remapAxisIndexForStructure(startValue, normalized);
  const nextEnd = remapAxisIndexForStructure(endValue, normalized);
  if (nextStart !== null && nextEnd !== null) {
    return `${remapSingleReference(start, change)}:${remapSingleReference(end, change)}`;
  }
  if (nextStart === null && nextEnd === null) {
    return "#REF!";
  }
  if (normalized.kind === "insert") {
    return `${remapSingleReference(start, change)}:${remapSingleReference(end, change)}`;
  }
  // One endpoint fell inside the deleted span: clamp it to the surviving edge
  // so a range like A1:A10 deleting rows 2-3 becomes A1:A8, not #REF!:A8.
  const startNext = nextStart === null ? normalized.index : nextStart;
  const endNext = nextEnd === null ? normalized.index - 1 : nextEnd;
  if (endNext < startNext) {
    return "#REF!";
  }
  const startRef = normalized.axis === "row"
    ? formatReference(start, start.column, startNext + 1)
    : formatReference(start, startNext + 1, start.row);
  const endRef = normalized.axis === "row"
    ? formatReference(end, end.column, endNext + 1)
    : formatReference(end, endNext + 1, end.row);
  return `${startRef}:${endRef}`;
}

/**
 * Rewrites the A1 references in a formula (or a defined-name reference) after a
 * row/column insertion or deletion, the way Excel does.
 *
 * The rewrite includes absolute (`$`) references, because moving a row moves
 * every reference to it regardless of anchoring. String literals, quoted sheet
 * names, function names that look like addresses (`LOG10`) and references to
 * other sheets are left untouched. A reference removed by a deletion becomes
 * `#REF!`.
 *
 * @param text Formula body, with or without a leading `=`.
 * @param change Structure change to apply.
 * @param options `sheetName` limits the rewrite to that sheet's references.
 */
export function remapA1References(
  text: string,
  change: XlsxStructureChange,
  options: ReferenceRemapOptions = {},
): string {
  const normalized = normalizeChange(change);
  let result = "";
  let index = 0;
  while (index < text.length) {
    const char = text[index] ?? "";

    if (char === "\"") {
      let end = index + 1;
      while (end < text.length) {
        if (text[end] === "\"") {
          if (text[end + 1] === "\"") {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      result += text.slice(index, end);
      index = end;
      continue;
    }

    // `'My Sheet'!A1` and `Sheet1!A1` need the prefix parsed first so a quoted
    // sheet name is not mistaken for a string literal.
    const prefix = readSheetPrefix(text, index);
    if (prefix) {
      const reference = readCellReference(text, index + prefix.length);
      if (reference) {
        const shouldRemap = !options.sheetName || sameSheetName(prefix.name, options.sheetName);
        const tokenStart = index + prefix.length;
        const rangeColon = text[tokenStart + reference.token.length];
        let end = tokenStart + reference.token.length;
        let replacement = reference.token;
        if (rangeColon === ":") {
          const rangeEnd = readCellReference(text, end + 1);
          if (rangeEnd) {
            replacement = shouldRemap
              ? remapRangeReferences(reference, rangeEnd, normalized)
              : text.slice(tokenStart, end + 1 + rangeEnd.token.length);
            end = end + 1 + rangeEnd.token.length;
          } else if (shouldRemap) {
            replacement = remapSingleReference(reference, normalized);
          }
        } else if (shouldRemap) {
          replacement = remapSingleReference(reference, normalized);
        }
        result += text.slice(index, tokenStart) + replacement;
        index = end;
        continue;
      }
    }

    if (char === "'") {
      let end = index + 1;
      while (end < text.length) {
        if (text[end] === "'") {
          if (text[end + 1] === "'") {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      result += text.slice(index, end);
      index = end;
      continue;
    }

    const reference = readCellReference(text, index);
    if (reference) {
      const rangeColon = text[index + reference.token.length];
      let end = index + reference.token.length;
      if (rangeColon === ":") {
        const rangeEnd = readCellReference(text, end + 1);
        if (rangeEnd) {
          end = end + 1 + rangeEnd.token.length;
        }
      }
      if (options.skipUnqualified) {
        result += text.slice(index, end);
        index = end;
        continue;
      }
      let replacement = remapSingleReference(reference, normalized);
      if (rangeColon === ":") {
        const rangeEnd = readCellReference(text, index + reference.token.length + 1);
        if (rangeEnd) {
          replacement = remapRangeReferences(reference, rangeEnd, normalized);
        }
      }
      result += replacement;
      index = end;
      continue;
    }

    result += char;
    index += 1;
  }
  return result;
}

// ── Workbook-level structure mutation ────────────────────────────────────────

function remapHyperlinkTarget(
  target: string,
  change: XlsxStructureChange,
  sheetName: string | undefined,
) {
  // `#Sheet!A1` links point inside the workbook; external URLs (even with a
  // fragment) are never rewritten.
  if (!target.startsWith("#")) {
    return target;
  }
  return `#${remapA1References(target.slice(1), change, { sheetName })}`;
}

type ShiftSheetStructureOptions = {
  /** Sheet name the change applies to; limits formula/reference rewrites. */
  sheetName?: string;
};

/**
 * Rewrites every authored object that lives in the worksheet XML after a
 * row/column insertion or deletion.
 *
 * Merged-cell geometry and cell contents are mutated on the workbook model by
 * the viewer; this function handles the parts the WASM engine does not expose:
 * hyperlinks, comments, conditional-format ranges and thresholds, data
 * validations, tables and frozen panes. Drawing anchors are updated separately
 * through the image/chart asset APIs.
 *
 * @returns `true` when the worksheet part was rewritten.
 */
export function shiftWorkbookSheetStructure(
  assets: WorkbookAuthoringAssets,
  workbookSheetIndex: number,
  change: XlsxStructureChange,
  options: ShiftSheetStructureOptions = {},
): boolean {
  const state = readWorkbookSheetAuthoringState(assets, workbookSheetIndex);
  if (!state) {
    return false;
  }
  const normalized = normalizeChange(change);

  const rules = state.conditionalFormatRules.flatMap((rule) => {
    const next = remapConditionalFormatRuleForStructure(rule, normalized);
    return next ? [next] : [];
  });
  // Raw (unmodeled) rules are preserved by the XML writer, so their sqref and
  // formulas have to be shifted before the modeled rules are rewritten.
  updateWorkbookConditionalFormatSqrefs(
    assets,
    workbookSheetIndex,
    (ranges) => remapRangesForStructure(ranges, normalized),
    (formula) => remapA1References(formula, normalized, { sheetName: options.sheetName })
  );
  replaceWorkbookConditionalFormatRules(assets, workbookSheetIndex, rules);

  const validations = state.dataValidations.flatMap((validation) => {
    const next = remapDataValidationForStructure(validation, normalized);
    return next ? [next] : [];
  });
  replaceWorkbookDataValidations(assets, workbookSheetIndex, validations);

  const nextFreeze = remapFreezePanesForStructure(state.freezePanes, normalized);
  if (JSON.stringify(nextFreeze) !== JSON.stringify(state.freezePanes)) {
    setWorkbookFreezePanes(assets, workbookSheetIndex, nextFreeze);
  }

  const hyperlinks = state.hyperlinks.flatMap((hyperlink) => {
    const cell = remapCellForStructure(hyperlink.cell, normalized);
    return cell ? [{ ...hyperlink, cell, target: remapHyperlinkTarget(hyperlink.target, normalized, options.sheetName) }] : [];
  });
  if (state.hyperlinks.length > 0) {
    for (const hyperlink of state.hyperlinks) {
      setWorkbookCellHyperlink(assets, workbookSheetIndex, hyperlink.cell, null);
    }
    for (const hyperlink of hyperlinks) {
      setWorkbookCellHyperlink(assets, workbookSheetIndex, hyperlink.cell, {
        target: hyperlink.target,
        tooltip: hyperlink.tooltip,
      });
    }
  }

  const comments = state.comments.flatMap((comment) => {
    const cell = remapCellForStructure(comment.cell, normalized);
    return cell ? [{ ...comment, cell }] : [];
  });
  if (state.comments.length > 0) {
    for (const comment of state.comments) {
      setWorkbookCellComment(assets, workbookSheetIndex, comment.cell, null);
    }
    for (const comment of comments) {
      setWorkbookCellComment(assets, workbookSheetIndex, comment.cell, {
        author: comment.author,
        text: comment.text,
      });
    }
  }

  for (const table of state.tables) {
    const nextRange = remapTableForStructure(table, normalized);
    if (!nextRange) {
      deleteWorkbookTable(assets, workbookSheetIndex, table.name);
      continue;
    }
    updateWorkbookTableRange(assets, workbookSheetIndex, table.name, nextRange);
  }

  // The pure remappers are exported for hosts that keep their own copies of
  // these objects; `options.sheetName` is accepted so callers can pass the same
  // filter they use for formula rewrites without a second signature.
  void options;
  return true;
}

// ── Defined names ────────────────────────────────────────────────────────────

/** Normalizes the engine's `namedRanges` payload into the public shape. */
export function normalizeWorkbookDefinedNames(value: unknown): XlsxNamedRange[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.flatMap((entry) => {
    if (!entry || typeof entry !== "object") {
      return [];
    }
    const record = entry as Record<string, unknown>;
    const name = typeof record.name === "string" ? record.name.trim() : "";
    const refersTo = typeof record.refersTo === "string" ? record.refersTo.trim() : "";
    if (!name || !refersTo) {
      return [];
    }
    const sheetIndex = typeof record.sheetIndex === "number" && Number.isFinite(record.sheetIndex)
      ? Math.max(0, Math.floor(record.sheetIndex))
      : undefined;
    return [{
      comment: typeof record.comment === "string" ? record.comment : undefined,
      hidden: record.hidden === true,
      name,
      refersTo,
      sheetIndex,
    } satisfies XlsxNamedRange];
  });
}

function escapeXmlText(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const DEFINED_NAMES_FOLLOWERS = [
  "calcPr",
  "oleSize",
  "customWorkbookViews",
  "pivotCaches",
  "smartTagPr",
  "smartTagTypes",
  "webPublishing",
  "fileRecoveryPr",
  "webPublishObjects",
  "extLst",
];

/**
 * Replaces the workbook's `definedNames` list in saved XLSX bytes.
 *
 * The WASM engine can add and overwrite names but cannot delete one, so the
 * viewer writes the effective list (engine names plus Name Manager edits) into
 * the saved workbook part. Everything else in `xl/workbook.xml` is preserved.
 */
export function applyWorkbookDefinedNames(
  bytes: Uint8Array,
  names: readonly XlsxNamedRange[],
): Uint8Array {
  let archive: ArchiveEntries;
  try {
    archive = unzipSync(bytes);
  } catch {
    return new Uint8Array(bytes);
  }
  const workbookXml = readArchiveText(archive, "xl/workbook.xml");
  if (!workbookXml) {
    return new Uint8Array(bytes);
  }
  const document = parseXml(workbookXml);
  const root = document?.documentElement;
  if (!document || !root) {
    return new Uint8Array(bytes);
  }

  for (const existing of getLocalElements(document, "definedNames")) {
    existing.parentNode?.removeChild(existing);
  }
  if (names.length > 0) {
    const container = document.createElementNS(SPREADSHEET_NS, "definedNames");
    for (const name of names) {
      const node = document.createElementNS(SPREADSHEET_NS, "definedName");
      node.setAttribute("name", name.name);
      if (name.sheetIndex !== undefined) {
        node.setAttribute("localSheetId", String(name.sheetIndex));
      }
      if (name.comment) {
        node.setAttribute("comment", name.comment);
      }
      if (name.hidden) {
        node.setAttribute("hidden", "1");
      }
      node.textContent = name.refersTo;
      container.appendChild(node);
    }
    const follower = Array.from(root.childNodes).find((child) => {
      if (child.nodeType !== 1) {
        return false;
      }
      return DEFINED_NAMES_FOLLOWERS.includes((child as Element).localName);
    });
    if (follower) {
      root.insertBefore(container, follower);
    } else {
      root.appendChild(container);
    }
  }

  archive[normalizeArchivePath("xl/workbook.xml")] = strToU8(serializeXml(document));
  return zipSync(archive, { level: 6 });
}

/** Reads the defined names out of saved XLSX bytes. */
export function readWorkbookDefinedNames(bytes: Uint8Array): XlsxNamedRange[] {
  let archive: ArchiveEntries;
  try {
    archive = unzipSync(bytes);
  } catch {
    return [];
  }
  const workbookXml = readArchiveText(archive, "xl/workbook.xml");
  const document = workbookXml ? parseXml(workbookXml) : null;
  if (!document) {
    return [];
  }
  return getLocalElements(document, "definedName").flatMap((node) => {
    const name = node.getAttribute("name")?.trim() ?? "";
    const refersTo = (node.textContent ?? "").trim();
    if (!name || !refersTo) {
      return [];
    }
    const localSheetId = node.getAttribute("localSheetId");
    const parsedSheetIndex = localSheetId === null ? null : Number(localSheetId);
    return [{
      comment: node.getAttribute("comment") ?? undefined,
      hidden: node.getAttribute("hidden") === "1",
      name,
      refersTo,
      sheetIndex: parsedSheetIndex !== null && Number.isFinite(parsedSheetIndex)
        ? Math.max(0, Math.floor(parsedSheetIndex))
        : undefined,
    } satisfies XlsxNamedRange];
  });
}
