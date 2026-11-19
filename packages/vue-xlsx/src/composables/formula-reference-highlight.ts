import type { XlsxCellRange } from "@arcships/xlsx-core";

/**
 * Excel's formula-editing palette. Each distinct reference gets the next colour
 * so two references in one formula are always visually distinguishable, and a
 * reference repeated later in the formula keeps the colour of its first
 * occurrence.
 */
export const FORMULA_REFERENCE_COLORS = [
  "#2563eb",
  "#16a34a",
  "#dc2626",
  "#9333ea",
  "#ea580c",
  "#0891b2",
  "#db2777",
  "#65a30d",
] as const;

/** Excel's hard worksheet bounds, repeated here so this module is dependency-free. */
const SHEET_MAX_COLUMNS = 16_384;
const SHEET_MAX_ROWS = 1_048_576;

export type FormulaReferenceKind = "cell" | "range" | "column" | "row" | "name";

export interface FormulaReferenceNameTarget {
  range: XlsxCellRange;
  sheetName?: string | null;
}

export interface FormulaReferenceOptions {
  /** Worksheet the formula is being edited on; used by host UIs to filter cross-sheet outlines. */
  activeSheetName?: string | null;
  /**
   * Resolves a defined name to the range it points at. Names are only
   * highlighted when they resolve; an unknown name is treated as plain text so
   * a typo is not silently coloured like a valid reference.
   */
  resolveName?: (
    name: string,
    sheetName: string | null,
  ) => FormulaReferenceNameTarget | null;
}

export interface FormulaReference {
  /** Index of the first character in the formula text. */
  start: number;
  /** Index just past the last character in the formula text. */
  end: number;
  /** Reference text exactly as typed, including `$` anchors and any sheet qualifier. */
  text: string;
  kind: FormulaReferenceKind;
  color: string;
  /** Zero-based target range on its worksheet. */
  range: XlsxCellRange;
  /** Canonical range text used by the grid's outline attribute and by hosts. */
  rangeAddress: string;
  /** Quoted sheet qualifier without quotes, or `null` for the active sheet. */
  sheetName: string | null;
  /** Defined name for `kind === "name"`. */
  name?: string;
}

export interface FormulaReferenceSegment {
  text: string;
  reference: FormulaReference | null;
  color: string | null;
  active: boolean;
}

export interface FormulaReferenceCollection {
  references: FormulaReference[];
  /** Index into `references` of the occurrence under the caret, or `-1`. */
  activeIndex: number;
  activeReference: FormulaReference | null;
}

const EMPTY_COLLECTION: FormulaReferenceCollection = {
  activeIndex: -1,
  activeReference: null,
  references: [],
};

const CELL_AT = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y;
const COLUMN_AT = /(\$?)([A-Za-z]{1,3})/y;
const ROW_AT = /(\$?)(\d{1,7})/y;
const NAME_AT = /[\p{L}_\\][\p{L}\p{N}_.\\]*/uy;
const SHEET_PREFIX_AT = /(?:'(?:[^']|'')+'|[\p{L}_\\][\p{L}\p{N}_.\\]*)!/uy;
const CELL_LIKE = /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/;
const FORMULA_LITERAL = /^(?:TRUE|FALSE|NULL)$/i;

function columnLabel(col: number): string {
  let label = "";
  let next = col;
  while (next >= 0) {
    label = String.fromCharCode(65 + (next % 26)) + label;
    next = Math.floor(next / 26) - 1;
  }
  return label;
}

function columnIndexFromLabel(label: string): number {
  let index = 0;
  for (const char of label.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index;
}

function normalizeRange(start: { row: number; col: number }, end: { row: number; col: number }): XlsxCellRange {
  return {
    start: {
      row: Math.min(start.row, end.row),
      col: Math.min(start.col, end.col),
    },
    end: {
      row: Math.max(start.row, end.row),
      col: Math.max(start.col, end.col),
    },
  };
}

function rangeAddress(range: XlsxCellRange): string {
  const start = `${columnLabel(range.start.col)}${range.start.row + 1}`;
  const end = `${columnLabel(range.end.col)}${range.end.row + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function isSheetPrefixBoundary(text: string, start: number): boolean {
  const before = start > 0 ? text[start - 1] : "";
  return !before || !/[\p{L}\p{N}_.$']/u.test(before);
}

function isTokenBoundary(text: string, start: number, end: number): boolean {
  const before = start > 0 ? text[start - 1] : "";
  const after = end < text.length ? text[end] : "";
  if (before && /[\p{L}\p{N}_.$]/u.test(before)) return false;
  if (after && /[\p{L}\p{N}_(!.$\[]/u.test(after)) return false;
  return true;
}

function skipQuoted(text: string, index: number, quote: string): number {
  let cursor = index + 1;
  while (cursor < text.length) {
    if (text[cursor] === quote) {
      if (text[cursor + 1] === quote) {
        cursor += 2;
        continue;
      }
      return cursor + 1;
    }
    cursor += 1;
  }
  return text.length;
}

function skipSpaces(text: string, index: number): number {
  let cursor = index;
  while (cursor < text.length && /\s/.test(text[cursor] ?? "")) cursor += 1;
  return cursor;
}

function readRangeColon(text: string, index: number): number {
  let cursor = skipSpaces(text, index);
  if (text[cursor] !== ":") return -1;
  cursor += 1;
  return skipSpaces(text, cursor);
}

function unquoteSheetName(value: string): string {
  if (!value.startsWith("'")) return value;
  return value.slice(1, -1).replace(/''/g, "'");
}

interface ParsedSheetPrefix {
  sheetName: string;
  end: number;
}

function matchSheetPrefix(text: string, index: number): ParsedSheetPrefix | null {
  if (!isSheetPrefixBoundary(text, index)) return null;
  SHEET_PREFIX_AT.lastIndex = index;
  const match = SHEET_PREFIX_AT.exec(text);
  if (!match) return null;
  const raw = match[0].slice(0, -1);
  return { sheetName: unquoteSheetName(raw), end: index + match[0].length };
}

interface ParsedCell {
  col: number;
  row: number;
  token: string;
}

function matchCell(text: string, index: number): ParsedCell | null {
  CELL_AT.lastIndex = index;
  const match = CELL_AT.exec(text);
  if (!match) return null;
  const col = columnIndexFromLabel(match[2] ?? "");
  const row = Number(match[4]);
  if (col < 1 || col > SHEET_MAX_COLUMNS || row < 1 || row > SHEET_MAX_ROWS) return null;
  return { col: col - 1, row: row - 1, token: match[0] };
}

interface ParsedColumn {
  col: number;
  token: string;
}

function matchColumn(text: string, index: number): ParsedColumn | null {
  COLUMN_AT.lastIndex = index;
  const match = COLUMN_AT.exec(text);
  if (!match) return null;
  const col = columnIndexFromLabel(match[2] ?? "");
  if (col < 1 || col > SHEET_MAX_COLUMNS) return null;
  return { col: col - 1, token: match[0] };
}

interface ParsedRow {
  row: number;
  token: string;
}

function matchRow(text: string, index: number): ParsedRow | null {
  ROW_AT.lastIndex = index;
  const match = ROW_AT.exec(text);
  if (!match) return null;
  const row = Number(match[2]);
  if (row < 1 || row > SHEET_MAX_ROWS) return null;
  return { row: row - 1, token: match[0] };
}

interface ParsedReferenceBody {
  end: number;
  kind: FormulaReferenceKind;
  range: XlsxCellRange;
  bodyAddress: string;
  name?: string;
  /** Sheet a defined name resolves to when its `refersTo` names one. */
  targetSheetName?: string | null;
}

function readReferenceBody(
  text: string,
  index: number,
  options: FormulaReferenceOptions,
): ParsedReferenceBody | null {
  const cell = matchCell(text, index);
  if (cell) {
    const afterCell = index + cell.token.length;
    const secondStart = readRangeColon(text, afterCell);
    if (secondStart >= 0) {
      const second = matchCell(text, secondStart);
      if (second) {
        const end = secondStart + second.token.length;
        if (isTokenBoundary(text, index, end)) {
          const range = normalizeRange(
            { row: cell.row, col: cell.col },
            { row: second.row, col: second.col },
          );
          return { bodyAddress: rangeAddress(range), end, kind: "range", range };
        }
      }
    } else if (isTokenBoundary(text, index, afterCell)) {
      const range = normalizeRange(
        { row: cell.row, col: cell.col },
        { row: cell.row, col: cell.col },
      );
      return {
        bodyAddress: rangeAddress(range),
        end: afterCell,
        kind: "cell",
        range,
      };
    }
  }

  const column = matchColumn(text, index);
  if (column) {
    const afterColumn = index + column.token.length;
    const secondStart = readRangeColon(text, afterColumn);
    if (secondStart >= 0) {
      const second = matchColumn(text, secondStart);
      if (second) {
        const end = secondStart + second.token.length;
        if (isTokenBoundary(text, index, end)) {
          const range = normalizeRange(
            { row: 0, col: column.col },
            { row: SHEET_MAX_ROWS - 1, col: second.col },
          );
          const startLabel = columnLabel(Math.min(column.col, second.col));
          const endLabel = columnLabel(Math.max(column.col, second.col));
          return {
            bodyAddress: `${startLabel}:${endLabel}`,
            end,
            kind: "column",
            range,
          };
        }
      }
    }
  }

  const row = matchRow(text, index);
  if (row) {
    const afterRow = index + row.token.length;
    const secondStart = readRangeColon(text, afterRow);
    if (secondStart >= 0) {
      const second = matchRow(text, secondStart);
      if (second) {
        const end = secondStart + second.token.length;
        if (isTokenBoundary(text, index, end)) {
          const range = normalizeRange(
            { row: row.row, col: 0 },
            { row: second.row, col: SHEET_MAX_COLUMNS - 1 },
          );
          const startLabel = Math.min(row.row, second.row) + 1;
          const endLabel = Math.max(row.row, second.row) + 1;
          return {
            bodyAddress: `${startLabel}:${endLabel}`,
            end,
            kind: "row",
            range,
          };
        }
      }
    }
  }

  NAME_AT.lastIndex = index;
  const nameMatch = NAME_AT.exec(text);
  if (!nameMatch) return null;
  const name = nameMatch[0];
  const end = index + name.length;
  if (!isTokenBoundary(text, index, end)) return null;
  if (CELL_LIKE.test(name) || FORMULA_LITERAL.test(name)) return null;
  const after = skipSpaces(text, end);
  if (text[after] === "(" || text[after] === "[") return null;
  if (name.toLowerCase() === "true" || name.toLowerCase() === "false") return null;
  const target = options.resolveName?.(name, null);
  if (!target) return null;
  return {
    bodyAddress: rangeAddress(target.range),
    end,
    kind: "name",
    name,
    range: target.range,
    targetSheetName: target.sheetName ?? null,
  };
}

interface ParsedReference extends ParsedReferenceBody {
  sheetName: string | null;
  start: number;
  text: string;
}

function tryReadReference(
  text: string,
  index: number,
  options: FormulaReferenceOptions,
): ParsedReference | null {
  let cursor = index;
  let sheetName: string | null = null;
  const prefix = matchSheetPrefix(text, index);
  if (prefix) {
    sheetName = prefix.sheetName;
    cursor = prefix.end;
  }
  const body = readReferenceBody(text, cursor, {
    ...options,
    resolveName: options.resolveName
      ? (name, _sheetName) => options.resolveName?.(name, sheetName) ?? null
      : undefined,
  });
  if (!body) return null;
  return {
    ...body,
    sheetName: body.kind === "name"
      ? body.targetSheetName ?? sheetName ?? null
      : sheetName ?? body.targetSheetName ?? null,
    start: index,
    text: text.slice(index, body.end),
  };
}

function skipExternalReference(text: string, index: number): number {
  const close = text.indexOf("]", index + 1);
  if (close < 0) return text.length;
  const bang = text.indexOf("!", close + 1);
  if (bang < 0) return close + 1;
  const body = readReferenceBody(text, bang + 1, {});
  return body?.end ?? bang + 1;
}

function scanReferenceTokens(
  text: string,
  options: FormulaReferenceOptions,
): ParsedReference[] {
  const references: ParsedReference[] = [];
  let index = 0;
  while (index < text.length) {
    const char = text[index] ?? "";
    if (char === '"') {
      index = skipQuoted(text, index, '"');
      continue;
    }
    if (char === "'") {
      const prefix = matchSheetPrefix(text, index);
      if (!prefix) {
        index = skipQuoted(text, index, "'");
        continue;
      }
    }
    if (char === "[") {
      index = skipExternalReference(text, index);
      continue;
    }
    const reference = tryReadReference(text, index, options);
    if (reference) {
      references.push(reference);
      index = reference.end;
      continue;
    }
    const prefix = matchSheetPrefix(text, index);
    index = prefix ? prefix.end : index + 1;
  }
  return references;
}

function collectReferences(
  text: string,
  caret: number,
  options: FormulaReferenceOptions,
): FormulaReferenceCollection {
  const parsed = scanReferenceTokens(text, options);
  if (parsed.length === 0) return EMPTY_COLLECTION;

  const colorByKey = new Map<string, string>();
  const references = parsed.map((entry) => {
    const key = `${entry.sheetName?.toLowerCase() ?? ""}!${entry.bodyAddress}`;
    let color = colorByKey.get(key);
    if (!color) {
      color = FORMULA_REFERENCE_COLORS[colorByKey.size % FORMULA_REFERENCE_COLORS.length] ?? FORMULA_REFERENCE_COLORS[0];
      colorByKey.set(key, color);
    }
    return {
      color,
      end: entry.end,
      kind: entry.kind,
      name: entry.name,
      range: {
        end: { ...entry.range.end },
        start: { ...entry.range.start },
      },
      rangeAddress: entry.bodyAddress,
      sheetName: entry.sheetName,
      start: entry.start,
      text: entry.text,
    } satisfies FormulaReference;
  });

  const safeCaret = Number.isFinite(caret) ? Math.max(0, Math.min(caret, text.length)) : text.length;
  const activeIndex = references.findIndex(
    (reference) => safeCaret >= reference.start && safeCaret <= reference.end,
  );
  return {
    activeIndex,
    activeReference: activeIndex >= 0 ? references[activeIndex] ?? null : null,
    references,
  };
}

/**
 * Parses every cell, range, whole-row/column and defined-name reference in a
 * formula and assigns Excel's editing palette. String literals, quoted sheet
 * names, function calls, external workbook references and unparseable names are
 * left alone. The formula is never mutated.
 */
export function parseFormulaReferences(
  formula: string,
  options: FormulaReferenceOptions = {},
): FormulaReference[] {
  if (!formula.startsWith("=")) return [];
  return collectReferences(formula, formula.length, options).references;
}

/**
 * Parses the formula and marks the reference the caret is inside. A caret just
 * past a reference still counts as inside it, so pressing End in the formula
 * bar keeps the last reference active.
 */
export function collectFormulaReferences(
  formula: string,
  caret: number = formula.length,
  options: FormulaReferenceOptions = {},
): FormulaReferenceCollection {
  if (!formula.startsWith("=")) return EMPTY_COLLECTION;
  return collectReferences(formula, caret, options);
}

/**
 * Resolves a defined name's `refersTo` expression to its first static range.
 * Dynamic formulas such as `OFFSET(...)` resolve to the first literal
 * reference they contain; array constants and `#REF!` resolve to `null`.
 */
export function resolveNamedRangeTarget(
  refersTo: string,
  fallbackSheetName: string | null = null,
): FormulaReferenceNameTarget | null {
  const expression = refersTo.trim().replace(/^=\s*/, "");
  if (!expression || expression.startsWith("{")) return null;
  const first = scanReferenceTokens(expression, {})[0];
  if (!first) return null;
  return {
    range: normalizedReferenceRange(first.range),
    sheetName: first.sheetName ?? fallbackSheetName,
  };
}

function normalizedReferenceRange(range: XlsxCellRange): XlsxCellRange {
  return normalizeRange(range.start, range.end);
}

/** Splits the formula into plain and reference-coloured segments. */
export function formulaReferenceSegments(
  formula: string,
  references: readonly FormulaReference[],
  activeIndex = -1,
): FormulaReferenceSegment[] {
  if (references.length === 0) {
    return formula ? [{ active: false, color: null, reference: null, text: formula }] : [];
  }

  const ordered = references
    .map((reference, index) => ({ index, reference }))
    .sort((left, right) => left.reference.start - right.reference.start);
  const segments: FormulaReferenceSegment[] = [];
  let cursor = 0;
  for (const { index, reference } of ordered) {
    if (reference.start < cursor) continue;
    if (reference.start > cursor) {
      segments.push({
        active: false,
        color: null,
        reference: null,
        text: formula.slice(cursor, reference.start),
      });
    }
    segments.push({
      active: index === activeIndex,
      color: reference.color,
      reference,
      text: formula.slice(reference.start, reference.end),
    });
    cursor = reference.end;
  }
  if (cursor < formula.length) {
    segments.push({
      active: false,
      color: null,
      reference: null,
      text: formula.slice(cursor),
    });
  }
  return segments;
}

/** Palette entry for the n-th distinct reference (wraps after eight). */
export function formulaReferenceColor(index: number): string {
  const size = FORMULA_REFERENCE_COLORS.length;
  const safe = Number.isFinite(index) ? Math.max(0, Math.trunc(index)) : 0;
  return FORMULA_REFERENCE_COLORS[safe % size] ?? FORMULA_REFERENCE_COLORS[0];
}
