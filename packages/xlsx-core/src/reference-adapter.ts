import type {
  ManualRegionLocator,
  OfficeDocumentRevision,
  OfficeObjectDescriptor,
  OfficeObjectFingerprint,
  OfficeObjectReference,
  OfficeObjectReferenceDraft,
  OfficeObjectReliability,
  ResolveReferenceResult,
  XlsxObjectLocator,
  XlsxSheetLocator,
} from "@arcships/office-interaction";
import type {
  XlsxCellAddress,
  XlsxCellRange,
  XlsxChart,
  XlsxChartSeries,
} from "./types";
import type { XlsxBehaviorRecord, XlsxBehaviorScope } from "./xlsx-behaviors";

export type XlsxOfficeReference = Extract<
  OfficeObjectReference,
  { document: { format: "xlsx" } }
>;

export type XlsxOfficeReferenceDraft = Extract<
  OfficeObjectReferenceDraft,
  { document: { format: "xlsx" } }
>;

/** Stable worksheet identity copied from an XLSX workbook tab. */
export interface XlsxReferenceSheet {
  id?: string;
  index: number;
  name: string;
  workbookSheetIndex?: number;
}

export interface XlsxReferenceCellSnapshot {
  displayValue?: string;
  formula?: string;
}

/** One workbook table part visible on a worksheet. */
export interface XlsxReferenceTable {
  name: string;
  displayName?: string;
  sheet: XlsxReferenceSheet;
  /** Inclusive A1 range of the table body, headers included. */
  a1: string;
  id?: string;
}

/** One drawing object anchored on a worksheet. */
export interface XlsxReferenceDrawing {
  id: string;
  objectType: "image" | "shape" | "form-control";
  sheet: XlsxReferenceSheet;
  name?: string;
  alt?: string;
}

/** One cell comment. */
export interface XlsxReferenceComment {
  sheet: XlsxReferenceSheet;
  cellA1: string;
  text?: string;
  author?: string;
  resolved?: boolean;
  nativeId?: string;
}

export interface XlsxReferenceContext {
  revision: OfficeDocumentRevision & { format: "xlsx" };
  sheets: readonly XlsxReferenceSheet[];
  charts?: readonly XlsxChart[];
  tables?: readonly XlsxReferenceTable[];
  drawings?: readonly XlsxReferenceDrawing[];
  comments?: readonly XlsxReferenceComment[];
  behaviors?: readonly XlsxBehaviorRecord[];
  /** Optional visible/model cell evidence used in descriptors and fingerprints. */
  getCellSnapshot?: (
    sheet: XlsxReferenceSheet,
    cell: XlsxCellAddress,
  ) => XlsxReferenceCellSnapshot | undefined;
}

const TEXT_QUOTE_LIMIT = 256;

const XLSX_SHEET_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["xlsx.native-worksheet"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["xlsx.worksheet-boundary"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.workbook-tab"] },
  relocation: { level: "likely", reasonCodes: ["xlsx.sheet-identity"] },
};

const XLSX_RANGE_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["xlsx.native-cell-range"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["xlsx.a1-range"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.sheet-identity"] },
  relocation: { level: "likely", reasonCodes: ["xlsx.sheet-and-a1"] },
};

const XLSX_CHART_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["xlsx.native-chart"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["xlsx.drawing-id"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.sheet-identity"] },
  relocation: { level: "likely", reasonCodes: ["xlsx.chart-id-or-name"] },
};

const XLSX_TABLE_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["xlsx.native-table"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["xlsx.table-a1"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.table-sheet"] },
  relocation: { level: "likely", reasonCodes: ["xlsx.table-name"] },
};

const XLSX_DRAWING_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["xlsx.native-drawing"] },
  boundary: { level: "likely", reasonCodes: ["xlsx.drawing-anchor"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.drawing-sheet"] },
  relocation: { level: "likely", reasonCodes: ["xlsx.drawing-id-or-name"] },
};

const XLSX_COMMENT_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["xlsx.native-comment"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["xlsx.comment-cell"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.comment-sheet"] },
  relocation: { level: "uncertain", reasonCodes: ["xlsx.comment-cell"] },
};

const XLSX_CHART_ELEMENT_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["xlsx.native-chart-element"] },
  boundary: { level: "likely", reasonCodes: ["xlsx.chart-plot-geometry"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.chart-element-index"] },
  relocation: { level: "likely", reasonCodes: ["xlsx.chart-series-name"] },
};

const XLSX_BEHAVIOR_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "exact", score: 1, reasonCodes: ["xlsx.native-behavior"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["xlsx.behavior-scope"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.sheet-identity"] },
  relocation: { level: "likely", reasonCodes: ["xlsx.behavior-instance-id"] },
};

const XLSX_REGION_RELIABILITY: OfficeObjectReliability = {
  semantic: { level: "unknown", reasonCodes: ["manual-region"] },
  boundary: { level: "exact", score: 1, reasonCodes: ["xlsx.sheet-region"] },
  hierarchy: { level: "exact", score: 1, reasonCodes: ["xlsx.sheet-identity"] },
  relocation: { level: "likely", reasonCodes: ["xlsx.sheet-and-cell-offsets"] },
};

function copyReliability(value: OfficeObjectReliability): OfficeObjectReliability {
  return {
    semantic: { ...value.semantic, reasonCodes: [...value.semantic.reasonCodes] },
    boundary: { ...value.boundary, reasonCodes: [...value.boundary.reasonCodes] },
    hierarchy: { ...value.hierarchy, reasonCodes: [...value.hierarchy.reasonCodes] },
    relocation: { ...value.relocation, reasonCodes: [...value.relocation.reasonCodes] },
  };
}

function textContentHash(text: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    hash ^= code & 0xff;
    hash = Math.imul(hash, 0x01000193);
    hash ^= code >>> 8;
    hash = Math.imul(hash, 0x01000193);
  }
  return `fnv1a32:${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

function sameRevision(left: OfficeDocumentRevision, right: OfficeDocumentRevision): boolean {
  return left.format === right.format
    && left.documentId === right.documentId
    && left.revision === right.revision;
}

function assertContext(context: XlsxReferenceContext): void {
  if (context.revision.format !== "xlsx") throw new TypeError("XLSX reference context requires format=xlsx");
  if (!context.revision.documentId.trim()) throw new TypeError("XLSX documentId must not be empty");
  if (!context.revision.revision.trim()) throw new TypeError("XLSX revision must not be empty");
  const indices = new Set<number>();
  for (const sheet of context.sheets) {
    if (!Number.isSafeInteger(sheet.index) || sheet.index < 0 || indices.has(sheet.index)) {
      throw new TypeError("XLSX reference sheets require unique non-negative indices");
    }
    if (!sheet.name.trim()) throw new TypeError("XLSX reference sheet names must not be empty");
    if (sheet.id !== undefined && !sheet.id.trim()) throw new TypeError("XLSX reference sheet ids must not be empty");
    indices.add(sheet.index);
  }
}

function sheetLocator(sheet: XlsxReferenceSheet): XlsxSheetLocator {
  return {
    ...(sheet.id ? { sheetId: sheet.id } : {}),
    name: sheet.name,
    index: sheet.index,
  };
}

function sheetKey(sheet: XlsxReferenceSheet): string {
  return sheet.id ?? sheet.name;
}

function sameSheet(left: XlsxReferenceSheet, right: XlsxReferenceSheet): boolean {
  if (left.id && right.id) return left.id === right.id;
  return left.name === right.name && left.index === right.index;
}

function requireContextSheet(context: XlsxReferenceContext, requested: XlsxReferenceSheet): XlsxReferenceSheet {
  const found = context.sheets.find((sheet) => sameSheet(sheet, requested));
  if (!found) throw new RangeError("XLSX worksheet is not present in the reference context");
  return found;
}

function columnLabel(column: number): string {
  let value = column + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function assertCell(cell: XlsxCellAddress): void {
  if (!Number.isSafeInteger(cell.row) || cell.row < 0 || !Number.isSafeInteger(cell.col) || cell.col < 0) {
    throw new TypeError("XLSX cell coordinates must be non-negative safe integers");
  }
}

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  assertCell(range.start);
  assertCell(range.end);
  return {
    start: { row: Math.min(range.start.row, range.end.row), col: Math.min(range.start.col, range.end.col) },
    end: { row: Math.max(range.start.row, range.end.row), col: Math.max(range.start.col, range.end.col) },
  };
}

function rangeA1(range: XlsxCellRange): string {
  const normalized = normalizeRange(range);
  const start = `${columnLabel(normalized.start.col)}${normalized.start.row + 1}`;
  const end = `${columnLabel(normalized.end.col)}${normalized.end.row + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function columnIndex(label: string): number {
  let value = 0;
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64;
  return value - 1;
}

function rangeFromA1(a1: string): XlsxCellRange | undefined {
  const match = /^\$?([A-Z]{1,4})\$?([1-9][0-9]*)(?::\$?([A-Z]{1,4})\$?([1-9][0-9]*))?$/iu.exec(a1);
  if (!match) return undefined;
  const start = { col: columnIndex(match[1]!), row: Number(match[2]) - 1 };
  const end = { col: columnIndex(match[3] ?? match[1]!), row: Number(match[4] ?? match[2]) - 1 };
  if (![start.col, start.row, end.col, end.row].every(Number.isSafeInteger)) return undefined;
  return normalizeRange({ start, end });
}

function snapshotFingerprint(
  sheet: XlsxReferenceSheet,
  a1: string,
  snapshot?: XlsxReferenceCellSnapshot,
): OfficeObjectFingerprint {
  const displayValue = snapshot?.displayValue;
  const formula = snapshot?.formula;
  const signature = displayValue !== undefined || formula !== undefined
    ? `value:${displayValue ?? ""}\nformula:${formula ?? ""}`
    : undefined;
  return {
    ...(displayValue && displayValue.length <= TEXT_QUOTE_LIMIT ? { exactText: displayValue } : {}),
    ...(signature ? { contentHash: textContentHash(signature) } : {}),
    ancestorKeys: [`xlsx.sheet:${sheetKey(sheet)}`, `xlsx.a1:${a1}`],
  };
}

function chartFingerprint(sheet: XlsxReferenceSheet, chart: XlsxChart): OfficeObjectFingerprint {
  const name = chart.name?.trim() || chart.title?.trim();
  return {
    ...(name ? { objectName: name } : {}),
    contentHash: textContentHash(`chart:${chart.chartType}:${chart.series.length}:${chart.title ?? ""}`),
    ancestorKeys: [
      `xlsx.sheet:${sheetKey(sheet)}`,
      `xlsx.chart-id:${chart.id}`,
      `xlsx.chart-type:${chart.chartType}`,
    ],
  };
}

function assertSheetRegion(
  sheet: XlsxReferenceSheet,
  region: Extract<ManualRegionLocator, { space: "sheet" }>,
): void {
  if (region.sheetId !== sheetKey(sheet)) {
    throw new RangeError("XLSX region sheetId must match the selected worksheet identity");
  }
  for (const point of [region.start, region.end]) {
    assertCell(point);
    if (![point.xOffset, point.yOffset].every(Number.isFinite)
      || point.xOffset < 0 || point.xOffset > 1 || point.yOffset < 0 || point.yOffset > 1) {
      throw new RangeError("XLSX region offsets must be finite values between zero and one");
    }
  }
  if (region.start.row > region.end.row || region.start.col > region.end.col
    || region.start.row === region.end.row && region.start.yOffset >= region.end.yOffset
    || region.start.col === region.end.col && region.start.xOffset >= region.end.xOffset) {
    throw new RangeError("XLSX region must describe a positive forward sheet area");
  }
}

export function createXlsxWorksheetReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, requestedSheet);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "worksheet",
    source: "native",
    locator: { type: "format", format: "xlsx", value: { kind: "worksheet", sheet: sheetLocator(sheet) } },
    fingerprint: { objectName: sheet.name, ancestorKeys: [`xlsx.sheet:${sheetKey(sheet)}`] },
    reliability: copyReliability(XLSX_SHEET_RELIABILITY),
  };
}

export function createXlsxCellReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
  cell: XlsxCellAddress,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, requestedSheet);
  assertCell(cell);
  const a1 = rangeA1({ start: cell, end: cell });
  const snapshot = context.getCellSnapshot?.(sheet, cell);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "cell",
    source: "native",
    locator: { type: "format", format: "xlsx", value: { kind: "range", sheet: sheetLocator(sheet), a1 } },
    fingerprint: snapshotFingerprint(sheet, a1, snapshot),
    reliability: copyReliability(XLSX_RANGE_RELIABILITY),
  };
}

export function createXlsxRangeReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
  range: XlsxCellRange,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, requestedSheet);
  const a1 = rangeA1(range);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "cell-range",
    source: "native",
    locator: { type: "format", format: "xlsx", value: { kind: "range", sheet: sheetLocator(sheet), a1 } },
    fingerprint: { ancestorKeys: [`xlsx.sheet:${sheetKey(sheet)}`, `xlsx.a1:${a1}`] },
    reliability: copyReliability(XLSX_RANGE_RELIABILITY),
  };
}

function assertAxisRange(start: number, end: number, axis: "row" | "column"): { start: number; end: number } {
  if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(end) || end < 0) {
    throw new TypeError(`XLSX ${axis} indices must be non-negative safe integers`);
  }
  return { start: Math.min(start, end), end: Math.max(start, end) };
}

function createXlsxAxisReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
  kind: "row" | "column",
  start: number,
  end: number,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, requestedSheet);
  const range = assertAxisRange(start, end, kind);
  const label = kind === "row"
    ? `${range.start + 1}:${range.end + 1}`
    : `${columnLabel(range.start)}:${columnLabel(range.end)}`;
  return {
    schemaVersion: 1,
    document: context.revision,
    kind,
    source: "native",
    locator: {
      type: "format",
      format: "xlsx",
      value: { kind, sheet: sheetLocator(sheet), ...range },
    },
    fingerprint: { ancestorKeys: [`xlsx.sheet:${sheetKey(sheet)}`, `xlsx.${kind}:${label}`] },
    reliability: copyReliability(XLSX_RANGE_RELIABILITY),
  };
}

/** Create a stable reference to one row or an inclusive row span. Indices are zero-based. */
export function createXlsxRowReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
  start: number,
  end = start,
): XlsxOfficeReferenceDraft {
  return createXlsxAxisReferenceDraft(context, requestedSheet, "row", start, end);
}

/** Create a stable reference to one column or an inclusive column span. Indices are zero-based. */
export function createXlsxColumnReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
  start: number,
  end = start,
): XlsxOfficeReferenceDraft {
  return createXlsxAxisReferenceDraft(context, requestedSheet, "column", start, end);
}

export function createXlsxChartReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
  chart: XlsxChart,
  options: { fallbackRegion?: Extract<ManualRegionLocator, { space: "sheet" }> } = {},
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, requestedSheet);
  const currentChart = context.charts?.find((candidate) => candidate.id === chart.id
    && chartBelongsToSheet(candidate, sheet));
  if (!currentChart) throw new RangeError("XLSX chart is not present in the reference context");
  if (options.fallbackRegion) assertSheetRegion(sheet, options.fallbackRegion);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "chart",
    source: "native",
    locator: {
      type: "format",
      format: "xlsx",
      value: {
        kind: "drawing",
        sheet: sheetLocator(sheet),
        objectType: "chart",
        id: currentChart.id,
        ...(currentChart.name ? { name: currentChart.name } : {}),
      },
    },
    fingerprint: chartFingerprint(sheet, currentChart),
    ...(options.fallbackRegion ? { fallbackRegion: options.fallbackRegion } : {}),
    reliability: copyReliability(XLSX_CHART_RELIABILITY),
  };
}

function tableFingerprint(table: XlsxReferenceTable): OfficeObjectFingerprint {
  const name = table.name.trim();
  return {
    ...(name ? { objectName: name } : {}),
    contentHash: textContentHash(`table:${name}:${table.a1}`),
    ancestorKeys: [
      `xlsx.sheet:${sheetKey(table.sheet)}`,
      `xlsx.table:${name}`,
      `xlsx.a1:${table.a1}`,
    ],
  };
}

/** Create a stable reference to a workbook table part. */
export function createXlsxTableReferenceDraft(
  context: XlsxReferenceContext,
  table: XlsxReferenceTable,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, table.sheet);
  const candidates = (context.tables ?? []).filter((candidate) =>
    candidate.name === table.name && sameSheet(candidate.sheet, sheet));
  if (candidates.length === 0) throw new RangeError("XLSX table is not present in the reference context");
  if (candidates.length > 1) throw new RangeError("XLSX table name is ambiguous in the reference context");
  const current = candidates[0]!;
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "table",
    source: "native",
    locator: {
      type: "format",
      format: "xlsx",
      value: {
        kind: "table",
        sheet: sheetLocator(sheet),
        tableName: current.name,
        ...(current.id ? { id: current.id } : {}),
      },
    },
    fingerprint: tableFingerprint({ ...current, sheet }),
    reliability: copyReliability(XLSX_TABLE_RELIABILITY),
  };
}

function drawingFingerprint(drawing: XlsxReferenceDrawing): OfficeObjectFingerprint {
  const name = drawing.name?.trim() || drawing.alt?.trim();
  return {
    ...(name ? { objectName: name } : {}),
    contentHash: textContentHash(`drawing:${drawing.objectType}:${drawing.id}`),
    ancestorKeys: [
      `xlsx.sheet:${sheetKey(drawing.sheet)}`,
      `xlsx.drawing:${drawing.objectType}:${drawing.id}`,
    ],
  };
}

/** Create a stable reference to an image, shape or form control. */
export function createXlsxDrawingReferenceDraft(
  context: XlsxReferenceContext,
  drawing: XlsxReferenceDrawing,
  options: { fallbackRegion?: Extract<ManualRegionLocator, { space: "sheet" }> } = {},
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, drawing.sheet);
  const current = (context.drawings ?? []).find((candidate) =>
    candidate.id === drawing.id
    && candidate.objectType === drawing.objectType
    && sameSheet(candidate.sheet, sheet));
  if (!current) throw new RangeError("XLSX drawing is not present in the reference context");
  if (options.fallbackRegion) assertSheetRegion(sheet, options.fallbackRegion);
  const kind = current.objectType === "form-control"
    ? "form-control"
    : current.objectType === "image"
      ? "image"
      : "shape";
  return {
    schemaVersion: 1,
    document: context.revision,
    kind,
    source: "native",
    locator: {
      type: "format",
      format: "xlsx",
      value: {
        kind: "drawing",
        sheet: sheetLocator(sheet),
        objectType: current.objectType,
        id: current.id,
        ...(current.name ? { name: current.name } : {}),
      },
    },
    fingerprint: drawingFingerprint({ ...current, sheet }),
    ...(options.fallbackRegion ? { fallbackRegion: options.fallbackRegion } : {}),
    reliability: copyReliability(XLSX_DRAWING_RELIABILITY),
  };
}

/** Create a stable reference to one cell comment. */
export function createXlsxCommentReferenceDraft(
  context: XlsxReferenceContext,
  comment: XlsxReferenceComment,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, comment.sheet);
  const a1 = comment.cellA1.trim().toUpperCase();
  const current = (context.comments ?? []).find((candidate) =>
    candidate.cellA1.toUpperCase() === a1
    && sameSheet(candidate.sheet, sheet)
    && (comment.nativeId === undefined || candidate.nativeId === comment.nativeId));
  if (!current) throw new RangeError("XLSX comment is not present in the reference context");
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "comment",
    source: "native",
    locator: {
      type: "format",
      format: "xlsx",
      value: {
        kind: "comment",
        sheet: sheetLocator(sheet),
        cellA1: a1,
        ...(current.nativeId ? { nativeId: current.nativeId } : {}),
      },
    },
    fingerprint: {
      ...(current.text ? { contentHash: textContentHash(current.text) } : {}),
      ancestorKeys: [`xlsx.sheet:${sheetKey(sheet)}`, `xlsx.comment-cell:${a1}`],
    },
    reliability: copyReliability(XLSX_COMMENT_RELIABILITY),
  };
}

export type XlsxChartElementTarget =
  | { element: "series"; seriesIndex: number }
  | { element: "point"; seriesIndex: number; pointIndex: number }
  | { element: "legend-entry"; legendIndex: number; seriesIndex?: number };

const XLSX_CHART_ELEMENT_KIND = {
  series: "chart-series",
  point: "chart-point",
  "legend-entry": "chart-legend-entry",
} as const;

function seriesFingerprint(
  sheet: XlsxReferenceSheet,
  chart: XlsxChart,
  series: XlsxChartSeries,
  seriesIndex: number,
): OfficeObjectFingerprint {
  const name = series.name?.trim();
  const signature = `series:${name ?? seriesIndex}:${series.values.length}:${series.values.slice(0, 8).join(",")}`;
  return {
    ...(name ? { objectName: name } : {}),
    contentHash: textContentHash(signature),
    ancestorKeys: [
      `xlsx.sheet:${sheetKey(sheet)}`,
      `xlsx.chart-id:${chart.id}`,
      `xlsx.chart-element:series:${seriesIndex}`,
    ],
  };
}

function assertChartElementTarget(chart: XlsxChart, target: XlsxChartElementTarget): void {
  const index = target.element === "legend-entry" ? target.legendIndex : target.seriesIndex;
  if (!Number.isSafeInteger(index) || index < 0) {
    throw new TypeError("XLSX chart element indices must be non-negative safe integers");
  }
  if (target.element === "legend-entry") {
    if (index >= chart.series.length) {
      throw new RangeError("XLSX chart legend entry is outside the current chart");
    }
    if (target.seriesIndex !== undefined
      && (!Number.isSafeInteger(target.seriesIndex) || target.seriesIndex < 0 || target.seriesIndex >= chart.series.length)) {
      throw new RangeError("XLSX chart legend series is outside the current chart");
    }
    return;
  }
  const series = chart.series[target.seriesIndex];
  if (!series) throw new RangeError("XLSX chart series is outside the current chart");
  if (target.element === "point") {
    if (!Number.isSafeInteger(target.pointIndex) || target.pointIndex < 0) {
      throw new TypeError("XLSX chart point index must be a non-negative safe integer");
    }
    if (series.values.length > 0 && target.pointIndex >= series.values.length) {
      throw new RangeError("XLSX chart point is outside the current series");
    }
  }
}

/**
 * Create a reference to an element inside a chart (stage 3).
 *
 * Series, point and legend-entry references reuse the chart's drawing identity,
 * so a chart that moves on the sheet keeps the same internal references.
 */
export function createXlsxChartElementReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
  chartId: string,
  target: XlsxChartElementTarget,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, requestedSheet);
  const chart = (context.charts ?? []).find((candidate) =>
    candidate.id === chartId && chartBelongsToSheet(candidate, sheet));
  if (!chart) throw new RangeError("XLSX chart is not present in the reference context");
  assertChartElementTarget(chart, target);
  const seriesIndex = target.element === "legend-entry"
    ? target.seriesIndex ?? target.legendIndex
    : target.seriesIndex;
  const series = chart.series[seriesIndex]!;
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: XLSX_CHART_ELEMENT_KIND[target.element],
    source: "native",
    locator: {
      type: "format",
      format: "xlsx",
      value: {
        kind: "chart-element",
        sheet: sheetLocator(sheet),
        chartId: chart.id,
        ...target,
      },
    },
    fingerprint: seriesFingerprint(sheet, chart, series, seriesIndex),
    reliability: copyReliability(XLSX_CHART_ELEMENT_RELIABILITY),
  };
}

function behaviorLocator(record: XlsxBehaviorRecord): XlsxObjectLocator {
  if (record.behavior === "named-range" || record.behavior === "protection") {
    const scope = record.scope.kind === "workbook"
      ? { kind: "workbook" as const }
      : { kind: "sheet" as const, sheet: (record.scope as Extract<XlsxBehaviorScope, { kind: "sheet" }>).sheet };
    return { kind: "behavior", behavior: record.behavior, scope, instanceId: record.instanceId };
  }
  if (record.scope.kind !== "sheet") {
    throw new RangeError(`XLSX ${record.behavior} requires a sheet scope`);
  }
  return {
    kind: "behavior",
    behavior: record.behavior,
    scope: {
      kind: "sheet",
      sheet: record.scope.sheet,
      appliesToA1: (record.scope as Extract<XlsxBehaviorScope, { kind: "sheet"; appliesToA1: string }>).appliesToA1,
    },
    instanceId: record.instanceId,
  };
}

/** Create a stage 4 invisible-behavior reference (formula, rule, name, protection). */
export function createXlsxBehaviorReferenceDraft(
  context: XlsxReferenceContext,
  record: XlsxBehaviorRecord,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const current = (context.behaviors ?? []).find((candidate) => candidate.instanceId === record.instanceId);
  if (!current) throw new RangeError("XLSX behavior is not present in the reference context");
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: XLSX_BEHAVIOR_KIND[current.behavior],
    source: "native",
    locator: { type: "format", format: "xlsx", value: behaviorLocator(current) },
    fingerprint: {
      ancestorKeys: [
        `xlsx.behavior:${current.behavior}`,
        `xlsx.behavior-instance:${current.instanceId}`,
      ],
    },
    reliability: copyReliability(XLSX_BEHAVIOR_RELIABILITY),
  };
}

const XLSX_BEHAVIOR_KIND = {
  formula: "formula",
  "conditional-format": "conditional-format",
  "data-validation": "data-validation",
  "named-range": "named-range",
  protection: "protection",
} as const;

function referenceFromDraft(draft: XlsxOfficeReferenceDraft): XlsxOfficeReference {
  return { ...draft, referenceId: "descriptor" } as XlsxOfficeReference;
}

/** Describe a table without committing a reference (object-mode hover/list). */
export function describeXlsxTable(
  context: XlsxReferenceContext,
  table: XlsxReferenceTable,
): OfficeObjectDescriptor {
  return tableDescriptor(referenceFromDraft(createXlsxTableReferenceDraft(context, table)), table.sheet, table);
}

/** Describe a drawing object without committing a reference. */
export function describeXlsxDrawing(
  context: XlsxReferenceContext,
  drawing: XlsxReferenceDrawing,
): OfficeObjectDescriptor {
  return drawingDescriptor(referenceFromDraft(createXlsxDrawingReferenceDraft(context, drawing)), drawing.sheet, drawing);
}

/** Describe one cell comment without committing a reference. */
export function describeXlsxComment(
  context: XlsxReferenceContext,
  comment: XlsxReferenceComment,
): OfficeObjectDescriptor {
  return commentDescriptor(referenceFromDraft(createXlsxCommentReferenceDraft(context, comment)), comment.sheet, comment);
}

/** Describe a chart without committing a reference. */
export function describeXlsxChart(
  context: XlsxReferenceContext,
  sheet: XlsxReferenceSheet,
  chart: XlsxChart,
): OfficeObjectDescriptor {
  return chartDescriptor(referenceFromDraft(createXlsxChartReferenceDraft(context, sheet, chart)), sheet, chart);
}

/** Describe one chart element without committing a reference. */
export function describeXlsxChartElement(
  context: XlsxReferenceContext,
  sheet: XlsxReferenceSheet,
  chart: XlsxChart,
  target: XlsxChartElementTarget,
): OfficeObjectDescriptor {
  const draft = createXlsxChartElementReferenceDraft(context, sheet, chart.id, target);
  const seriesIndex = target.element === "legend-entry"
    ? target.seriesIndex ?? target.legendIndex
    : target.seriesIndex;
  const series = chart.series[seriesIndex]!;
  return chartElementDescriptor(referenceFromDraft(draft), sheet, chart, target, series);
}

/** Describe one stage 4 behavior without committing a reference. */
export function describeXlsxBehavior(
  context: XlsxReferenceContext,
  record: XlsxBehaviorRecord,
): OfficeObjectDescriptor {
  return behaviorDescriptor(referenceFromDraft(createXlsxBehaviorReferenceDraft(context, record)), record);
}

export function createXlsxRegionReferenceDraft(
  context: XlsxReferenceContext,
  requestedSheet: XlsxReferenceSheet,
  region: Extract<ManualRegionLocator, { space: "sheet" }>,
): XlsxOfficeReferenceDraft {
  assertContext(context);
  const sheet = requireContextSheet(context, requestedSheet);
  assertSheetRegion(sheet, region);
  return {
    schemaVersion: 1,
    document: context.revision,
    kind: "region",
    source: "manual",
    locator: { type: "manual-region", format: "xlsx", value: region },
    fingerprint: { ancestorKeys: [`xlsx.sheet:${sheetKey(sheet)}`] },
    fallbackRegion: region,
    reliability: copyReliability(XLSX_REGION_RELIABILITY),
  };
}

function draftFromReference(reference: XlsxOfficeReference): XlsxOfficeReferenceDraft {
  const { referenceId: _referenceId, ...draft } = reference;
  return draft;
}

function currentFormatReference(
  context: XlsxReferenceContext,
  reference: XlsxOfficeReference,
  locator: XlsxObjectLocator,
  fingerprint: OfficeObjectFingerprint | undefined = reference.fingerprint,
): XlsxOfficeReference {
  return {
    ...reference,
    document: context.revision,
    locator: { type: "format", format: "xlsx", value: locator },
    ...(fingerprint ? { fingerprint } : {}),
  } as XlsxOfficeReference;
}

function currentRegionReference(
  context: XlsxReferenceContext,
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
): XlsxOfficeReference {
  if (reference.locator.type !== "manual-region") return reference;
  const value = { ...reference.locator.value, sheetId: sheetKey(sheet) };
  return {
    ...reference,
    document: context.revision,
    locator: { type: "manual-region", format: "xlsx", value },
    fallbackRegion: value,
  } as XlsxOfficeReference;
}

function sheetsForLocator(
  context: XlsxReferenceContext,
  locator: XlsxSheetLocator,
  currentRevision: boolean,
): { sheets: readonly XlsxReferenceSheet[]; reasonCode: string } {
  if (locator.sheetId) {
    const matches = context.sheets.filter((sheet) => sheet.id === locator.sheetId);
    if (matches.length) return { sheets: matches, reasonCode: "xlsx.sheet-id-match" };
  }
  const nameMatches = context.sheets.filter((sheet) => sheet.name === locator.name);
  if (nameMatches.length) return { sheets: nameMatches, reasonCode: "xlsx.sheet-name-match" };
  if (currentRevision) {
    const indexMatch = context.sheets.filter((sheet) => sheet.index === locator.index);
    if (indexMatch.length) return { sheets: indexMatch, reasonCode: "xlsx.sheet-index-match" };
  }
  return { sheets: [], reasonCode: "xlsx.sheet-not-found" };
}

function sheetForRegion(
  context: XlsxReferenceContext,
  region: Extract<ManualRegionLocator, { space: "sheet" }>,
): readonly XlsxReferenceSheet[] {
  return context.sheets.filter((sheet) => sheetKey(sheet) === region.sheetId);
}

function snapshotForRange(
  context: XlsxReferenceContext,
  sheet: XlsxReferenceSheet,
  range: XlsxCellRange,
): XlsxReferenceCellSnapshot | undefined {
  if (range.start.row !== range.end.row || range.start.col !== range.end.col) return undefined;
  return context.getCellSnapshot?.(sheet, range.start);
}

function rangeDescriptor(
  context: XlsxReferenceContext,
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  a1: string,
  range: XlsxCellRange,
): OfficeObjectDescriptor {
  const singleCell = range.start.row === range.end.row && range.start.col === range.end.col;
  const snapshot = snapshotForRange(context, sheet, range);
  const visual = reference.fallbackRegion
    ? {
        fragments: [{
          container: { space: "sheet" as const, sheetId: sheetKey(sheet) },
          region: reference.fallbackRegion,
        }],
        layoutVersion: reference.document.revision,
      }
    : undefined;
  return {
    objectId: `xlsx:${sheetKey(sheet)}:${a1}`,
    draft: draftFromReference(reference),
    label: `${sheet.name}!${a1}`,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
      { kind: singleCell ? "cell" : "cell-range", label: a1 },
    ],
    childrenState: "none",
    ...((snapshot?.displayValue !== undefined || snapshot?.formula !== undefined) ? {
      content: {
        ...(snapshot.displayValue !== undefined ? { value: snapshot.displayValue } : {}),
        ...(snapshot.formula !== undefined ? { formula: snapshot.formula } : {}),
      },
    } : {}),
    ...(visual ? { visual } : {}),
    facets: { sheetId: sheetKey(sheet), a1 },
  };
}

function axisDescriptor(
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  locator: Extract<XlsxObjectLocator, { kind: "row" | "column" }>,
): OfficeObjectDescriptor {
  const label = locator.kind === "row"
    ? `${locator.start + 1}:${locator.end + 1}`
    : `${columnLabel(locator.start)}:${columnLabel(locator.end)}`;
  return {
    objectId: `xlsx:${sheetKey(sheet)}:${locator.kind}:${locator.start}:${locator.end}`,
    draft: draftFromReference(reference),
    label: `${sheet.name}!${label}`,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
      { kind: locator.kind, label },
    ],
    childrenState: "available",
    facets: {
      sheetId: sheetKey(sheet),
      axis: locator.kind,
      start: locator.start,
      end: locator.end,
    },
  };
}

function worksheetDescriptor(reference: XlsxOfficeReference, sheet: XlsxReferenceSheet): OfficeObjectDescriptor {
  return {
    objectId: `xlsx:sheet:${sheetKey(sheet)}`,
    draft: draftFromReference(reference),
    label: sheet.name,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
    ],
    childrenState: "available",
    facets: { sheetId: sheetKey(sheet), sheetIndex: sheet.index },
  };
}

function chartBelongsToSheet(chart: XlsxChart, sheet: XlsxReferenceSheet): boolean {
  if (sheet.workbookSheetIndex !== undefined) return chart.workbookSheetIndex === sheet.workbookSheetIndex;
  return chart.sheetIndex === sheet.index;
}

function chartDescriptor(
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  chart: XlsxChart,
): OfficeObjectDescriptor {
  const name = chart.name?.trim() || chart.title?.trim() || "Chart";
  const visual = reference.fallbackRegion
    ? {
        fragments: [{
          container: { space: "sheet" as const, sheetId: sheetKey(sheet) },
          region: reference.fallbackRegion,
          zIndex: chart.zIndex,
        }],
        layoutVersion: reference.document.revision,
      }
    : undefined;
  return {
    objectId: `xlsx:chart:${sheetKey(sheet)}:${chart.id}`,
    draft: draftFromReference(reference),
    label: name,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
      { kind: "chart", label: name },
    ],
    childrenState: chart.series.length ? "available" : "none",
    ...(visual ? { visual } : {}),
    facets: {
      sheetId: sheetKey(sheet),
      chartId: chart.id,
      chartType: chart.chartType,
      seriesCount: chart.series.length,
    },
  };
}

function tableDescriptor(
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  table: XlsxReferenceTable,
): OfficeObjectDescriptor {
  return {
    objectId: `xlsx:table:${sheetKey(sheet)}:${table.name}`,
    draft: draftFromReference(reference),
    label: table.displayName?.trim() || table.name,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
      { kind: "table", label: table.name },
    ],
    childrenState: "available",
    facets: {
      sheetId: sheetKey(sheet),
      tableName: table.name,
      a1: table.a1,
      ...(table.id ? { tableId: table.id } : {}),
    },
  };
}

function drawingDescriptor(
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  drawing: XlsxReferenceDrawing,
): OfficeObjectDescriptor {
  const kind = drawing.objectType === "form-control"
    ? "form-control"
    : drawing.objectType === "image" ? "image" : "shape";
  const name = drawing.name?.trim() || drawing.alt?.trim();
  const visual = reference.fallbackRegion
    ? {
        fragments: [{
          container: { space: "sheet" as const, sheetId: sheetKey(sheet) },
          region: reference.fallbackRegion,
        }],
        layoutVersion: reference.document.revision,
      }
    : undefined;
  return {
    objectId: `xlsx:drawing:${sheetKey(sheet)}:${drawing.objectType}:${drawing.id}`,
    draft: draftFromReference(reference),
    label: name || `${drawing.objectType} ${drawing.id}`,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
      { kind, label: name || drawing.objectType },
    ],
    childrenState: "none",
    ...(drawing.alt ? { content: { text: drawing.alt } } : {}),
    ...(visual ? { visual } : {}),
    facets: { sheetId: sheetKey(sheet), drawingId: drawing.id, objectType: drawing.objectType },
  };
}

function commentDescriptor(
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  comment: XlsxReferenceComment,
): OfficeObjectDescriptor {
  const text = comment.text?.replace(/\s+/gu, " ").trim();
  return {
    objectId: `xlsx:comment:${sheetKey(sheet)}:${comment.cellA1}`,
    draft: draftFromReference(reference),
    label: text ? `Comment: ${text.slice(0, 80)}` : `Comment ${comment.cellA1}`,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
      { kind: "cell", label: comment.cellA1 },
      { kind: "comment", label: comment.author ? `Comment by ${comment.author}` : "Comment" },
    ],
    childrenState: "none",
    ...(comment.text ? { content: { text: comment.text } } : {}),
    facets: {
      sheetId: sheetKey(sheet),
      cellA1: comment.cellA1,
      ...(comment.author ? { author: comment.author } : {}),
      ...(comment.resolved === undefined ? {} : { resolved: comment.resolved }),
    },
  };
}

function chartElementDescriptor(
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  chart: XlsxChart,
  target: XlsxChartElementTarget,
  series: XlsxChartSeries,
): OfficeObjectDescriptor {
  const seriesName = series.name?.trim() || `Series ${target.element === "legend-entry" ? target.legendIndex + 1 : target.seriesIndex + 1}`;
  const label = target.element === "series"
    ? `Series: ${seriesName}`
    : target.element === "point"
      ? `Point ${target.pointIndex + 1}`
      : `Legend entry ${target.legendIndex + 1}`;
  const kind = XLSX_CHART_ELEMENT_KIND[target.element];
  const pointIndex = target.element === "point" ? target.pointIndex : undefined;
  const value = pointIndex === undefined ? undefined : series.values[pointIndex];
  return {
    objectId: `xlsx:chart-element:${sheetKey(sheet)}:${chart.id}:${target.element}:${pointIndex ?? (target.element === "legend-entry" ? target.legendIndex : target.seriesIndex)}`,
    draft: draftFromReference(reference),
    label,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
      { kind: "chart", label: chart.name?.trim() || chart.title?.trim() || "Chart" },
      { kind, label },
    ],
    childrenState: target.element === "series" ? "available" : "none",
    ...(value !== undefined && value !== null ? { content: { value: String(value) } } : {}),
    facets: {
      sheetId: sheetKey(sheet),
      chartId: chart.id,
      seriesIndex: target.element === "legend-entry" ? target.seriesIndex ?? target.legendIndex : target.seriesIndex,
      ...(pointIndex === undefined ? {} : { pointIndex }),
      ...(target.element === "legend-entry" ? { legendIndex: target.legendIndex } : {}),
      seriesName,
    },
  };
}

function behaviorDescriptor(
  reference: XlsxOfficeReference,
  record: XlsxBehaviorRecord,
): OfficeObjectDescriptor {
  return {
    objectId: `xlsx:behavior:${record.instanceId}`,
    draft: draftFromReference(reference),
    label: record.label,
    path: record.path.map((entry) => ({ ...entry })),
    childrenState: "none",
    ...(record.content === undefined ? {} : { content: { ...record.content } }),
    facets: { ...record.facets, behavior: record.behavior, instanceId: record.instanceId },
  };
}

function regionDescriptor(reference: XlsxOfficeReference, sheet: XlsxReferenceSheet): OfficeObjectDescriptor {
  const region = reference.locator.type === "manual-region" ? reference.locator.value : reference.fallbackRegion!;
  return {
    objectId: `xlsx:region:${sheetKey(sheet)}:${JSON.stringify(region)}`,
    draft: draftFromReference(reference),
    label: `Region on ${sheet.name}`,
    path: [
      { kind: "workbook", label: "Workbook" },
      { kind: "worksheet", label: sheet.name },
      { kind: "region", label: "Selected region" },
    ],
    childrenState: "none",
    visual: {
      fragments: [{ container: { space: "sheet", sheetId: sheetKey(sheet) }, region }],
      layoutVersion: reference.document.revision,
    },
  };
}

function ambiguousSheets(
  context: XlsxReferenceContext,
  reference: XlsxOfficeReference,
  locator: Extract<XlsxObjectLocator, { kind: "worksheet" | "range" | "row" | "column" }>,
  sheets: readonly XlsxReferenceSheet[],
): ResolveReferenceResult {
  const candidates = sheets.flatMap((sheet) => {
    if (locator.kind === "worksheet") {
      const current = currentFormatReference(context, reference, { ...locator, sheet: sheetLocator(sheet) });
      return [worksheetDescriptor(current, sheet)];
    }
    const current = currentFormatReference(context, reference, { ...locator, sheet: sheetLocator(sheet) });
    if ("start" in locator) {
      return [axisDescriptor(current, sheet, { ...locator, sheet: sheetLocator(sheet) })];
    }
    const range = rangeFromA1(locator.a1);
    if (!range) return [];
    return [rangeDescriptor(context, current, sheet, locator.a1, range)];
  });
  return { status: "ambiguous", candidates, reasonCodes: ["xlsx.multiple-sheet-matches"] };
}

function tableOnSheet(
  context: XlsxReferenceContext,
  sheet: XlsxReferenceSheet,
  locator: Extract<XlsxObjectLocator, { kind: "table" }>,
): readonly XlsxReferenceTable[] {
  return (context.tables ?? []).filter((table) =>
    sameSheet(table.sheet, sheet)
    && (table.name === locator.tableName || (locator.id !== undefined && table.id === locator.id)));
}

function drawingOnSheet(
  context: XlsxReferenceContext,
  sheet: XlsxReferenceSheet,
  locator: Extract<XlsxObjectLocator, { kind: "drawing" }>,
): readonly XlsxReferenceDrawing[] {
  const exact = (context.drawings ?? []).filter((drawing) =>
    sameSheet(drawing.sheet, sheet)
    && drawing.objectType === locator.objectType
    && drawing.id === locator.id);
  if (exact.length > 0) return exact;
  const name = locator.name?.trim();
  if (!name) return [];
  return (context.drawings ?? []).filter((drawing) =>
    sameSheet(drawing.sheet, sheet)
    && drawing.objectType === locator.objectType
    && (drawing.name?.trim() === name || drawing.alt?.trim() === name));
}

function commentOnSheet(
  context: XlsxReferenceContext,
  sheet: XlsxReferenceSheet,
  locator: Extract<XlsxObjectLocator, { kind: "comment" }>,
): readonly XlsxReferenceComment[] {
  const a1 = locator.cellA1.toUpperCase();
  return (context.comments ?? []).filter((comment) =>
    sameSheet(comment.sheet, sheet)
    && comment.cellA1.toUpperCase() === a1
    && (locator.nativeId === undefined || comment.nativeId === locator.nativeId));
}

function chartElementTarget(
  locator: Extract<XlsxObjectLocator, { kind: "chart-element" }>,
): XlsxChartElementTarget {
  if (locator.element === "series") return { element: "series", seriesIndex: locator.seriesIndex };
  if (locator.element === "point") {
    return { element: "point", seriesIndex: locator.seriesIndex, pointIndex: locator.pointIndex };
  }
  return {
    element: "legend-entry",
    legendIndex: locator.legendIndex,
    ...(locator.seriesIndex === undefined ? {} : { seriesIndex: locator.seriesIndex }),
  };
}

function chartElementForLocator(
  context: XlsxReferenceContext,
  sheet: XlsxReferenceSheet,
  locator: Extract<XlsxObjectLocator, { kind: "chart-element" }>,
): { chart: XlsxChart; target: XlsxChartElementTarget; series: XlsxChartSeries; seriesIndex: number } | undefined {
  const chart = (context.charts ?? []).find((candidate) =>
    candidate.id === locator.chartId && chartBelongsToSheet(candidate, sheet));
  if (!chart) return undefined;
  const target = chartElementTarget(locator);
  try {
    assertChartElementTarget(chart, target);
  } catch {
    return undefined;
  }
  const seriesIndex = target.element === "legend-entry"
    ? target.seriesIndex ?? target.legendIndex
    : target.seriesIndex;
  const series = chart.series[seriesIndex];
  if (!series) return undefined;
  return { chart, target, series, seriesIndex };
}

function resolveXlsxTableOrDrawingOrComment(
  context: XlsxReferenceContext,
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  currentRevision: boolean,
): ResolveReferenceResult | undefined {
  const locator = reference.locator.type === "format" ? reference.locator.value : undefined;
  if (!locator) return undefined;
  if (locator.kind === "table") {
    const matches = tableOnSheet(context, sheet, locator);
    if (matches.length === 0) return { status: "not-found", reasonCode: "xlsx.table-not-found" };
    if (matches.length > 1) {
      return {
        status: "ambiguous",
        candidates: matches.map((table) => tableDescriptor(
          currentFormatReference(context, reference, { ...locator, sheet: sheetLocator(sheet) }, tableFingerprint({ ...table, sheet })),
          sheet,
          table,
        )),
        reasonCodes: ["xlsx.multiple-table-matches"],
      };
    }
    const table = matches[0]!;
    const fingerprint = tableFingerprint({ ...table, sheet });
    if (currentRevision) {
      return { status: "exact", reference, descriptor: tableDescriptor(reference, sheet, table) };
    }
    const relocated = currentFormatReference(context, reference, {
      ...locator,
      sheet: sheetLocator(sheet),
    }, fingerprint);
    return {
      status: "relocated",
      reference: relocated,
      descriptor: tableDescriptor(relocated, sheet, table),
      reasonCodes: ["xlsx.revision-changed", "xlsx.table-name-match"],
    };
  }
  if (locator.kind === "drawing") {
    const matches = drawingOnSheet(context, sheet, locator);
    if (matches.length === 0) return { status: "not-found", reasonCode: "xlsx.drawing-not-found" };
    if (matches.length > 1) {
      return {
        status: "ambiguous",
        candidates: matches.map((drawing) => drawingDescriptor(
          currentFormatReference(context, reference, {
            ...locator,
            sheet: sheetLocator(sheet),
            id: drawing.id,
          }, drawingFingerprint({ ...drawing, sheet })),
          sheet,
          drawing,
        )),
        reasonCodes: ["xlsx.multiple-drawing-matches"],
      };
    }
    const drawing = matches[0]!;
    if (currentRevision) return { status: "exact", reference, descriptor: drawingDescriptor(reference, sheet, drawing) };
    const relocated = currentFormatReference(context, reference, {
      ...locator,
      sheet: sheetLocator(sheet),
      id: drawing.id,
    }, drawingFingerprint({ ...drawing, sheet }));
    return {
      status: "relocated",
      reference: relocated,
      descriptor: drawingDescriptor(relocated, sheet, drawing),
      reasonCodes: ["xlsx.revision-changed", "xlsx.drawing-identity-match"],
    };
  }
  if (locator.kind === "comment") {
    const matches = commentOnSheet(context, sheet, locator);
    if (matches.length === 0) return { status: "not-found", reasonCode: "xlsx.comment-not-found" };
    const comment = matches[0]!;
    if (currentRevision) return { status: "exact", reference, descriptor: commentDescriptor(reference, sheet, comment) };
    const relocated = currentFormatReference(context, reference, { ...locator, sheet: sheetLocator(sheet) });
    return {
      status: "relocated",
      reference: relocated,
      descriptor: commentDescriptor(relocated, sheet, comment),
      reasonCodes: ["xlsx.revision-changed", "xlsx.comment-cell-match"],
    };
  }
  return undefined;
}

function resolveXlsxChartElement(
  context: XlsxReferenceContext,
  reference: XlsxOfficeReference,
  sheet: XlsxReferenceSheet,
  locator: Extract<XlsxObjectLocator, { kind: "chart-element" }>,
  currentRevision: boolean,
): ResolveReferenceResult {
  const found = chartElementForLocator(context, sheet, locator);
  if (!found) return { status: "not-found", reasonCode: "xlsx.chart-element-not-found" };
  const fingerprint = seriesFingerprint(sheet, found.chart, found.series, found.seriesIndex);
  if (currentRevision) {
    return {
      status: "exact",
      reference,
      descriptor: chartElementDescriptor(reference, sheet, found.chart, found.target, found.series),
    };
  }
  const relocated = currentFormatReference(context, reference, {
    ...locator,
    chartId: found.chart.id,
    sheet: sheetLocator(sheet),
    ...found.target,
  }, fingerprint);
  return {
    status: "relocated",
    reference: relocated,
    descriptor: chartElementDescriptor(relocated, sheet, found.chart, found.target, found.series),
    reasonCodes: ["xlsx.revision-changed", "xlsx.chart-series-name-match"],
  };
}

function resolveXlsxBehavior(
  context: XlsxReferenceContext,
  reference: XlsxOfficeReference,
  locator: Extract<XlsxObjectLocator, { kind: "behavior" }>,
  currentRevision: boolean,
): ResolveReferenceResult {
  const record = (context.behaviors ?? []).find((candidate) => candidate.instanceId === locator.instanceId);
  if (!record || record.behavior !== locator.behavior) {
    return { status: "not-found", reasonCode: "xlsx.behavior-not-found" };
  }
  if (currentRevision) {
    return { status: "exact", reference, descriptor: behaviorDescriptor(reference, record) };
  }
  const relocated = currentFormatReference(context, reference, behaviorLocator(record));
  return {
    status: "relocated",
    reference: relocated,
    descriptor: behaviorDescriptor(relocated, record),
    reasonCodes: ["xlsx.revision-changed", "xlsx.behavior-instance-id"],
  };
}

export function resolveXlsxReference(
  context: XlsxReferenceContext,
  reference: OfficeObjectReference,
): ResolveReferenceResult {
  assertContext(context);
  if (reference.document.format !== "xlsx") return { status: "unsupported", reasonCode: "xlsx.format-mismatch" };
  const xlsxReference = reference as XlsxOfficeReference;
  if (xlsxReference.document.documentId !== context.revision.documentId) {
    return { status: "not-found", reasonCode: "xlsx.document-id-mismatch" };
  }
  const currentRevision = sameRevision(xlsxReference.document, context.revision);

  if (xlsxReference.locator.type === "manual-region") {
    const sheets = sheetForRegion(context, xlsxReference.locator.value);
    if (sheets.length === 0) return { status: "not-found", reasonCode: "xlsx.region-sheet-not-found" };
    if (sheets.length > 1) {
      return {
        status: "ambiguous",
        candidates: sheets.map((sheet) => regionDescriptor(currentRegionReference(context, xlsxReference, sheet), sheet)),
        reasonCodes: ["xlsx.multiple-region-sheet-matches"],
      };
    }
    if (currentRevision) {
      return { status: "exact", reference: xlsxReference, descriptor: regionDescriptor(xlsxReference, sheets[0]!) };
    }
    const relocated = currentRegionReference(context, xlsxReference, sheets[0]!);
    return {
      status: "relocated",
      reference: relocated,
      descriptor: regionDescriptor(relocated, sheets[0]!),
      reasonCodes: ["xlsx.revision-changed", "xlsx.sheet-region-stable"],
    };
  }

  const locator = xlsxReference.locator.value;

  // Workbook-scoped behaviors have no worksheet; resolve them before any
  // sheet lookup so a missing sheet cannot hide a named range or protection.
  if (locator.kind === "behavior" && locator.scope.kind === "workbook") {
    return resolveXlsxBehavior(context, xlsxReference, locator, currentRevision);
  }

  const sheetLocatorForReference = locator.kind === "behavior"
    ? locator.scope.kind === "sheet" ? locator.scope.sheet : undefined
    : "sheet" in locator ? locator.sheet : undefined;
  if (!sheetLocatorForReference) {
    return { status: "unsupported", reasonCode: "xlsx.locator-sheet-unavailable" };
  }
  const sheetResult = sheetsForLocator(context, sheetLocatorForReference, currentRevision);
  if (sheetResult.sheets.length === 0) return { status: "not-found", reasonCode: "xlsx.sheet-not-found" };
  if (sheetResult.sheets.length > 1) {
    if (locator.kind === "drawing" && locator.objectType === "chart") {
      const candidates = sheetResult.sheets.flatMap((sheet) => (context.charts ?? [])
        .filter((chart) => chartBelongsToSheet(chart, sheet) && chart.id === locator.id)
        .map((chart) => {
          const current = currentFormatReference(context, xlsxReference, {
            ...locator,
            sheet: sheetLocator(sheet),
          }, chartFingerprint(sheet, chart));
          return chartDescriptor(current, sheet, chart);
        }));
      return { status: "ambiguous", candidates, reasonCodes: ["xlsx.multiple-sheet-matches"] };
    }
    if (locator.kind === "table" || locator.kind === "drawing" || locator.kind === "comment") {
      const candidates: OfficeObjectDescriptor[] = [];
      for (const candidateSheet of sheetResult.sheets) {
        const descriptorReference = currentFormatReference(context, xlsxReference, {
          ...locator,
          sheet: sheetLocator(candidateSheet),
        });
        if (locator.kind === "table") {
          const table = tableOnSheet(context, candidateSheet, locator)[0];
          if (table) candidates.push(tableDescriptor(descriptorReference, candidateSheet, table));
        } else if (locator.kind === "drawing") {
          const drawing = drawingOnSheet(context, candidateSheet, locator)[0];
          if (drawing) candidates.push(drawingDescriptor(descriptorReference, candidateSheet, drawing));
        } else {
          const comment = commentOnSheet(context, candidateSheet, locator)[0];
          if (comment) candidates.push(commentDescriptor(descriptorReference, candidateSheet, comment));
        }
      }
      return { status: "ambiguous", candidates, reasonCodes: ["xlsx.multiple-sheet-matches"] };
    }
    if (locator.kind === "chart-element" || locator.kind === "behavior") {
      return { status: "ambiguous", candidates: [], reasonCodes: ["xlsx.multiple-sheet-matches"] };
    }
    if (locator.kind !== "worksheet" && locator.kind !== "range"
      && locator.kind !== "row" && locator.kind !== "column") {
      return { status: "unsupported", reasonCode: "xlsx.reference-kind-unsupported" };
    }
    return ambiguousSheets(context, xlsxReference, locator, sheetResult.sheets);
  }
  const sheet = sheetResult.sheets[0]!;

  if (locator.kind === "table" || locator.kind === "comment"
    || (locator.kind === "drawing" && locator.objectType !== "chart")) {
    const result = resolveXlsxTableOrDrawingOrComment(context, xlsxReference, sheet, currentRevision);
    if (result) return result;
  }
  if (locator.kind === "chart-element") {
    return resolveXlsxChartElement(context, xlsxReference, sheet, locator, currentRevision);
  }
  if (locator.kind === "behavior") {
    return resolveXlsxBehavior(context, xlsxReference, locator, currentRevision);
  }

  if (locator.kind === "worksheet") {
    const current = currentFormatReference(context, xlsxReference, { ...locator, sheet: sheetLocator(sheet) });
    if (currentRevision) return { status: "exact", reference: xlsxReference, descriptor: worksheetDescriptor(xlsxReference, sheet) };
    return {
      status: "relocated",
      reference: current,
      descriptor: worksheetDescriptor(current, sheet),
      reasonCodes: ["xlsx.revision-changed", sheetResult.reasonCode],
    };
  }

  if (locator.kind === "range") {
    const range = rangeFromA1(locator.a1);
    if (!range) return { status: "unsupported", reasonCode: "xlsx.invalid-a1-range" };
    const current = currentFormatReference(context, xlsxReference, {
      ...locator,
      sheet: sheetLocator(sheet),
    }, xlsxReference.kind === "cell"
      ? snapshotFingerprint(sheet, locator.a1, snapshotForRange(context, sheet, range))
      : xlsxReference.fingerprint);
    if (currentRevision) {
      return { status: "exact", reference: xlsxReference, descriptor: rangeDescriptor(context, xlsxReference, sheet, locator.a1, range) };
    }
    return {
      status: "relocated",
      reference: current,
      descriptor: rangeDescriptor(context, current, sheet, locator.a1, range),
      reasonCodes: ["xlsx.revision-changed", sheetResult.reasonCode, "xlsx.a1-stable"],
    };
  }

  if ("start" in locator) {
    let axis: { start: number; end: number };
    try {
      axis = assertAxisRange(locator.start, locator.end, locator.kind);
    } catch {
      return { status: "unsupported", reasonCode: `xlsx.invalid-${locator.kind}-range` };
    }
    const currentLocator = { ...locator, ...axis, sheet: sheetLocator(sheet) };
    const current = currentFormatReference(context, xlsxReference, currentLocator);
    if (currentRevision) {
      return { status: "exact", reference: xlsxReference, descriptor: axisDescriptor(xlsxReference, sheet, currentLocator) };
    }
    return {
      status: "relocated",
      reference: current,
      descriptor: axisDescriptor(current, sheet, currentLocator),
      reasonCodes: ["xlsx.revision-changed", sheetResult.reasonCode, `xlsx.${locator.kind}-stable`],
    };
  }

  if (locator.kind !== "drawing") {
    return { status: "unsupported", reasonCode: "xlsx.reference-kind-unsupported" };
  }
  const sheetCharts = (context.charts ?? []).filter((chart) => chartBelongsToSheet(chart, sheet));
  let charts = sheetCharts.filter((chart) => chart.id === locator.id);
  let chartReason = "xlsx.chart-id-match";
  if (charts.length === 0 && !currentRevision) {
    const name = xlsxReference.fingerprint?.objectName ?? locator.name;
    const typeKey = xlsxReference.fingerprint?.ancestorKeys?.find((key) => key.startsWith("xlsx.chart-type:"));
    const chartType = typeKey?.slice("xlsx.chart-type:".length);
    if (name) {
      charts = sheetCharts.filter((chart) => (chart.name === name || chart.title === name)
        && (!chartType || chart.chartType === chartType));
      chartReason = "xlsx.chart-name-type-match";
    }
  }
  if (charts.length === 0) return { status: "not-found", reasonCode: "xlsx.chart-not-found" };
  const chartReferences = charts.map((chart) => currentFormatReference(context, xlsxReference, {
    kind: "drawing",
    sheet: sheetLocator(sheet),
    objectType: "chart",
    id: chart.id,
    ...(chart.name ? { name: chart.name } : {}),
  }, chartFingerprint(sheet, chart)));
  if (charts.length > 1) {
    return {
      status: "ambiguous",
      candidates: charts.map((chart, index) => chartDescriptor(chartReferences[index]!, sheet, chart)),
      reasonCodes: ["xlsx.multiple-chart-matches"],
    };
  }
  if (currentRevision) {
    return { status: "exact", reference: xlsxReference, descriptor: chartDescriptor(xlsxReference, sheet, charts[0]!) };
  }
  return {
    status: "relocated",
    reference: chartReferences[0]!,
    descriptor: chartDescriptor(chartReferences[0]!, sheet, charts[0]!),
    reasonCodes: ["xlsx.revision-changed", sheetResult.reasonCode, chartReason],
  };
}

export function describeXlsxReference(
  context: XlsxReferenceContext,
  reference: OfficeObjectReference,
): OfficeObjectDescriptor | undefined {
  const result = resolveXlsxReference(context, reference);
  return result.status === "exact" || result.status === "relocated" ? result.descriptor : undefined;
}
