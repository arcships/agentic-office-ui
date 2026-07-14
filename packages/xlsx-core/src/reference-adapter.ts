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
} from "./types";

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

export interface XlsxReferenceContext {
  revision: OfficeDocumentRevision & { format: "xlsx" };
  sheets: readonly XlsxReferenceSheet[];
  charts?: readonly XlsxChart[];
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
  if (locator.kind !== "worksheet" && locator.kind !== "range"
    && locator.kind !== "row" && locator.kind !== "column"
    && !(locator.kind === "drawing" && locator.objectType === "chart")) {
    return { status: "unsupported", reasonCode: "xlsx.reference-kind-unsupported" };
  }
  const sheetResult = sheetsForLocator(context, locator.sheet, currentRevision);
  if (sheetResult.sheets.length === 0) return { status: "not-found", reasonCode: "xlsx.sheet-not-found" };
  if (sheetResult.sheets.length > 1) {
    if (locator.kind === "drawing") {
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
    return ambiguousSheets(context, xlsxReference, locator, sheetResult.sheets);
  }
  const sheet = sheetResult.sheets[0]!;

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
