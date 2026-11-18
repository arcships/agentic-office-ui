import {
  XLSX_MAX_COLUMNS,
  XLSX_MAX_ROWS,
  type XlsxCellAddress,
  type XlsxCellRange,
  type XlsxChart,
  type XlsxChartElementSelection,
  type XlsxChartReference,
  type XlsxChartSeries,
  type XlsxRangeHighlight,
} from "@arcships/xlsx-core";

/** Colour per chart data-source role; categories and values stay distinct. */
export const CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE = {
  bubbleSize: "#059669",
  category: "#2563eb",
  seriesName: "#7c3aed",
  value: "#ea580c",
} as const;

export type ChartSourceHighlightRole = keyof typeof CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE;

/** Colour used by `XlsxViewerController.highlightRanges` when none is passed. */
export const XLSX_RANGE_HIGHLIGHT_DEFAULT_COLOR = "#f59e0b";

/** Worksheet identity needed to resolve a sheet-qualified chart formula. */
export interface ChartSourceSheet {
  /** Visible worksheet index used by charts, images and highlights. */
  index: number;
  name: string;
  workbookSheetIndex: number;
}

/** Parsed chart source reference: worksheet index plus a normalized range. */
export interface ChartSourceRange {
  range: XlsxCellRange;
  sheetIndex: number;
}

export interface ParseChartSourceRangeOptions {
  /** Visible sheet used by formulas without a `Sheet!` qualifier. */
  fallbackSheetIndex?: number;
  /** Workbook worksheets, used to resolve a qualified sheet name. */
  sheets?: readonly ChartSourceSheet[];
  /**
   * Resolves a defined name to a sheet-qualified formula. Return `null` for
   * names that do not resolve; chart source ranges from named ranges are
   * otherwise skipped.
   */
  resolveNamedRange?: (name: string) => string | null | undefined;
}

export interface CollectChartRangeHighlightsOptions extends ParseChartSourceRangeOptions {
  /** The selected chart element. Without one there is nothing to highlight. */
  selection?: XlsxChartElementSelection | null;
}

function assertCell(cell: XlsxCellAddress): void {
  if (
    !Number.isSafeInteger(cell.row)
    || !Number.isSafeInteger(cell.col)
    || cell.row < 0
    || cell.col < 0
  ) {
    throw new TypeError("XLSX highlight ranges require non-negative safe integer cells");
  }
}

/** Returns a normalized copy of `range` and rejects invalid coordinates. */
export function normalizeHighlightRange(range: XlsxCellRange): XlsxCellRange {
  assertCell(range.start);
  assertCell(range.end);
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

function columnIndex(label: string): number {
  let value = 0;
  for (const character of label.toUpperCase()) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value - 1;
}

function parseCellPart(part: string): XlsxCellAddress | null {
  const match = /^([A-Z]{1,3})(\d{1,7})$/iu.exec(part);
  if (!match) return null;
  const col = columnIndex(match[1]!);
  const row = Number(match[2]) - 1;
  if (
    !Number.isSafeInteger(row)
    || !Number.isSafeInteger(col)
    || row < 0
    || col < 0
    || row >= XLSX_MAX_ROWS
    || col >= XLSX_MAX_COLUMNS
  ) {
    return null;
  }
  return { col, row };
}

function parseColumnPart(part: string): number | null {
  if (!/^[A-Z]{1,3}$/iu.test(part)) return null;
  const col = columnIndex(part);
  return col >= 0 && col < XLSX_MAX_COLUMNS ? col : null;
}

function parseRowPart(part: string): number | null {
  if (!/^\d{1,7}$/u.test(part)) return null;
  const row = Number(part) - 1;
  return row >= 0 && row < XLSX_MAX_ROWS ? row : null;
}

/** Parses `A1`, `A1:B2`, `A:A` or `1:1` (with optional `$` anchors). */
export function parseChartSourceA1Range(reference: string): XlsxCellRange | null {
  const cleaned = reference.trim().replace(/\$/g, "");
  if (!cleaned) return null;
  const parts = cleaned.split(":");
  if (parts.length > 2 || parts.some((part) => part.length === 0)) return null;
  const [startPart, endPart = startPart] = parts;
  if (!startPart || !endPart) return null;

  const startCell = parseCellPart(startPart);
  const endCell = parseCellPart(endPart);
  if (startCell && endCell) {
    return normalizeHighlightRange({ end: endCell, start: startCell });
  }

  const startColumn = parseColumnPart(startPart);
  const endColumn = parseColumnPart(endPart);
  if (startColumn !== null && endColumn !== null) {
    return normalizeHighlightRange({
      end: { col: endColumn, row: XLSX_MAX_ROWS - 1 },
      start: { col: startColumn, row: 0 },
    });
  }

  const startRow = parseRowPart(startPart);
  const endRow = parseRowPart(endPart);
  if (startRow !== null && endRow !== null) {
    return normalizeHighlightRange({
      end: { col: XLSX_MAX_COLUMNS - 1, row: endRow },
      start: { col: 0, row: startRow },
    });
  }

  return null;
}

function unquoteSheetName(value: string): string {
  const trimmed = value.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

function splitSheetQualifier(reference: string): { range: string; sheetName?: string } {
  let quoted = false;
  for (let index = 0; index < reference.length; index += 1) {
    const character = reference[index];
    if (character === "'") {
      quoted = !quoted;
    } else if (character === "!" && !quoted) {
      return {
        range: reference.slice(index + 1),
        sheetName: unquoteSheetName(reference.slice(0, index)),
      };
    }
  }
  return { range: reference };
}

function resolveSheetIndex(
  sheetName: string | undefined,
  options: ParseChartSourceRangeOptions,
): number | null {
  if (!sheetName) {
    return options.fallbackSheetIndex ?? null;
  }
  if (!options.sheets?.length) {
    return null;
  }
  const needle = sheetName.toLocaleLowerCase();
  const match = options.sheets.find((sheet) => sheet.name.toLocaleLowerCase() === needle);
  return match ? match.index : null;
}

/**
 * Parses a chart reference formula such as `Sheet1!$B$2:$B$5` into the visible
 * sheet index and cell range it points at.
 *
 * Quoted sheet names, `$` anchors, single cells, ranges, whole columns and
 * whole rows are accepted. External-workbook references (`[Book.xlsx]…`),
 * unions and unparseable formulas return `null` instead of highlighting the
 * wrong cells. A formula without a sheet qualifier uses
 * `fallbackSheetIndex`; a defined name is resolved through
 * `resolveNamedRange`.
 */
export function parseChartSourceRange(
  formula: string | null | undefined,
  options: ParseChartSourceRangeOptions = {},
): ChartSourceRange | null {
  return parseChartSourceRangeInternal(formula, options, new Set());
}

function parseChartSourceRangeInternal(
  formula: string | null | undefined,
  options: ParseChartSourceRangeOptions,
  seenNames: Set<string>,
): ChartSourceRange | null {
  let reference = (formula ?? "").trim();
  if (!reference) return null;
  if (reference.startsWith("=")) reference = reference.slice(1).trim();
  if (!reference || reference.includes("[")) return null;

  const split = splitSheetQualifier(reference);
  if (split.range.includes(",") || (split.range.startsWith("(") && split.range.endsWith(")"))) {
    return null;
  }

  const range = parseChartSourceA1Range(split.range);
  if (range) {
    const sheetIndex = resolveSheetIndex(split.sheetName, options);
    return sheetIndex === null ? null : { range, sheetIndex };
  }

  if (!split.sheetName && options.resolveNamedRange && !seenNames.has(reference)) {
    seenNames.add(reference);
    const resolved = options.resolveNamedRange(reference);
    if (resolved && resolved !== reference && resolved !== formula) {
      return parseChartSourceRangeInternal(resolved, options, seenNames);
    }
  }
  return null;
}

function seriesIndexForSelection(
  chart: XlsxChart,
  selection: XlsxChartElementSelection,
): number {
  if (selection.kind === "chart") return -1;
  const byId = chart.series.findIndex((series) => series.id === selection.seriesId);
  if (byId >= 0) return byId;
  return selection.seriesIndex >= 0 && selection.seriesIndex < chart.series.length
    ? selection.seriesIndex
    : -1;
}

function selectedSeriesIndexes(
  chart: XlsxChart,
  selection: XlsxChartElementSelection,
): number[] {
  if (selection.kind === "chart") {
    return chart.series.map((_, index) => index);
  }
  const index = seriesIndexForSelection(chart, selection);
  return index >= 0 ? [index] : [];
}

/** Maps a series point index onto the cell of `range` it represents. */
export function chartRangeCellAt(range: XlsxCellRange, pointIndex: number): XlsxCellAddress | null {
  if (!Number.isSafeInteger(pointIndex) || pointIndex < 0) return null;
  const normalized = normalizeHighlightRange(range);
  const rowSpan = normalized.end.row - normalized.start.row + 1;
  const colSpan = normalized.end.col - normalized.start.col + 1;
  if (pointIndex >= rowSpan && pointIndex >= colSpan) return null;

  if (rowSpan === 1) {
    return { col: normalized.start.col + pointIndex, row: normalized.start.row };
  }
  if (colSpan === 1) {
    return { col: normalized.start.col, row: normalized.start.row + pointIndex };
  }
  // A rectangular source is usually a category/value block with one point per
  // row; prefer the row axis, matching how Excel lays out chart data.
  return { col: normalized.start.col, row: normalized.start.row + pointIndex };
}

function referenceHighlights(
  chart: XlsxChart,
  series: XlsxChartSeries,
  seriesIndex: number,
  selection: XlsxChartElementSelection,
  options: CollectChartRangeHighlightsOptions,
): XlsxRangeHighlight[] {
  const fallbackSheetIndex = options.fallbackSheetIndex ?? chart.sheetIndex ?? 0;
  const references: Array<{
    reference: XlsxChartReference | null | undefined;
    role: ChartSourceHighlightRole;
  }> = [
    { reference: series.categoriesRef, role: "category" },
    { reference: series.valuesRef, role: "value" },
    { reference: series.bubbleSizeRef, role: "bubbleSize" },
  ];
  const result: XlsxRangeHighlight[] = [];

  for (const { reference, role } of references) {
    if (!reference?.formula) continue;
    const source = parseChartSourceRange(reference.formula, {
      ...options,
      fallbackSheetIndex,
    });
    if (!source) continue;

    let range = source.range;
    if (selection.kind === "point") {
      const cell = chartRangeCellAt(range, selection.pointIndex);
      if (!cell) continue;
      range = { end: cell, start: cell };
    }

    const seriesKey = series.id || String(seriesIndex);
    result.push({
      color: CHART_SOURCE_HIGHLIGHT_COLOR_BY_ROLE[role],
      id: [
        chart.id,
        selection.kind,
        selection.kind === "point" ? selection.pointIndex : "all",
        seriesKey,
        role,
        range.start.row,
        range.start.col,
        range.end.row,
        range.end.col,
      ].join(":"),
      range,
      role,
      seriesIndex,
      seriesName: series.name,
      sheetIndex: source.sheetIndex,
    });
  }

  return result;
}

/**
 * Derives the source ranges a selected chart element reads from.
 *
 * A chart-level selection highlights every series; a series, point or legend
 * entry selection highlights only that series (a point narrows category and
 * value references to the point's own cell). Category, value and bubble-size
 * ranges keep their own colours so a legend entry and its source are easy to
 * match. Duplicate ranges (several series sharing categories) are emitted once.
 */
export function collectChartRangeHighlights(
  chart: XlsxChart,
  options: CollectChartRangeHighlightsOptions = {},
): XlsxRangeHighlight[] {
  const { selection } = options;
  if (!selection || selection.chartId !== chart.id) return [];

  const result: XlsxRangeHighlight[] = [];
  const seen = new Set<string>();
  for (const seriesIndex of selectedSeriesIndexes(chart, selection)) {
    const series = chart.series[seriesIndex];
    if (!series) continue;
    for (const highlight of referenceHighlights(chart, series, seriesIndex, selection, options)) {
      const key = [
        highlight.sheetIndex,
        highlight.range.start.row,
        highlight.range.start.col,
        highlight.range.end.row,
        highlight.range.end.col,
        highlight.role,
      ].join(":");
      if (seen.has(key)) continue;
      seen.add(key);
      result.push(highlight);
    }
  }
  return result;
}
