/**
 * Platform-neutral XLSX data entry.
 *
 * XML parsing, Worker/WASM setup, object URLs, rendering, and process-wide
 * caches belong to explicit adapters and must not enter this dependency graph.
 */

import type { XlsxChart } from "./types/chart-types";
import type { XlsxImageAnchor } from "./types/image-types";

export type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartDataLabels,
  XlsxChartElementSelection,
  XlsxChartLegend,
  XlsxChartPointDataLabel,
  XlsxChartPointStyle,
  XlsxChartReference,
  XlsxChartSeries,
  XlsxChartsheet,
  XlsxChartTypeGroup,
  XlsxChartWall,
} from "./types/chart-types";

export type {
  XlsxFormControl,
  XlsxFormControlKind,
  XlsxImage,
  XlsxImageAnchor,
  XlsxImageMarker,
  XlsxImageRect,
  XlsxImageResizeHandlePosition,
  XlsxShape,
  XlsxShapeFill,
  XlsxShapeParagraph,
  XlsxShapeStroke,
  XlsxShapeTextBox,
  XlsxShapeTextRun,
} from "./types/image-types";

export {
  resolveWorkbookColor,
  resolveWorkbookFillColor,
  resolveWorkbookFillStyle,
} from "./colors";

export const EMU_PER_PIXEL = 9525;
const DEFAULT_COLUMN_WIDTH_EMU = 64 * EMU_PER_PIXEL;
const DEFAULT_ROW_HEIGHT_EMU = 20 * EMU_PER_PIXEL;

export interface DrawingRectEmu {
  cx: number;
  cy: number;
  x: number;
  y: number;
}

export interface GroupTransform {
  chX: number;
  chY: number;
  scaleX: number;
  scaleY: number;
  x: number;
  y: number;
}

export interface CellAddress {
  col: number;
  row: number;
}

export interface CellRange {
  end: CellAddress;
  start: CellAddress;
}

export type ParsedChartSeriesFormula = {
  bubbleSizeFormula?: string;
  categoryFormula: string;
  nameFormula?: string;
  nameLiteral?: string;
  order: number;
  valueFormula: string;
};

export function columnLabel(col: number): string {
  let label = "";
  let nextValue = col;
  while (nextValue >= 0) {
    label = String.fromCharCode(65 + (nextValue % 26)) + label;
    nextValue = Math.floor(nextValue / 26) - 1;
  }
  return label;
}

export function rangeToA1(range: CellRange): string {
  const startRow = Math.min(range.start.row, range.end.row);
  const startCol = Math.min(range.start.col, range.end.col);
  const endRow = Math.max(range.start.row, range.end.row);
  const endCol = Math.max(range.start.col, range.end.col);
  const start = `${columnLabel(startCol)}${startRow + 1}`;
  const end = `${columnLabel(endCol)}${endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

export function buildA1RangeFormula(
  sheetName: string,
  start: CellAddress,
  end: CellAddress,
): string {
  const escapedSheetName = sheetName.replace(/'/g, "''");
  return `'${escapedSheetName}'!$${columnLabel(start.col)}$${start.row + 1}:$${columnLabel(end.col)}$${end.row + 1}`;
}

function quoteSeriesFormulaString(value: string) {
  return `"${value.replace(/"/g, "\"\"")}"`;
}

function unquoteSeriesFormulaString(value: string) {
  const trimmed = value.trim();
  if (trimmed.length < 2 || !trimmed.startsWith("\"") || !trimmed.endsWith("\"")) {
    return null;
  }
  return trimmed.slice(1, -1).replace(/""/g, "\"");
}

function splitTopLevelSeriesArguments(value: string) {
  const args: string[] = [];
  let current = "";
  let doubleQuoted = false;
  let singleQuoted = false;
  let depth = 0;

  for (let index = 0; index < value.length; index += 1) {
    const char = value[index] ?? "";
    const next = value[index + 1] ?? "";
    if (doubleQuoted) {
      current += char;
      if (char === "\"" && next === "\"") {
        current += next;
        index += 1;
      } else if (char === "\"") {
        doubleQuoted = false;
      }
      continue;
    }
    if (singleQuoted) {
      current += char;
      if (char === "'" && next === "'") {
        current += next;
        index += 1;
      } else if (char === "'") {
        singleQuoted = false;
      }
      continue;
    }
    if (char === "\"") {
      doubleQuoted = true;
      current += char;
      continue;
    }
    if (char === "'") {
      singleQuoted = true;
      current += char;
      continue;
    }
    if (char === "(") {
      depth += 1;
      current += char;
      continue;
    }
    if (char === ")") {
      depth = Math.max(0, depth - 1);
      current += char;
      continue;
    }
    if (char === "," && depth === 0) {
      args.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }
  args.push(current.trim());
  return args;
}

export function buildChartSeriesFormula(
  chart: XlsxChart | null | undefined,
  seriesIndex: number,
): string {
  const series = chart?.series[seriesIndex];
  if (!chart || !series) return "";
  const raw = series.raw && typeof series.raw === "object"
    ? series.raw as Record<string, unknown>
    : null;
  const rawName = typeof raw?.name === "string" && raw.name.length > 0
    ? raw.name
    : null;
  const nameArgument = rawName
    ?? quoteSeriesFormulaString(series.name ?? `Series ${seriesIndex + 1}`);
  const bubbleArgument = series.bubbleSizeRef?.formula;
  return [
    `=SERIES(${nameArgument}`,
    series.categoriesRef?.formula ?? "",
    series.valuesRef?.formula ?? "",
    String(seriesIndex + 1),
    ...(chart.chartType === "Bubble" || bubbleArgument ? [bubbleArgument ?? ""] : []),
  ].join(",") + ")";
}

export function parseChartSeriesFormula(
  formula: string,
  chart: XlsxChart | null | undefined,
): ParsedChartSeriesFormula | null {
  const trimmed = formula.trim();
  const withoutEquals = trimmed.startsWith("=") ? trimmed.slice(1).trim() : trimmed;
  const match = /^SERIES\s*\(([\s\S]*)\)$/i.exec(withoutEquals);
  if (!match) return null;
  const args = splitTopLevelSeriesArguments(match[1]);
  const isBubble = chart?.chartType === "Bubble";
  if (args.length < 4 || args.length > 5 || (isBubble && args.length !== 5)) {
    return null;
  }
  const [nameArg = "", categoryFormula = "", valueFormula = "", orderArg = "", bubbleSizeFormula] = args;
  if (!categoryFormula || !valueFormula) return null;
  const order = Number(orderArg);
  if (!Number.isFinite(order)) return null;
  const nameLiteral = unquoteSeriesFormulaString(nameArg);
  return {
    bubbleSizeFormula: bubbleSizeFormula && bubbleSizeFormula.length > 0
      ? bubbleSizeFormula
      : undefined,
    categoryFormula,
    nameFormula: nameLiteral == null && nameArg.length > 0 ? nameArg : undefined,
    nameLiteral: nameLiteral ?? undefined,
    order,
    valueFormula,
  };
}

export function anchorToRect(anchor: XlsxImageAnchor): DrawingRectEmu {
  if (anchor.kind === "absolute") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: anchor.positionEmu.x,
      y: anchor.positionEmu.y,
    };
  }
  if (anchor.kind === "one-cell") {
    return {
      cx: anchor.sizeEmu.cx,
      cy: anchor.sizeEmu.cy,
      x: anchor.from.colOffsetEmu,
      y: anchor.from.rowOffsetEmu,
    };
  }
  return {
    cx: Math.max(
      0,
      (anchor.to.col - anchor.from.col) * DEFAULT_COLUMN_WIDTH_EMU
        + anchor.to.colOffsetEmu
        - anchor.from.colOffsetEmu,
    ),
    cy: Math.max(
      0,
      (anchor.to.row - anchor.from.row) * DEFAULT_ROW_HEIGHT_EMU
        + anchor.to.rowOffsetEmu
        - anchor.from.rowOffsetEmu,
    ),
    x: anchor.from.colOffsetEmu,
    y: anchor.from.rowOffsetEmu,
  };
}

export function applyGroupTransform(
  rect: DrawingRectEmu,
  group: GroupTransform,
): DrawingRectEmu {
  return {
    cx: rect.cx * group.scaleX,
    cy: rect.cy * group.scaleY,
    x: group.x + (rect.x - group.chX) * group.scaleX,
    y: group.y + (rect.y - group.chY) * group.scaleY,
  };
}

export function emuToPixels(value: number): number {
  return value / EMU_PER_PIXEL;
}

export function pixelsToEmu(value: number): number {
  return value * EMU_PER_PIXEL;
}

export function rectToAbsoluteAnchor(rect: DrawingRectEmu): XlsxImageAnchor {
  return {
    kind: "absolute",
    positionEmu: { x: rect.x, y: rect.y },
    sizeEmu: { cx: rect.cx, cy: rect.cy },
  };
}
