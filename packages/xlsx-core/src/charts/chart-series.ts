import type { Workbook } from "@dukelib/sheets-wasm";
import {
  normalizeHexColor,
  EMU_PER_PIXEL,
} from "./chart-colors";
import type {
  XlsxChart,
  XlsxChartAxis,
  XlsxChartSeries,
  XlsxChartReference,
  XlsxChartTypeGroup,
} from "../types";

export const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const DRAWING_SPREADSHEET_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
export const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
export const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
export const CHART_EX_REL_TYPE = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
export const CHART_STYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
export const CHART_COLOR_STYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
export const SERIES_COLORS = [
  "#4472c4",
  "#ed7d31",
  "#a5a5a5",
  "#ffc000",
  "#5b9bd5",
  "#70ad47",
  "#264478",
  "#9e480e",
  "#636363",
  "#997300"
];

export type ParsedChartSeriesFormula = {
  bubbleSizeFormula?: string;
  categoryFormula: string;
  nameFormula?: string;
  nameLiteral?: string;
  order: number;
  valueFormula: string;
};

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

function readSeriesNameFormula(series: XlsxChartSeries) {
  const raw = series.raw && typeof series.raw === "object" ? series.raw as Record<string, unknown> : null;
  return typeof raw?.name === "string" && raw.name.length > 0 ? raw.name : null;
}

export function buildChartSeriesFormula(chart: XlsxChart | null | undefined, seriesIndex: number) {
  const series = chart?.series[seriesIndex];
  if (!chart || !series) {
    return "";
  }

  const nameFormula = readSeriesNameFormula(series);
  const nameArgument = nameFormula ?? quoteSeriesFormulaString(series.name ?? `Series ${seriesIndex + 1}`);
  const categoryArgument = series.categoriesRef?.formula ?? "";
  const valueArgument = series.valuesRef?.formula ?? "";
  const orderArgument = String(seriesIndex + 1);
  const bubbleArgument = series.bubbleSizeRef?.formula;

  return [
    `=SERIES(${nameArgument}`,
    categoryArgument,
    valueArgument,
    orderArgument,
    ...(chart.chartType === "Bubble" || bubbleArgument ? [bubbleArgument ?? ""] : [])
  ].join(",") + ")";
}

export function parseChartSeriesFormula(formula: string, chart: XlsxChart | null | undefined): ParsedChartSeriesFormula | null {
  const trimmed = formula.trim();
  const withoutEquals = trimmed.startsWith("=") ? trimmed.slice(1).trim() : trimmed;
  const match = /^SERIES\s*\(([\s\S]*)\)$/i.exec(withoutEquals);
  if (!match) {
    return null;
  }

  const args = splitTopLevelSeriesArguments(match[1]);
  const isBubble = chart?.chartType === "Bubble";
  if (args.length < 4 || args.length > 5 || (isBubble && args.length !== 5)) {
    return null;
  }

  const [nameArg = "", categoryFormula = "", valueFormula = "", orderArg = "", bubbleSizeFormula] = args;
  if (!categoryFormula || !valueFormula) {
    return null;
  }

  const parsedOrder = Number(orderArg);
  if (!Number.isFinite(parsedOrder)) {
    return null;
  }

  const nameLiteral = unquoteSeriesFormulaString(nameArg);
  return {
    bubbleSizeFormula: bubbleSizeFormula && bubbleSizeFormula.length > 0 ? bubbleSizeFormula : undefined,
    categoryFormula,
    nameFormula: nameLiteral == null && nameArg.length > 0 ? nameArg : undefined,
    nameLiteral: nameLiteral ?? undefined,
    order: parsedOrder,
    valueFormula
  };
}

export function applyChartSeriesFormula(
  chart: XlsxChart,
  seriesIndex: number,
  formula: string,
  workbook: Workbook | null
) {
  const parsed = parseChartSeriesFormula(formula, chart);
  const currentSeries = chart.series[seriesIndex];
  if (!parsed || !currentSeries) {
    return null;
  }

  const categoriesRef: XlsxChartReference = {
    ...(currentSeries.categoriesRef ?? {}),
    formula: parsed.categoryFormula
  };
  const valuesRef: XlsxChartReference = {
    ...(currentSeries.valuesRef ?? {}),
    formula: parsed.valueFormula
  };
  const bubbleSizeRef: XlsxChartReference | null = chart.chartType === "Bubble" || parsed.bubbleSizeFormula
    ? {
        ...(currentSeries.bubbleSizeRef ?? {}),
        formula: parsed.bubbleSizeFormula
      }
    : currentSeries.bubbleSizeRef ?? null;
  const raw = {
    ...(currentSeries.raw ?? {})
  };
  if (parsed.nameFormula) {
    raw.name = parsed.nameFormula;
  } else {
    delete raw.name;
  }

  const nextSeries: XlsxChartSeries = {
    ...currentSeries,
    bubbleSizeRef,
    bubbleSizes: workbook && bubbleSizeRef?.formula
      ? resolveReferenceValues(workbook, chart.workbookSheetIndex, bubbleSizeRef, "value").map((value) => (
          typeof value === "number" && Number.isFinite(value) ? value : null
        ))
      : currentSeries.bubbleSizes,
    categories: workbook
      ? resolveReferenceValues(workbook, chart.workbookSheetIndex, categoriesRef, "category")
      : currentSeries.categories,
    categoriesRef,
    name: parsed.nameLiteral ?? (
      parsed.nameFormula && workbook
        ? resolveSeriesName(workbook, chart.workbookSheetIndex, parsed.nameFormula)
        : parsed.nameFormula ?? currentSeries.name
    ),
    raw,
    values: workbook
      ? resolveReferenceValues(workbook, chart.workbookSheetIndex, valuesRef, "value").map((value) => (
          typeof value === "number" && Number.isFinite(value) ? value : null
        ))
      : currentSeries.values,
    valuesRef
  };

  return {
    ...chart,
    series: chart.series.map((series, index) => index === seriesIndex ? nextSeries : series)
  };
}

function unquoteSheetName(value: string) {
  const trimmed = value.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

function splitSheetReference(reference: string) {
  let bangIndex = -1;
  let quoted = false;
  for (let index = 0; index < reference.length; index += 1) {
    const char = reference[index];
    if (char === "'") {
      quoted = !quoted;
    } else if (char === "!" && !quoted) {
      bangIndex = index;
      break;
    }
  }

  if (bangIndex < 0) {
    return null;
  }

  return {
    range: reference.slice(bangIndex + 1),
    sheetName: unquoteSheetName(reference.slice(0, bangIndex))
  };
}

function parseA1Cell(reference: string) {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(reference.trim());
  if (!match) {
    return null;
  }

  let col = 0;
  for (const char of match[1].toUpperCase()) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }

  return {
    col: col - 1,
    row: Number(match[2]) - 1
  };
}

function parseA1Range(reference: string) {
  const [startRef, endRef = startRef] = reference.split(":");
  const start = parseA1Cell(startRef ?? "");
  const end = parseA1Cell(endRef ?? "");
  if (!start || !end) {
    return null;
  }

  return {
    end: {
      col: Math.max(start.col, end.col),
      row: Math.max(start.row, end.row)
    },
    start: {
      col: Math.min(start.col, end.col),
      row: Math.min(start.row, end.row)
    }
  };
}

function formatA1Column(col: number) {
  let current = col + 1;
  let label = "";
  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }
  return label;
}

export function buildA1RangeFormula(sheetName: string, start: { col: number; row: number }, end: { col: number; row: number }) {
  const escapedSheetName = sheetName.replace(/'/g, "''");
  return `'${escapedSheetName}'!$${formatA1Column(start.col)}$${start.row + 1}:$${formatA1Column(end.col)}$${end.row + 1}`;
}

export function resolveReferenceSheet(workbook: Workbook, fallbackSheetIndex: number, formula?: string | null) {
  if (!formula) {
    return {
      range: null,
      sheet: workbook.getSheet(fallbackSheetIndex),
      sheetName: workbook.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }

  const trimmedFormula = formula.trim();
  if (trimmedFormula.length > 0 && !trimmedFormula.includes("!")) {
    try {
      const namedRange = workbook.getNamedRange(trimmedFormula);
      if (typeof namedRange === "string" && namedRange.length > 0 && namedRange !== trimmedFormula) {
        return resolveReferenceSheet(workbook, fallbackSheetIndex, namedRange);
      }
    } catch {
      // Fall back to direct A1 parsing when the workbook has no matching name.
    }
  }

  const split = splitSheetReference(trimmedFormula);
  if (!split) {
    return {
      range: parseA1Range(trimmedFormula),
      sheet: workbook.getSheet(fallbackSheetIndex),
      sheetName: workbook.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }

  try {
    return {
      range: parseA1Range(split.range),
      sheet: workbook.getSheetByName(split.sheetName),
      sheetName: split.sheetName
    };
  } catch {
    return {
      range: parseA1Range(split.range),
      sheet: workbook.getSheet(fallbackSheetIndex),
      sheetName: workbook.getSheet(fallbackSheetIndex)?.name ?? ""
    };
  }
}

export function resolveChartReferenceLabel(
  workbook: Workbook,
  fallbackSheetIndex: number,
  reference: XlsxChartReference | null | undefined,
  fallbackLabel: string
) {
  if (!reference?.formula) {
    return fallbackLabel;
  }

  const resolved = resolveReferenceSheet(workbook, fallbackSheetIndex, reference.formula);
  if (!resolved.sheet || !resolved.range) {
    return fallbackLabel;
  }

  const { start } = resolved.range;
  if (start.row > 0) {
    const headerDisplay = cellValueToDisplay(
      typeof resolved.sheet.getFormattedValueAt === "function"
        ? resolved.sheet.getFormattedValueAt(start.row - 1, start.col)
        : null
    );
    if (headerDisplay.length > 0) {
      return headerDisplay;
    }
  }

  const firstDisplay = cellValueToDisplay(
    typeof resolved.sheet.getFormattedValueAt === "function"
      ? resolved.sheet.getFormattedValueAt(start.row, start.col)
      : null
  );
  return firstDisplay.length > 0 ? firstDisplay : fallbackLabel;
}

export function resolveReferenceRowPaths(
  workbook: Workbook,
  fallbackSheetIndex: number,
  reference: XlsxChartReference | null | undefined
) {
  if (!reference?.formula) {
    return [] as string[][];
  }

  const resolved = resolveReferenceSheet(workbook, fallbackSheetIndex, reference.formula);
  if (!resolved.sheet || !resolved.range) {
    return [];
  }

  const rows: string[][] = [];
  for (let row = resolved.range.start.row; row <= resolved.range.end.row; row += 1) {
    const parts: string[] = [];
    for (let col = resolved.range.start.col; col <= resolved.range.end.col; col += 1) {
      const calculated = typeof resolved.sheet.getCalculatedValueAt === "function"
        ? resolved.sheet.getCalculatedValueAt(row, col)
        : null;
      const formatted = typeof resolved.sheet.getFormattedValueAt === "function"
        ? resolved.sheet.getFormattedValueAt(row, col)
        : calculated;
      const display = cellValueToDisplay(formatted ?? calculated);
      const numeric = cellValueToNumber(calculated ?? formatted);
      const label = display.length > 0 ? display : (numeric != null ? String(numeric) : "");
      if (label.length > 0) {
        parts.push(label);
      }
    }
    rows.push(parts);
  }
  return rows;
}

export function cellValueToNumber(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (value && typeof value === "object") {
    if ((value as { is_empty?: boolean }).is_empty) {
      return null;
    }
    const candidates: unknown[] = [];
    if (typeof (value as { asNumber?: () => unknown }).asNumber === "function") {
      candidates.push((value as { asNumber: () => unknown }).asNumber());
    }
    if (typeof (value as { toJs?: () => unknown }).toJs === "function") {
      candidates.push((value as { toJs: () => unknown }).toJs());
    }
    if (typeof (value as { asText?: () => unknown }).asText === "function") {
      candidates.push((value as { asText: () => unknown }).asText());
    }
    if (typeof (value as { toString?: () => unknown }).toString === "function") {
      candidates.push((value as { toString: () => unknown }).toString());
    }

    for (const candidate of candidates) {
      if (typeof candidate === "number" && Number.isFinite(candidate)) {
        return candidate;
      }
      if (typeof candidate === "string") {
        const parsed = Number(candidate.replace(/,/g, ""));
        if (Number.isFinite(parsed)) {
          return parsed;
        }
      }
    }
  }
  if (typeof value === "string") {
    const parsed = Number(value.replace(/,/g, ""));
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

export function cellValueToDisplay(value: unknown) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value && typeof value === "object") {
    if ((value as { is_empty?: boolean }).is_empty) {
      return "";
    }
    const candidates: unknown[] = [];
    if (typeof (value as { asText?: () => unknown }).asText === "function") {
      candidates.push((value as { asText: () => unknown }).asText());
    }
    if (typeof (value as { toJs?: () => unknown }).toJs === "function") {
      candidates.push((value as { toJs: () => unknown }).toJs());
    }
    if (typeof (value as { toString?: () => unknown }).toString === "function") {
      candidates.push((value as { toString: () => unknown }).toString());
    }

    for (const candidate of candidates) {
      if (candidate === null || candidate === undefined) {
        continue;
      }
      if (typeof candidate === "string") {
        return candidate;
      }
      return String(candidate);
    }
  }
  return String(value);
}

export function resolveReferenceValues(
  workbook: Workbook,
  fallbackSheetIndex: number,
  reference: XlsxChartReference | null | undefined,
  mode: "category" | "value"
): Array<number | string | null> {
  if (!reference?.formula) {
    return reference?.values ?? [];
  }

  const resolved = resolveReferenceSheet(workbook, fallbackSheetIndex, reference.formula);
  if (!resolved.sheet || !resolved.range) {
    return reference.values ?? [];
  }

  const values: Array<number | string | null> = [];
  for (let row = resolved.range.start.row; row <= resolved.range.end.row; row += 1) {
    for (let col = resolved.range.start.col; col <= resolved.range.end.col; col += 1) {
      const calculated = typeof resolved.sheet.getCalculatedValueAt === "function"
        ? resolved.sheet.getCalculatedValueAt(row, col)
        : null;
      const formatted = typeof resolved.sheet.getFormattedValueAt === "function"
        ? resolved.sheet.getFormattedValueAt(row, col)
        : calculated;
      if (mode === "value") {
        values.push(cellValueToNumber(calculated ?? formatted));
      } else {
        const display = cellValueToDisplay(formatted ?? calculated);
        const numeric = cellValueToNumber(calculated ?? formatted);
        values.push(display.length > 0 ? display : (numeric !== null ? numeric : null));
      }
    }
  }

  return values;
}

export function resolveSeriesName(workbook: Workbook, fallbackSheetIndex: number, rawName: unknown) {
  if (typeof rawName !== "string" || !rawName) {
    return undefined;
  }

  const resolved = resolveReferenceSheet(workbook, fallbackSheetIndex, rawName);
  if (!resolved.sheet || !resolved.range) {
    return rawName;
  }

  const value = typeof resolved.sheet.getFormattedValueAt === "function"
    ? resolved.sheet.getFormattedValueAt(resolved.range.start.row, resolved.range.start.col)
    : null;
  const display = cellValueToDisplay(value);
  return display || rawName;
}

export function normalizeChartSeries(
  workbook: Workbook,
  workbookSheetIndex: number,
  chartId: string,
  raw: unknown,
  index: number
): XlsxChartSeries {
  const series = raw && typeof raw === "object" ? raw as Record<string, unknown> : {};
  const categoriesRef = normalizeChartReference(series.categories);
  const valuesRef = normalizeChartReference(series.values);
  const shapeProperties = series.shapeProperties && typeof series.shapeProperties === "object"
    ? series.shapeProperties as Record<string, unknown>
    : undefined;
  const rawFillColor = typeof shapeProperties?.solidFillHex === "string"
    ? normalizeHexColor(shapeProperties.solidFillHex)
    : null;
  const rawLineColor = typeof shapeProperties?.lineColorHex === "string"
    ? normalizeHexColor(shapeProperties.lineColorHex)
    : null;
  const bubbleSizeRef = normalizeChartReference(series.bubbleSize ?? series.bubbleSizes ?? series.bubbles);

  return {
    bubbleSizeRef,
    bubbleSizes: resolveReferenceValues(workbook, workbookSheetIndex, bubbleSizeRef, "value").map((value) => (
      typeof value === "number" && Number.isFinite(value) ? value : null
    )),
    categories: resolveReferenceValues(workbook, workbookSheetIndex, categoriesRef, "category"),
    categoriesRef,
    color: rawFillColor ?? undefined,
    dataPoints: Array.isArray(series.dataPoints) ? series.dataPoints : [],
    dataPointStyles: undefined,
    id: `${chartId}-series-${index}`,
    invertIfNegative: typeof series.invertIfNegative === "boolean" ? series.invertIfNegative : undefined,
    lineColor: rawLineColor ?? rawFillColor ?? undefined,
    lineWidthPx: typeof shapeProperties?.lineWidth === "number"
      ? Math.max(1, Number(shapeProperties.lineWidth) / EMU_PER_PIXEL)
      : undefined,
    marker: series.marker && typeof series.marker === "object" ? series.marker as Record<string, unknown> : undefined,
    markerColor: undefined,
    markerLineColor: undefined,
    markerSize: series.marker && typeof series.marker === "object" && typeof (series.marker as Record<string, unknown>).size === "number"
      ? Number((series.marker as Record<string, unknown>).size)
      : undefined,
    markerSymbol: series.marker && typeof series.marker === "object" && typeof (series.marker as Record<string, unknown>).symbol === "string"
      ? String((series.marker as Record<string, unknown>).symbol)
      : undefined,
    name: resolveSeriesName(workbook, workbookSheetIndex, series.name),
    negativeColor: undefined,
    negativeLineColor: undefined,
    raw: series,
    shapeProperties,
    smooth: typeof series.smooth === "boolean" ? series.smooth : undefined,
    values: resolveReferenceValues(workbook, workbookSheetIndex, valuesRef, "value").map((value) => (
      typeof value === "number" && Number.isFinite(value) ? value : null
    )),
    valuesRef
  };
}


export function normalizeChartReference(raw: unknown): XlsxChartReference | null {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const record = raw as Record<string, unknown>;
  return {
    formula: typeof record.formula === "string" ? record.formula : undefined,
    refType: typeof record.refType === "string" ? record.refType : undefined,
    values: Array.isArray(record.values) ? record.values as Array<number | string | null> : undefined
  };
}

