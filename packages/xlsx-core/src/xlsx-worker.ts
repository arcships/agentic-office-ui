import type { Workbook } from "@dukelib/sheets-wasm";
import {
  DOMParser as WorkerDomParser,
  Node as WorkerNode,
  XMLSerializer as WorkerXmlSerializer,
} from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { OfficeLoadError } from "@arcships/office-runtime";
import { loadWorkbookChartAssets } from "./charts";
import {
  parseWorkbookChartStyleAssets,
  parseWorkbookImageAssets,
  parseWorkbookStructureAssets,
  normalizeWorkbookTableMetadata,
  resolveSheetColumnWidthPixels,
  resolveWorksheetDefaultColumnWidthPixels,
  resolveWorksheetDefaultRowHeightPixels,
  resolveWorksheetMergeMetadata,
  revokeWorkbookImageAssets
} from "./images";
import type { WorkbookStructureAssets } from "./images";
import { safeCalculate } from "./safe-calculate";
import { trackXlsxWorkbookLifetime } from "./workbook-lifetime";
import { getSheetsWasmModule, setWasmSource, type WorkerWasmSource } from "./wasm";
import type {
  XlsxChart,
  XlsxChartsheet,
  XlsxCellAddress,
  XlsxCellRange,
  XlsxDataValidation,
  XlsxFreezePanes,
  XlsxFormControl,
  XlsxImage,
  XlsxResolvedCellStyle,
  XlsxShape,
  XlsxSheetData,
  XlsxSheetVisibility,
  XlsxTable,
  XlsxTableStyleDefinition,
  XlsxWorkbookTab
} from "./types";
import {
  validateXlsxArchive,
  type XlsxRuntimeLimits,
} from "./resource-limits";
import type { XlsxWorkerErrorCode } from "./worker-client";

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_WIDTH = 80;
const DEFAULT_ZOOM_SCALE = 100;
const FORMULA_COUNT_THRESHOLD = 1000;
const MIN_ROW_HEIGHT_PX = 16;

const workerXmlGlobals = globalThis as typeof globalThis & {
  DOMParser?: typeof DOMParser;
  Node?: typeof Node;
  XMLSerializer?: typeof XMLSerializer;
};
workerXmlGlobals.DOMParser ??= WorkerDomParser as unknown as typeof DOMParser;
workerXmlGlobals.Node ??= WorkerNode as unknown as typeof Node;
workerXmlGlobals.XMLSerializer ??= WorkerXmlSerializer as unknown as typeof XMLSerializer;

type WorkerSheetState = Partial<NonNullable<WorkbookStructureAssets["sheetStatesByWorkbookSheetIndex"][number]>>;

function isLegacyXlsWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 8
    && bytes[0] === 0xd0
    && bytes[1] === 0xcf
    && bytes[2] === 0x11
    && bytes[3] === 0xe0
    && bytes[4] === 0xa1
    && bytes[5] === 0xb1
    && bytes[6] === 0x1a
    && bytes[7] === 0xe1;
}

function shouldSkipXmlParsingForWorkbook(bytes: Uint8Array, skipXmlParsing = false) {
  return skipXmlParsing || isLegacyXlsWorkbook(bytes);
}

function normalizeWorksheetVisibility(value: unknown): XlsxSheetVisibility {
  return value === "hidden" || value === "veryHidden" ? value : "visible";
}

type WorkerRequest =
  | {
      id: number;
      type: "load";
      payload: {
        buffer: ArrayBuffer;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
        wasmSource?: WorkerWasmSource;
        limits: Readonly<XlsxRuntimeLimits>;
      };
    }
  | {
      id: number;
      type: "parseCharts";
      payload: {
        buffer: ArrayBuffer;
        showHiddenSheets?: boolean;
        skipXmlParsing?: boolean;
        wasmSource?: WorkerWasmSource;
        limits: Readonly<XlsxRuntimeLimits>;
      };
    }
  | {
      id: number;
      type: "getCellSnapshot";
      payload: {
        workbookSheetIndex: number;
        row: number;
        col: number;
      };
    }
  | {
      id: number;
      type: "getRowsBatch";
      payload: {
        workbookSheetIndex: number;
        startRow: number;
        rowCount: number;
      };
    };

type WorkerSuccessResponse = {
  id: number;
  success: true;
  result:
    | {
        chartsByWorkbookSheetIndex: XlsxChart[][];
        chartsheets: XlsxChartsheet[];
        tabs: XlsxWorkbookTab[];
      }
    | {
        chartsByWorkbookSheetIndex: XlsxChart[][];
        chartsheets: XlsxChartsheet[];
        formControlsByWorkbookSheetIndex: XlsxFormControl[][];
        imagesByWorkbookSheetIndex: XlsxImage[][];
        objectUrls: string[];
        shapesByWorkbookSheetIndex: XlsxShape[][];
        sheets: XlsxSheetData[];
        tablesByWorkbookSheetIndex: XlsxTable[][];
        tabs: XlsxWorkbookTab[];
      }
    | {
        displayValue: string;
        formula: string;
      }
    | unknown[]
    | null;
};

type WorkerErrorResponse = {
  id: number;
  success: false;
  error: {
    code: XlsxWorkerErrorCode;
    message: string;
    phase?: string;
    limit?: string;
    actual?: number;
    allowed?: number;
  };
};

type WorkerResponse = WorkerSuccessResponse | WorkerErrorResponse;

let workbook: Workbook | null = null;
let chartsByWorkbookSheetIndex: XlsxChart[][] = [];
let chartsheets: XlsxChartsheet[] = [];
let sheets: XlsxSheetData[] = [];
let tablesByWorkbookSheetIndex: XlsxTable[][] = [];
let tabs: XlsxWorkbookTab[] = [];

function canParseXmlInWorker() {
  return typeof DOMParser !== "undefined";
}

function decodeXmlAttribute(value: string) {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function readXmlAttribute(tag: string, name: string) {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = new RegExp(`(?:^|\\s)${escapedName}="([^"]*)"`).exec(tag);
  return match ? decodeXmlAttribute(match[1] ?? "") : null;
}

function readArchiveText(archive: Record<string, Uint8Array>, path: string) {
  const entry = archive[path];
  return entry ? strFromU8(entry) : "";
}

function normalizeWorkbookRelationshipTarget(target: string) {
  if (target.startsWith("/")) {
    return target.replace(/^\/+/, "");
  }

  return target.startsWith("xl/")
    ? target
    : `xl/${target.replace(/^\.?\//, "")}`;
}

function parseWorkbookSheetPathsFromArchive(archive: Record<string, Uint8Array>) {
  const workbookXml = readArchiveText(archive, "xl/workbook.xml");
  const workbookRelationshipsXml = readArchiveText(archive, "xl/_rels/workbook.xml.rels");
  if (!workbookXml || !workbookRelationshipsXml) {
    return [] as string[];
  }

  const relationshipTargetById = new Map<string, string>();
  for (const match of workbookRelationshipsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const tag = match[0];
    const id = readXmlAttribute(tag, "Id");
    const target = readXmlAttribute(tag, "Target");
    if (id && target) {
      relationshipTargetById.set(id, normalizeWorkbookRelationshipTarget(target));
    }
  }

  const paths: string[] = [];
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const tag = match[0];
    const relationshipId = readXmlAttribute(tag, "r:id") ?? readXmlAttribute(tag, "id");
    const target = relationshipId ? relationshipTargetById.get(relationshipId) : null;
    if (target) {
      paths.push(target);
    }
  }

  return paths;
}

function parseWorkerSheetLayoutAssets(bytes: Uint8Array, sheetCount: number): Array<WorkerSheetState | null> {
  try {
    const archive = unzipSync(bytes);
    const workbookSheetPaths = parseWorkbookSheetPathsFromArchive(archive);
    const sheetPaths = workbookSheetPaths.length > 0
      ? workbookSheetPaths
      : Array.from({ length: sheetCount }, (_, index) => `xl/worksheets/sheet${index + 1}.xml`);

    return sheetPaths.slice(0, sheetCount).map((path) => {
      const xml = readArchiveText(archive, path);
      if (!xml) {
        return null;
      }

      const rowHeightOverridesPx: Record<number, number> = {};
      for (const match of xml.matchAll(/<row\b[^>]*>/g)) {
        const tag = match[0];
        const rowNumber = Number(readXmlAttribute(tag, "r") ?? Number.NaN);
        const height = Number(readXmlAttribute(tag, "ht") ?? Number.NaN);
        const rowIndex = rowNumber - 1;
        if (rowIndex >= 0 && Number.isFinite(height)) {
          rowHeightOverridesPx[rowIndex] = Math.max(MIN_ROW_HEIGHT_PX, Math.round(height * 1.33));
        }
      }

      return { rowHeightOverridesPx };
    });
  } catch {
    return [];
  }
}

function buildVisibleSheetIndexByWorkbookSheetIndex(nextWorkbook: Workbook, showHiddenSheets = false) {
  const mapping = new Map<number, number>();
  let visibleIndex = 0;
  for (let workbookSheetIndex = 0; workbookSheetIndex < nextWorkbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = nextWorkbook.getSheet(workbookSheetIndex);
    const visibility = normalizeWorksheetVisibility(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      continue;
    }
    mapping.set(workbookSheetIndex, visibleIndex);
    visibleIndex += 1;
  }
  return mapping;
}

function normalizeRange(range: XlsxCellRange): XlsxCellRange {
  return {
    start: {
      col: Math.min(range.start.col, range.end.col),
      row: Math.min(range.start.row, range.end.row)
    },
    end: {
      col: Math.max(range.start.col, range.end.col),
      row: Math.max(range.start.row, range.end.row)
    }
  };
}

function parseA1CellReference(reference: string): XlsxCellAddress | null {
  const match = /^([A-Z]+)(\d+)$/i.exec(reference.trim());
  if (!match) {
    return null;
  }

  const [, columnPart, rowPart] = match;
  let col = 0;
  for (const char of columnPart.toUpperCase()) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }

  return {
    col: col - 1,
    row: Number(rowPart) - 1
  };
}

function parseA1RangeReference(reference: string): XlsxCellRange | null {
  const [startRef, endRef = startRef] = reference.split(":");
  const start = parseA1CellReference(startRef ?? "");
  const end = parseA1CellReference(endRef ?? "");
  if (!start || !end) {
    return null;
  }

  return normalizeRange({ end, start });
}

function parseWorksheetFreezePanes(worksheet: ReturnType<Workbook["getSheet"]>): XlsxFreezePanes | null {
  const rawFreezePanes = worksheet.freezePanes as Record<string, unknown> | null | undefined;
  const row = typeof rawFreezePanes?.row === "number" && rawFreezePanes.row >= 0 ? rawFreezePanes.row : null;
  const col = typeof rawFreezePanes?.col === "number" && rawFreezePanes.col >= 0 ? rawFreezePanes.col : null;
  if (row === null && col === null) {
    return null;
  }

  return {
    col: col ?? 0,
    row: row ?? 0
  };
}

function parseWorksheetDataValidations(worksheet: ReturnType<Workbook["getSheet"]>): XlsxDataValidation[] {
  const rawDataValidations = Array.isArray(worksheet.dataValidations) ? worksheet.dataValidations : [];

  return rawDataValidations.flatMap((entry) => {
    if (!entry || typeof entry !== "object") {
      return [];
    }

    const validation = entry as Record<string, unknown>;
    const ranges = Array.isArray(validation.ranges)
      ? validation.ranges.flatMap((range) => {
          if (typeof range !== "string") {
            return [];
          }

          const parsedRange = parseA1RangeReference(range);
          return parsedRange ? [parsedRange] : [];
        })
      : [];
    const validationType = typeof validation.validationType === "string" ? validation.validationType : null;
    if (!validationType || ranges.length === 0) {
      return [];
    }

    return [{
      allowBlank: typeof validation.allowBlank === "boolean" ? validation.allowBlank : undefined,
      errorMessage: typeof validation.errorMessage === "string" ? validation.errorMessage : undefined,
      errorStyle: typeof validation.errorStyle === "string" ? validation.errorStyle : undefined,
      inputMessage: typeof validation.inputMessage === "string" ? validation.inputMessage : undefined,
      listSource: typeof validation.listSource === "string" ? validation.listSource : undefined,
      ranges,
      showDropdown: typeof validation.showDropdown === "boolean" ? validation.showDropdown : undefined,
      showErrorAlert: typeof validation.showErrorAlert === "boolean" ? validation.showErrorAlert : undefined,
      showInputMessage: typeof validation.showInputMessage === "boolean" ? validation.showInputMessage : undefined,
      validationType
    } satisfies XlsxDataValidation];
  });
}

function resolveWorksheetZoomScale(
  worksheet: ReturnType<Workbook["getSheet"]>,
  sheetState?: { zoomScale?: number } | null
) {
  const candidates = [
    sheetState?.zoomScale,
    typeof worksheet.zoomScale === "number" ? worksheet.zoomScale : undefined
  ];
  const value = candidates.find((entry): entry is number => typeof entry === "number" && Number.isFinite(entry) && entry > 0);
  return value ?? DEFAULT_ZOOM_SCALE;
}

function resolveSheetDisplayUsedRange(
  usedRange: [number, number, number, number],
  sheetState?: {
    maxContentCol?: number;
    maxContentRow?: number;
    maxHorizontalMergeEndCol?: number;
    maxVerticalMergeEndRow?: number;
    minContentCol?: number;
    minContentRow?: number;
  } | null
): [number, number, number, number] {
  const [minRow, minCol, maxRow, maxCol] = usedRange;
  const maxContentRow = sheetState?.maxContentRow ?? -1;
  const maxContentCol = sheetState?.maxContentCol ?? -1;
  const maxVerticalMergeEndRow = sheetState?.maxVerticalMergeEndRow ?? -1;
  const maxHorizontalMergeEndCol = sheetState?.maxHorizontalMergeEndCol ?? -1;
  const maxMeaningfulRow = Math.max(maxContentRow, maxVerticalMergeEndRow);
  const maxMeaningfulCol = Math.max(maxContentCol, maxHorizontalMergeEndCol);

  if (maxMeaningfulRow < 0 && maxMeaningfulCol < 0) {
    return usedRange;
  }

  return [
    sheetState?.minContentRow !== undefined && sheetState.minContentRow >= 0 ? Math.min(minRow, sheetState.minContentRow) : minRow,
    sheetState?.minContentCol !== undefined && sheetState.minContentCol >= 0 ? Math.min(minCol, sheetState.minContentCol) : minCol,
    maxMeaningfulRow >= 0
      ? (maxContentRow >= 0 ? Math.min(maxRow, maxMeaningfulRow) : Math.max(maxRow, maxMeaningfulRow))
      : maxRow,
    maxMeaningfulCol >= 0
      ? (maxContentCol >= 0 ? Math.min(maxCol, maxMeaningfulCol) : Math.max(maxCol, maxMeaningfulCol))
      : maxCol
  ];
}

function buildSheetList(
  nextWorkbook: Workbook,
  structureAssets?: WorkbookStructureAssets | null,
  sheetLayoutStates?: Array<WorkerSheetState | null>,
  showHiddenSheets = false
) {
  const sheetsByWorkbookSheetIndex: XlsxSheetData[] = [];

  for (let index = 0; index < nextWorkbook.sheetCount; index += 1) {
    const worksheet = nextWorkbook.getSheet(index);
    const sheetState = structureAssets?.sheetStatesByWorkbookSheetIndex[index] ?? sheetLayoutStates?.[index] ?? null;
    const mergeMetadata = resolveWorksheetMergeMetadata(worksheet);
    const effectiveSheetState = {
      ...sheetState,
      ...mergeMetadata
    };
    const defaultColWidthPx = resolveWorksheetDefaultColumnWidthPixels(
      worksheet,
      sheetState?.columnWidthCharacterWidthPx,
      sheetState?.defaultColWidthPx ?? DEFAULT_COL_WIDTH
    );
    const defaultRowHeightPx = resolveWorksheetDefaultRowHeightPixels(
      worksheet,
      sheetState?.defaultRowHeightPx ?? DEFAULT_ROW_HEIGHT
    );
    const visibility = normalizeWorksheetVisibility(worksheet.visibility);
    if (!showHiddenSheets && visibility !== "visible") {
      continue;
    }

    const resolveColumnWidthPx = (col: number) => {
      const width = worksheet.getColumnWidth(col);
      if (width !== undefined && width !== null) {
        return resolveSheetColumnWidthPixels(width, sheetState?.columnWidthCharacterWidthPx);
      }

      return sheetState?.colWidthOverridesPx?.[col] ?? defaultColWidthPx;
    };

    const resolveRowHeightPx = (row: number) => {
      const height = worksheet.getRowHeight(row);
      if (height !== undefined && height !== null) {
        return Math.max(Math.round(height * 1.33), 16);
      }

      return sheetState?.rowHeightOverridesPx?.[row] ?? defaultRowHeightPx;
    };

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    if (!usedRange) {
      sheetsByWorkbookSheetIndex.push({
        cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
        columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
        colCount: 0,
        colStyleIds: sheetState?.colStyleIds ?? {},
        colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
        colWidths: [],
        conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
        hyperlinks: sheetState?.hyperlinks ?? [],
        comments: sheetState?.comments ?? [],
        dataValidations: parseWorksheetDataValidations(worksheet),
        defaultColWidthPx,
        defaultRowHeightPx,
        freezePanes: parseWorksheetFreezePanes(worksheet),
        hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
        hasVerticalMerges: mergeMetadata.hasVerticalMerges,
        maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
        maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
        mergedRegions: sheetState?.mergedRegions ?? [],
        hiddenCols: sheetState?.hiddenCols ?? [],
        hiddenRows: sheetState?.hiddenRows ?? [],
        minUsedCol: -1,
        minUsedRow: -1,
        maxUsedCol: -1,
        maxUsedRow: -1,
        name: worksheet.name,
        visibility,
        namedCellStyleByName: structureAssets?.namedCellStyleByName ?? {},
        rowCount: 0,
        rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
        rowHeights: [],
        rowStyleIds: sheetState?.rowStyleIds ?? {},
        showGridLines: sheetState?.showGridLines ?? true,
        sparklines: sheetState?.sparklines ?? [],
        styleById: structureAssets?.styleById ?? {},
        tableStyleByName: structureAssets?.tableStyleByName ?? {},
        themePalette: structureAssets?.themePalette ?? { colorsByIndex: {} },
        visibleCols: [],
        visibleRows: [],
        workbookSheetIndex: index,
        zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
      });
      continue;
    }

    const [minRow, minCol, maxRow, maxCol] = resolveSheetDisplayUsedRange(usedRange, effectiveSheetState);
    const visibleRows: number[] = [];
    const hiddenRows: number[] = [];
    for (let row = 0; row <= maxRow; row += 1) {
      if (worksheet.isRowHidden(row)) {
        hiddenRows.push(row);
      } else {
        visibleRows.push(row);
      }
    }

    const visibleCols: number[] = [];
    const hiddenCols: number[] = [];
    for (let col = 0; col <= maxCol; col += 1) {
      if (worksheet.isColumnHidden(col)) {
        hiddenCols.push(col);
      } else {
        visibleCols.push(col);
      }
    }

    sheetsByWorkbookSheetIndex.push({
      cachedFormulaValues: sheetState?.cachedFormulaValues ?? {},
      columnWidthCharacterWidthPx: sheetState?.columnWidthCharacterWidthPx,
      colCount: visibleCols.length,
      colStyleIds: sheetState?.colStyleIds ?? {},
      colWidthOverridesPx: sheetState?.colWidthOverridesPx ?? {},
      colWidths: visibleCols.map(resolveColumnWidthPx),
      conditionalFormatRules: sheetState?.conditionalFormatRules ?? [],
      hyperlinks: sheetState?.hyperlinks ?? [],
      comments: sheetState?.comments ?? [],
      dataValidations: parseWorksheetDataValidations(worksheet),
      defaultColWidthPx,
      defaultRowHeightPx,
      freezePanes: parseWorksheetFreezePanes(worksheet),
      hasHorizontalMerges: mergeMetadata.hasHorizontalMerges,
      hasVerticalMerges: mergeMetadata.hasVerticalMerges,
      maxHorizontalMergeEndCol: mergeMetadata.maxHorizontalMergeEndCol,
      maxVerticalMergeEndRow: mergeMetadata.maxVerticalMergeEndRow,
      mergedRegions: sheetState?.mergedRegions ?? [],
      hiddenCols,
      hiddenRows,
      minUsedCol: minCol,
      minUsedRow: minRow,
      maxUsedCol: maxCol,
      maxUsedRow: maxRow,
      name: worksheet.name,
      visibility,
      namedCellStyleByName: structureAssets?.namedCellStyleByName ?? {},
      rowCount: visibleRows.length,
      rowHeightOverridesPx: sheetState?.rowHeightOverridesPx ?? {},
      rowHeights: visibleRows.map(resolveRowHeightPx),
      rowStyleIds: sheetState?.rowStyleIds ?? {},
      showGridLines: sheetState?.showGridLines ?? true,
      sparklines: sheetState?.sparklines ?? [],
      styleById: structureAssets?.styleById ?? {},
      tableStyleByName: structureAssets?.tableStyleByName ?? {},
      themePalette: structureAssets?.themePalette ?? { colorsByIndex: {} },
      visibleCols,
      visibleRows,
      workbookSheetIndex: index,
      zoomScale: resolveWorksheetZoomScale(worksheet, sheetState)
    });
  }

  return sheetsByWorkbookSheetIndex;
}

function mapWorksheetTables(worksheet: ReturnType<Workbook["getSheet"]> | null): XlsxTable[] {
  const rawTables = (worksheet?.tables ?? []) as Array<Record<string, unknown>>;
  return rawTables.flatMap((table, index) => {
    const rawColumns = Array.isArray(table.columns) ? table.columns : [];
    const rawName = typeof table.name === "string" ? table.name : `Table${index + 1}`;
    const rawDisplayName =
      typeof table.displayName === "string"
        ? table.displayName
        : typeof table.name === "string"
          ? table.name
          : `Table ${index + 1}`;
    const rawReference = typeof table.reference === "string" ? table.reference : "";
    const reference = rawReference;
    const parsedRange = parseA1RangeReference(reference);
    if (!parsedRange) {
      return [];
    }

    return [{
      columns: rawColumns.map((column, columnIndex) => ({
        id: typeof (column as { id?: unknown }).id === "number" ? ((column as { id?: number }).id ?? columnIndex + 1) : columnIndex + 1,
        index: columnIndex,
        name: typeof (column as { name?: unknown }).name === "string" ? ((column as { name?: string }).name ?? `Column ${columnIndex + 1}`) : `Column ${columnIndex + 1}`
      })),
      displayName: rawDisplayName,
      end: parsedRange.end,
      headerRowCount: resolveWorkbookTableCount(table.headerRowCount, 1),
      headerRowCellStyle: typeof table.headerRowCellStyle === "string" ? table.headerRowCellStyle : undefined,
      name: rawName,
      reference,
      start: parsedRange.start,
      styleInfo: table.styleInfo as XlsxTable["styleInfo"] | undefined,
      totalsRowCount: resolveWorkbookTableCount(table.totalsRowCount, 0),
      totalsRowShown: resolveWorkbookTableBoolean(table.totalsRowShown)
    }];
  });
}

function resolveWorkbookTableCount(value: unknown, fallback: number) {
  if (typeof value === "number" && Number.isFinite(value) && value >= 0) {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed) && parsed >= 0) {
      return parsed;
    }
  }

  return fallback;
}

function resolveWorkbookTableBoolean(value: unknown) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    return value !== 0;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "0" || normalized === "false" || normalized === "") {
      return false;
    }
    if (normalized === "1" || normalized === "true") {
      return true;
    }
  }

  return false;
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&#34;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function getCellDisplayValue(worksheet: ReturnType<Workbook["getSheet"]>, row: number, col: number, activeSheet?: XlsxSheetData | null) {
  const formula = worksheet.getFormulaAt(row, col);
  const cachedFormulaValue = formula ? activeSheet?.cachedFormulaValues?.[cellAddressToA1({ row, col })] : undefined;
  const formatted = worksheet.getFormattedValueAt(row, col);
  if (formatted && !(formula && cachedFormulaValue !== undefined && formatted.startsWith("#"))) {
    return decodeHtmlEntities(formatted);
  }

  const cellValue = worksheet.getCalculatedValueAt(row, col);
  try {
    if (formula && cachedFormulaValue !== undefined && cellValue.is_error) {
      return cachedFormulaValue;
    }
    if (cellValue.is_error) {
      return cellValue.asError() ?? "";
    }
    if (cellValue.is_empty) {
      return "";
    }

    return cellValue.toString();
  } finally {
    cellValue.free();
  }
}

function cellAddressToA1(cell: XlsxCellAddress) {
  let col = cell.col + 1;
  let label = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    col = Math.floor((col - 1) / 26);
  }
  return `${label}${cell.row + 1}`;
}

async function loadWorkbook(
  buffer: ArrayBuffer,
  skipXmlParsing = false,
  showHiddenSheets = false,
  limits?: XlsxRuntimeLimits,
) {
  validateXlsxArchive(buffer, limits);
  const wasmModule = await getSheetsWasmModule();
  const bytes = new Uint8Array(buffer);
  const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing);
  let activeWorkbook = trackXlsxWorkbookLifetime(wasmModule.Workbook.fromBytes(bytes));
  let totalFormulas = 0;
  for (let index = 0; index < activeWorkbook.sheetCount; index += 1) {
    const worksheet = activeWorkbook.getSheet(index);
    try {
      totalFormulas += worksheet.formulaCount;
    } finally {
      worksheet.free();
    }
  }

  if (totalFormulas <= FORMULA_COUNT_THRESHOLD) {
    const result = safeCalculate(activeWorkbook, {
      reparse: () => trackXlsxWorkbookLifetime(wasmModule.Workbook.fromBytes(bytes))
    });
    if (result.workbook !== activeWorkbook) {
      try { activeWorkbook.free(); } catch { /* trapped workbook is never reused */ }
    }
    activeWorkbook = result.workbook;
  }

  const nextWorkbook = activeWorkbook;
  const imageAssets = effectiveSkipXmlParsing || !canParseXmlInWorker()
    ? null
    : parseWorkbookImageAssets(bytes, limits);
  const structureAssets = imageAssets ?? (
    effectiveSkipXmlParsing || !canParseXmlInWorker()
      ? null
      : parseWorkbookStructureAssets(bytes, { includeCachedFormulaValues: true })
  );
  const sheetLayoutStates = structureAssets ? undefined : parseWorkerSheetLayoutAssets(bytes, nextWorkbook.sheetCount);
  const previousWorkbook = workbook;
  workbook = nextWorkbook;
  try {
    sheets = buildSheetList(nextWorkbook, structureAssets, sheetLayoutStates, showHiddenSheets);
    tablesByWorkbookSheetIndex = Array.from(
      { length: nextWorkbook.sheetCount },
      (_, workbookSheetIndex) => {
        const parsedTables = structureAssets?.tableMetadataByWorkbookSheetIndex[workbookSheetIndex];
        return parsedTables?.length
          ? normalizeWorkbookTableMetadata(parsedTables)
          : mapWorksheetTables(nextWorkbook.getSheet(workbookSheetIndex));
      }
    );
    const visibleSheetIndexByWorkbookSheetIndex = new Map(sheets.map((sheet, index) => [sheet.workbookSheetIndex, index]));
    const chartStyleAssets = effectiveSkipXmlParsing || !canParseXmlInWorker()
      ? null
      : parseWorkbookChartStyleAssets(bytes);
    const chartAssets = loadWorkbookChartAssets(
      nextWorkbook,
      chartStyleAssets,
      visibleSheetIndexByWorkbookSheetIndex,
      showHiddenSheets
    );
    chartsByWorkbookSheetIndex = chartAssets.chartsByWorkbookSheetIndex;
    chartsheets = chartAssets.chartsheets;
    tabs = chartAssets.tabs;
    if (previousWorkbook && previousWorkbook !== nextWorkbook) {
      try { previousWorkbook.free(); } catch { /* previous workbook is never reused */ }
    }
    return {
      chartsByWorkbookSheetIndex,
      chartsheets,
      formControlsByWorkbookSheetIndex: imageAssets?.formControlsByWorkbookSheetIndex ?? [],
      imagesByWorkbookSheetIndex: imageAssets?.imagesByWorkbookSheetIndex ?? [],
      objectUrls: imageAssets?.objectUrls ?? [],
      shapesByWorkbookSheetIndex: imageAssets?.shapesByWorkbookSheetIndex ?? [],
      sheets,
      tablesByWorkbookSheetIndex,
      tabs
    };
  } catch (error) {
    workbook = previousWorkbook;
    revokeWorkbookImageAssets(imageAssets);
    try { nextWorkbook.free(); } catch { /* failed load is discarded */ }
    throw error;
  }
}

async function parseCharts(
  buffer: ArrayBuffer,
  skipXmlParsing = false,
  showHiddenSheets = false,
  limits?: XlsxRuntimeLimits,
) {
  validateXlsxArchive(buffer, limits);
  const wasmModule = await getSheetsWasmModule();
  const bytes = new Uint8Array(buffer);
  const effectiveSkipXmlParsing = shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing);
  let activeWorkbook = trackXlsxWorkbookLifetime(wasmModule.Workbook.fromBytes(bytes));
  let totalFormulas = 0;
  for (let index = 0; index < activeWorkbook.sheetCount; index += 1) {
    const worksheet = activeWorkbook.getSheet(index);
    try {
      totalFormulas += worksheet.formulaCount;
    } finally {
      worksheet.free();
    }
  }
  if (totalFormulas <= FORMULA_COUNT_THRESHOLD) {
    const result = safeCalculate(activeWorkbook, {
      reparse: () => trackXlsxWorkbookLifetime(wasmModule.Workbook.fromBytes(bytes))
    });
    if (result.workbook !== activeWorkbook) {
      try { activeWorkbook.free(); } catch { /* trapped workbook is never reused */ }
    }
    activeWorkbook = result.workbook;
  }

  const nextWorkbook = activeWorkbook;
  try {
    const visibleSheetIndexByWorkbookSheetIndex = buildVisibleSheetIndexByWorkbookSheetIndex(nextWorkbook, showHiddenSheets);
    const chartStyleAssets = effectiveSkipXmlParsing || !canParseXmlInWorker()
      ? null
      : parseWorkbookChartStyleAssets(bytes);
    const chartAssets = loadWorkbookChartAssets(
      nextWorkbook,
      chartStyleAssets,
      visibleSheetIndexByWorkbookSheetIndex,
      showHiddenSheets
    );
    return {
      chartsByWorkbookSheetIndex: chartAssets.chartsByWorkbookSheetIndex,
      chartsheets: chartAssets.chartsheets,
      tabs: chartAssets.tabs
    };
  } finally {
    try { nextWorkbook.free(); } catch { /* chart-only workbook is temporary */ }
  }
}

function respond(message: WorkerResponse) {
  self.postMessage(message);
}

async function handleMessage(message: WorkerRequest) {
  switch (message.type) {
    case "load": {
      if (message.payload.wasmSource !== undefined) {
        setWasmSource(message.payload.wasmSource);
      }
      return loadWorkbook(
        message.payload.buffer,
        message.payload.skipXmlParsing,
        message.payload.showHiddenSheets,
        message.payload.limits,
      );
    }
    case "parseCharts": {
      if (message.payload.wasmSource !== undefined) {
        setWasmSource(message.payload.wasmSource);
      }
      return parseCharts(
        message.payload.buffer,
        message.payload.skipXmlParsing,
        message.payload.showHiddenSheets,
        message.payload.limits,
      );
    }
    case "getCellSnapshot": {
      if (!workbook) {
        return {
          displayValue: "",
          formula: ""
        };
      }

      const targetSheet = sheets.find((sheet) => sheet.workbookSheetIndex === message.payload.workbookSheetIndex) ?? null;
      const worksheet = workbook.getSheet(message.payload.workbookSheetIndex);
      try {
        return {
          displayValue: getCellDisplayValue(worksheet, message.payload.row, message.payload.col, targetSheet),
          formula: worksheet.getFormulaAt(message.payload.row, message.payload.col) ?? ""
        };
      } finally {
        worksheet.free();
      }
    }
    case "getRowsBatch": {
      if (!workbook) {
        return null;
      }

      const worksheet = workbook.getSheet(message.payload.workbookSheetIndex) as ReturnType<Workbook["getSheet"]> & {
        getRowsBatch?: (startRow: number, maxRows: number, options?: unknown) => unknown;
      };
      try {
        if (typeof worksheet.getRowsBatch !== "function") return null;
        return worksheet.getRowsBatch(message.payload.startRow, message.payload.rowCount, {
          includeFormulas: true,
          includeHyperlinks: true,
          includeMergeInfo: true,
          includeStyles: true,
          useFormattedValues: true
        }) as unknown[] | null;
      } finally {
        worksheet.free();
      }
    }
    default:
      return null;
  }
}

self.addEventListener("message", (event: MessageEvent<WorkerRequest>) => {
  const message = event.data;
  void handleMessage(message)
    .then((result) => {
      respond({
        id: message.id,
        result,
        success: true
      });
    })
    .catch((error: unknown) => {
      const officeError = error instanceof OfficeLoadError ? error : undefined;
      respond({
        error: {
          code: officeError?.code === "LIMIT_EXCEEDED"
            ? "LIMIT_EXCEEDED"
            : officeError?.code === "TIMEOUT"
              ? "TIMEOUT"
              : officeError?.code === "INVALID_SOURCE"
                ? "INVALID_SOURCE"
                : "WORKER_FAILED",
          message: error instanceof Error ? error.message : "Worker request failed.",
          ...(officeError?.phase ? { phase: officeError.phase } : {}),
          ...(officeError?.limit ? { limit: officeError.limit } : {}),
          ...(officeError?.actual !== undefined ? { actual: officeError.actual } : {}),
          ...(officeError?.allowed !== undefined ? { allowed: officeError.allowed } : {}),
        },
        id: message.id,
        success: false
      } satisfies WorkerErrorResponse);
    });
});
