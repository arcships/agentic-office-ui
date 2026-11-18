import type {
  XlsxAxisRange,
  XlsxCellRange,
  XlsxPageBreak,
  XlsxPageMargins,
  XlsxPageOrder,
  XlsxPageSetup,
  XlsxPrintOptions,
  XlsxPrintTitles,
} from "../types";
import {
  type ArchiveEntries,
  getChildElements,
  getFirstDescendant,
  getLocalElements,
  parseA1RangeReference,
  parseXml,
  readArchiveText,
} from "../images/image-parser";

// ── Defaults ─────────────────────────────────────────────────────────────

/** Excel's default page margins in inches. */
export const DEFAULT_PAGE_MARGINS: Readonly<XlsxPageMargins> = {
  bottom: 0.75,
  footer: 0.3,
  header: 0.3,
  left: 0.7,
  right: 0.7,
  top: 0.75,
};

/** Letter portrait at 100%, matching the OOXML defaults. */
export const DEFAULT_PAGE_SETUP: Readonly<XlsxPageSetup> = {
  blackAndWhite: false,
  cellComments: "none",
  copies: 1,
  draft: false,
  errors: "displayed",
  firstPageNumber: 1,
  fitToHeight: 1,
  fitToPage: false,
  fitToWidth: 1,
  orientation: "portrait",
  pageOrder: "downThenOver",
  paperSize: 1,
  scale: 100,
  useFirstPageNumber: false,
};

export const DEFAULT_PRINT_OPTIONS: Readonly<XlsxPrintOptions> = {
  gridLines: false,
  headings: false,
  horizontalCentered: false,
  verticalCentered: false,
};

export function defaultPrintTitles(): XlsxPrintTitles {
  return { columns: [], rows: [] };
}

// ── Attribute helpers ────────────────────────────────────────────────────

function readNumberAttribute(node: Element | null, name: string) {
  if (!node) return null;
  const raw = node.getAttribute(name);
  if (raw == null || raw.trim() === "") return null;
  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

function readBooleanAttribute(node: Element | null, name: string) {
  if (!node) return null;
  const raw = node.getAttribute(name);
  if (raw == null) return null;
  return raw === "1" || raw.toLowerCase() === "true";
}

function readStringAttribute(node: Element | null, name: string) {
  if (!node) return null;
  const raw = node.getAttribute(name);
  return raw == null || raw.trim() === "" ? null : raw;
}

function clampNumber(value: number, minimum: number, maximum: number) {
  return Math.min(maximum, Math.max(minimum, value));
}

function normalizeOrientation(value: unknown): XlsxPageSetup["orientation"] {
  return String(value).toLowerCase() === "landscape" ? "landscape" : "portrait";
}

function normalizePageOrder(value: unknown): XlsxPageOrder {
  return String(value).toLowerCase() === "overthendown" ? "overThenDown" : "downThenOver";
}

function normalizeCellComments(value: unknown): XlsxPageSetup["cellComments"] {
  const normalized = String(value ?? "").toLowerCase();
  if (normalized === "asdisplayed") return "asDisplayed";
  if (normalized === "atend") return "atEnd";
  return "none";
}

function normalizeErrors(value: unknown): XlsxPageSetup["errors"] {
  const normalized = String(value ?? "").toLowerCase();
  return normalized === "blank" || normalized === "dash" || normalized === "na"
    ? normalized === "na" ? "NA" : normalized as XlsxPageSetup["errors"]
    : "displayed";
}

function positiveIntegerOr(value: number | null, fallback: number, maximum = 32767) {
  return value != null && value > 0 ? Math.min(maximum, Math.floor(value)) : fallback;
}

// ── Worksheet-level parsing ──────────────────────────────────────────────

/** Parsed print fields of a single worksheet part. */
export interface WorksheetPrintState {
  pageSetup: XlsxPageSetup;
  printOptions: XlsxPrintOptions;
  pageMargins: XlsxPageMargins;
  rowBreaks: XlsxPageBreak[];
  colBreaks: XlsxPageBreak[];
}

function parseBreaks(document: XMLDocument, kind: "rowBreaks" | "colBreaks"): XlsxPageBreak[] {
  const container = getLocalElements(document, kind)[0];
  if (!container) return [];
  return getChildElements(container, "brk").flatMap((node) => {
    const id = readNumberAttribute(node, "id");
    if (id == null || id < 0) return [];
    return [{
      id: Math.floor(id),
      min: Math.max(0, Math.floor(readNumberAttribute(node, "min") ?? 0)),
      max: Math.max(0, Math.floor(readNumberAttribute(node, "max") ?? (kind === "rowBreaks" ? 16383 : 1048575))),
      manual: readBooleanAttribute(node, "man") ?? true,
    }];
  });
}

function parsePageMargins(document: XMLDocument): XlsxPageMargins {
  const node = getLocalElements(document, "pageMargins")[0] ?? null;
  return {
    bottom: readNumberAttribute(node, "bottom") ?? DEFAULT_PAGE_MARGINS.bottom,
    footer: readNumberAttribute(node, "footer") ?? DEFAULT_PAGE_MARGINS.footer,
    header: readNumberAttribute(node, "header") ?? DEFAULT_PAGE_MARGINS.header,
    left: readNumberAttribute(node, "left") ?? DEFAULT_PAGE_MARGINS.left,
    right: readNumberAttribute(node, "right") ?? DEFAULT_PAGE_MARGINS.right,
    top: readNumberAttribute(node, "top") ?? DEFAULT_PAGE_MARGINS.top,
  };
}

function parsePrintOptions(document: XMLDocument): XlsxPrintOptions {
  const node = getLocalElements(document, "printOptions")[0] ?? null;
  return {
    gridLines: readBooleanAttribute(node, "gridLines") ?? DEFAULT_PRINT_OPTIONS.gridLines,
    headings: readBooleanAttribute(node, "headings") ?? DEFAULT_PRINT_OPTIONS.headings,
    horizontalCentered: readBooleanAttribute(node, "horizontalCentered") ?? DEFAULT_PRINT_OPTIONS.horizontalCentered,
    verticalCentered: readBooleanAttribute(node, "verticalCentered") ?? DEFAULT_PRINT_OPTIONS.verticalCentered,
  };
}

function parsePageSetup(document: XMLDocument): XlsxPageSetup {
  const setup = getLocalElements(document, "pageSetup")[0] ?? null;
  const fitToPagePr = getFirstDescendant(document, "pageSetUpPr");
  const fitToWidth = positiveIntegerOr(readNumberAttribute(setup, "fitToWidth"), 1);
  const fitToHeight = positiveIntegerOr(readNumberAttribute(setup, "fitToHeight"), 1);
  return {
    blackAndWhite: readBooleanAttribute(setup, "blackAndWhite") ?? false,
    cellComments: normalizeCellComments(readStringAttribute(setup, "cellComments")),
    copies: positiveIntegerOr(readNumberAttribute(setup, "copies"), 1, 32767),
    draft: readBooleanAttribute(setup, "draft") ?? false,
    errors: normalizeErrors(readStringAttribute(setup, "errors")),
    firstPageNumber: positiveIntegerOr(readNumberAttribute(setup, "firstPageNumber"), 1, 32767),
    fitToHeight,
    fitToPage: readBooleanAttribute(fitToPagePr, "fitToPage") ?? false,
    fitToWidth,
    horizontalDpi: readNumberAttribute(setup, "horizontalDpi") ?? undefined,
    orientation: normalizeOrientation(readStringAttribute(setup, "orientation")),
    pageOrder: normalizePageOrder(readStringAttribute(setup, "pageOrder")),
    paperSize: positiveIntegerOr(readNumberAttribute(setup, "paperSize"), DEFAULT_PAGE_SETUP.paperSize, 118),
    scale: positiveIntegerOr(readNumberAttribute(setup, "scale"), DEFAULT_PAGE_SETUP.scale, 400),
    useFirstPageNumber: readBooleanAttribute(setup, "useFirstPageNumber") ?? false,
    verticalDpi: readNumberAttribute(setup, "verticalDpi") ?? undefined,
  };
}

/**
 * Parses `pageSetup`, `printOptions`, `pageMargins`, `rowBreaks`, `colBreaks`
 * and `sheetPr/pageSetUpPr` from a worksheet document.
 */
export function parseWorksheetPrintState(document: XMLDocument): WorksheetPrintState {
  return {
    colBreaks: parseBreaks(document, "colBreaks"),
    pageMargins: parsePageMargins(document),
    pageSetup: parsePageSetup(document),
    printOptions: parsePrintOptions(document),
    rowBreaks: parseBreaks(document, "rowBreaks"),
  };
}

// ── Workbook defined names ───────────────────────────────────────────────

/** Print metadata resolved from workbook defined names for one sheet. */
export interface WorkbookPrintMetadata {
  printAreas: XlsxCellRange[];
  printTitles: XlsxPrintTitles;
}

function splitReferenceList(value: string) {
  const references: string[] = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index] ?? "";
    if (quoted) {
      current += character;
      if (character === "'") {
        if (value[index + 1] === "'") {
          current += "'";
          index += 1;
        } else {
          quoted = false;
        }
      }
      continue;
    }
    if (character === "'") {
      quoted = true;
      current += character;
      continue;
    }
    if (character === ",") {
      if (current.trim()) references.push(current.trim());
      current = "";
      continue;
    }
    current += character;
  }
  if (current.trim()) references.push(current.trim());
  return references;
}

function unquoteSheetName(value: string) {
  const trimmed = value.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

/** Removes the optional `'Sheet name'!` qualifier and validates the sheet. */
function stripSheetQualifier(reference: string, sheetName: string) {
  let sheetPart = "";
  let rangePart = reference;
  let quoted = false;
  for (let index = 0; index < reference.length; index += 1) {
    const character = reference[index] ?? "";
    if (character === "'") {
      if (quoted && reference[index + 1] === "'") {
        index += 1;
        continue;
      }
      quoted = !quoted;
      continue;
    }
    if (character === "!" && !quoted) {
      sheetPart = reference.slice(0, index);
      rangePart = reference.slice(index + 1);
      break;
    }
  }
  if (sheetPart) {
    const qualifiedName = unquoteSheetName(sheetPart);
    if (qualifiedName.localeCompare(sheetName, undefined, { sensitivity: "accent" }) !== 0) {
      return null;
    }
  }
  return rangePart.trim();
}

function parsePrintAreaReferences(value: string, sheetName: string): XlsxCellRange[] {
  return splitReferenceList(value).flatMap((reference) => {
    const rangePart = stripSheetQualifier(reference, sheetName);
    if (!rangePart) return [];
    const range = parseA1RangeReference(rangePart.replace(/\$/g, ""));
    return range ? [range] : [];
  });
}

function toAxisRange(start: number, end: number): XlsxAxisRange {
  return { end: Math.max(start, end), start: Math.min(start, end) };
}

function parsePrintTitleReferences(value: string, sheetName: string) {
  const rows: XlsxAxisRange[] = [];
  const columns: XlsxAxisRange[] = [];
  for (const reference of splitReferenceList(value)) {
    const rangePart = stripSheetQualifier(reference, sheetName);
    if (!rangePart) continue;
    const normalized = rangePart.replace(/\$/g, "").toUpperCase();
    const rowRange = /^(\d+):(\d+)$/.exec(normalized);
    if (rowRange) {
      const start = Number(rowRange[1]) - 1;
      const end = Number(rowRange[2]) - 1;
      if (start >= 0 && end >= 0) rows.push(toAxisRange(start, end));
      continue;
    }
    const columnRange = /^([A-Z]+):([A-Z]+)$/.exec(normalized);
    if (columnRange) {
      const start = columnIndex(columnRange[1] ?? "");
      const end = columnIndex(columnRange[2] ?? "");
      if (start != null && end != null) columns.push(toAxisRange(start, end));
    }
  }
  return { columns, rows };
}

function columnIndex(label: string) {
  let value = 0;
  for (const character of label) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value > 0 ? value - 1 : null;
}

/**
 * Reads `_xlnm.Print_Area` / `_xlnm.Print_Titles` from `xl/workbook.xml` and
 * assigns each `localSheetId` its own print metadata.
 */
export function parseWorkbookPrintMetadata(
  archive: ArchiveEntries,
  sheets: ReadonlyArray<{ name: string }>,
): WorkbookPrintMetadata[] {
  const metadata: WorkbookPrintMetadata[] = sheets.map(() => ({
    printAreas: [],
    printTitles: defaultPrintTitles(),
  }));
  const xml = readArchiveText(archive, "xl/workbook.xml");
  const document = xml ? parseXml(xml) : null;
  if (!document) return metadata;

  for (const node of getLocalElements(document, "definedName")) {
    const name = (node.getAttribute("name") ?? "").trim().toLowerCase();
    const localSheetId = Number(node.getAttribute("localSheetId") ?? Number.NaN);
    if (!Number.isInteger(localSheetId) || localSheetId < 0 || localSheetId >= sheets.length) {
      continue;
    }
    const target = metadata[localSheetId];
    const sheetName = sheets[localSheetId]?.name ?? "";
    const value = (node.textContent ?? "").trim();
    if (!target || !value) continue;
    if (name === "_xlnm.print_area") {
      target.printAreas.push(...parsePrintAreaReferences(value, sheetName));
    } else if (name === "_xlnm.print_titles") {
      const titles = parsePrintTitleReferences(value, sheetName);
      target.printTitles.rows.push(...titles.rows);
      target.printTitles.columns.push(...titles.columns);
    }
  }
  return metadata;
}

// ── Engine fallback ──────────────────────────────────────────────────────

/** Structural subset of `@dukelib/sheets-wasm`'s Worksheet used here. */
export interface WorksheetPrintSource {
  pageSetup?: unknown;
  printArea?: string | undefined;
  repeatRows?: unknown;
  repeatCols?: unknown;
  rowBreaks?: unknown;
  colBreaks?: unknown;
}

function objectValues(value: unknown): unknown[] {
  if (Array.isArray(value)) return value;
  if (ArrayBuffer.isView(value)) {
    return Array.from(value as unknown as ArrayLike<unknown>);
  }
  if (value && typeof value === "object") {
    return Object.values(value as Record<string, unknown>);
  }
  return [];
}

function indicesToAxisRanges(indices: number[]): XlsxAxisRange[] {
  const sorted = [...new Set(indices.filter((index) => Number.isInteger(index) && index >= 0))].sort((left, right) => left - right);
  const ranges: XlsxAxisRange[] = [];
  for (const index of sorted) {
    const previous = ranges[ranges.length - 1];
    if (previous && index === previous.end + 1) previous.end = index;
    else ranges.push({ end: index, start: index });
  }
  return ranges;
}

function enginePageSetup(worksheet: WorksheetPrintSource | null | undefined): XlsxPageSetup {
  const raw = worksheet?.pageSetup && typeof worksheet.pageSetup === "object"
    ? worksheet.pageSetup as Record<string, unknown>
    : {};
  const fitToWidth = positiveIntegerOr(typeof raw.fitToWidth === "number" ? raw.fitToWidth : null, 1);
  const fitToHeight = positiveIntegerOr(typeof raw.fitToHeight === "number" ? raw.fitToHeight : null, 1);
  return {
    blackAndWhite: raw.blackAndWhite === true,
    cellComments: normalizeCellComments(raw.cellComments),
    copies: positiveIntegerOr(typeof raw.copies === "number" ? raw.copies : null, 1, 32767),
    draft: raw.draft === true,
    errors: normalizeErrors(raw.errors),
    firstPageNumber: positiveIntegerOr(typeof raw.firstPageNumber === "number" ? raw.firstPageNumber : null, 1, 32767),
    fitToHeight,
    fitToPage: raw.fitToPage === true || fitToWidth !== 1 || fitToHeight !== 1,
    fitToWidth,
    horizontalDpi: typeof raw.horizontalDpi === "number" ? raw.horizontalDpi : undefined,
    orientation: normalizeOrientation(raw.orientation),
    pageOrder: normalizePageOrder(raw.pageOrder),
    paperSize: positiveIntegerOr(typeof raw.paperSize === "number" ? raw.paperSize : null, DEFAULT_PAGE_SETUP.paperSize, 118),
    scale: positiveIntegerOr(typeof raw.scale === "number" ? raw.scale : null, DEFAULT_PAGE_SETUP.scale, 400),
    useFirstPageNumber: raw.useFirstPageNumber === true,
    verticalDpi: typeof raw.verticalDpi === "number" ? raw.verticalDpi : undefined,
  };
}

function enginePageMargins(worksheet: WorksheetPrintSource | null | undefined): XlsxPageMargins {
  const raw = worksheet?.pageSetup && typeof worksheet.pageSetup === "object"
    ? worksheet.pageSetup as Record<string, unknown>
    : {};
  const value = (name: string, fallback: number) =>
    typeof raw[name] === "number" && Number.isFinite(raw[name]) ? raw[name] as number : fallback;
  return {
    bottom: value("bottomMargin", DEFAULT_PAGE_MARGINS.bottom),
    footer: value("footerMargin", DEFAULT_PAGE_MARGINS.footer),
    header: value("headerMargin", DEFAULT_PAGE_MARGINS.header),
    left: value("leftMargin", DEFAULT_PAGE_MARGINS.left),
    right: value("rightMargin", DEFAULT_PAGE_MARGINS.right),
    top: value("topMargin", DEFAULT_PAGE_MARGINS.top),
  };
}

function enginePrintOptions(worksheet: WorksheetPrintSource | null | undefined): XlsxPrintOptions {
  const raw = worksheet?.pageSetup && typeof worksheet.pageSetup === "object"
    ? worksheet.pageSetup as Record<string, unknown>
    : {};
  return {
    gridLines: raw.printGridlines === true,
    headings: raw.printHeadings === true,
    horizontalCentered: raw.horizontalCentered === true,
    verticalCentered: raw.verticalCentered === true,
  };
}

function engineBreaks(value: unknown): XlsxPageBreak[] {
  return objectValues(value).flatMap((entry) => {
    if (!entry || typeof entry !== "object") return [];
    const raw = entry as Record<string, unknown>;
    const id = typeof raw.id === "number" ? raw.id : Number(raw.id);
    if (!Number.isFinite(id) || id < 0) return [];
    return [{
      id: Math.floor(id),
      min: typeof raw.min === "number" ? raw.min : 0,
      max: typeof raw.max === "number" ? raw.max : 16383,
      manual: raw.manual !== false,
    }];
  });
}

function enginePrintAreas(worksheet: WorksheetPrintSource | null | undefined): XlsxCellRange[] {
  const raw = worksheet?.printArea;
  if (typeof raw !== "string" || !raw.trim()) return [];
  return raw.split(",").flatMap((reference) => {
    const range = parseA1RangeReference(reference.trim());
    return range ? [range] : [];
  });
}

/**
 * Normalizes XML-parsed print fields or falls back to the workbook engine's
 * `pageSetup` / `printArea` / `repeatRows` / `repeatCols` / break metadata.
 *
 * The worksheet XML is the authoritative source when present. The engine
 * fallback keeps print preview working for `skipXmlParsing` sources and legacy
 * binaries where the worksheet part is unavailable.
 */
export function resolveWorksheetPrintState(
  state: Partial<WorksheetPrintState & WorkbookPrintMetadata> | null | undefined,
  worksheet?: WorksheetPrintSource | null,
): WorksheetPrintState & WorkbookPrintMetadata {
  const pageSetup = state?.pageSetup ?? enginePageSetup(worksheet);
  const printAreas = state?.printAreas ?? enginePrintAreas(worksheet);
  const printTitles = state?.printTitles ?? {
    columns: indicesToAxisRanges(objectValues(worksheet?.repeatCols).map(Number)),
    rows: indicesToAxisRanges(objectValues(worksheet?.repeatRows).map(Number)),
  };
  return {
    colBreaks: state?.colBreaks ?? engineBreaks(worksheet?.colBreaks),
    pageMargins: state?.pageMargins ?? enginePageMargins(worksheet),
    pageSetup: {
      ...DEFAULT_PAGE_SETUP,
      ...pageSetup,
      scale: clampNumber(Math.floor(Number.isFinite(pageSetup.scale) ? pageSetup.scale : DEFAULT_PAGE_SETUP.scale), 10, 400),
    },
    printAreas,
    printOptions: {
      ...DEFAULT_PRINT_OPTIONS,
      ...(state?.printOptions ?? enginePrintOptions(worksheet)),
    },
    printTitles,
    rowBreaks: state?.rowBreaks ?? engineBreaks(worksheet?.rowBreaks),
  };
}
