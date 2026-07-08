import {
  unzipSync,
  zipSync,
  strFromU8,
  strToU8
} from "fflate";
import type {
  XlsxCellRange,
  XlsxResolvedCellStyle,
  XlsxSheetData
} from "@extend-ai/xlsx-core";
import {
  MAX_INTERACTIVE_SHARED_STRINGS_BYTES,
  MAX_INTERACTIVE_TOTAL_XML_BYTES,
  MAX_INTERACTIVE_WORKSHEET_XML_BYTES
} from "./internal";

export type ZipEntryMetadata = {
  compressedSize: number;
  name: string;
  uncompressedSize: number;
};

export type WorkbookPreflightResult = {
  largestWorksheetXmlBytes: number;
  sharedStringsBytes: number;
  totalWorksheetXmlBytes: number;
  tooLarge: boolean;
};

export function formatBinaryBytes(bytes: number) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  return `${value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unitIndex]}`;
}

export function createWorkbookTooLargeError(preflight: WorkbookPreflightResult) {
  return new Error(
    `XLSX is too large to preview interactively. `
    + `Largest worksheet XML: ${formatBinaryBytes(preflight.largestWorksheetXmlBytes)}; `
    + `shared strings: ${formatBinaryBytes(preflight.sharedStringsBytes)}.`
  );
}

export function findZipEndOfCentralDirectoryOffset(bytes: Uint8Array) {
  const minLength = 22;
  if (bytes.byteLength < minLength) {
    return -1;
  }

  const searchStart = Math.max(0, bytes.byteLength - (0xffff + minLength));
  for (let offset = bytes.byteLength - minLength; offset >= searchStart; offset -= 1) {
    if (
      bytes[offset] === 0x50 &&
      bytes[offset + 1] === 0x4b &&
      bytes[offset + 2] === 0x05 &&
      bytes[offset + 3] === 0x06
    ) {
      return offset;
    }
  }

  return -1;
}

export function readZipCentralDirectoryEntries(buffer: ArrayBuffer): ZipEntryMetadata[] | null {
  const bytes = new Uint8Array(buffer);
  const eocdOffset = findZipEndOfCentralDirectoryOffset(bytes);
  if (eocdOffset < 0) {
    return null;
  }

  const view = new DataView(buffer, bytes.byteOffset, bytes.byteLength);
  const centralDirectorySize = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  const decoder = new TextDecoder();
  const entries: ZipEntryMetadata[] = [];

  let offset = centralDirectoryOffset;
  const endOffset = centralDirectoryOffset + centralDirectorySize;
  while (offset + 46 <= endOffset && offset + 46 <= bytes.byteLength) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      return null;
    }

    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const fileNameStart = offset + 46;
    const fileNameEnd = fileNameStart + fileNameLength;
    if (fileNameEnd > bytes.byteLength) {
      return null;
    }

    entries.push({
      compressedSize,
      name: decoder.decode(bytes.subarray(fileNameStart, fileNameEnd)),
      uncompressedSize
    });

    offset = fileNameEnd + extraLength + commentLength;
  }

  return entries;
}

export function preflightWorkbookBuffer(buffer: ArrayBuffer): WorkbookPreflightResult | null {
  const entries = readZipCentralDirectoryEntries(buffer);
  if (!entries) {
    return null;
  }

  let largestWorksheetXmlBytes = 0;
  let totalWorksheetXmlBytes = 0;
  let sharedStringsBytes = 0;

  for (const entry of entries) {
    if (/^xl\/worksheets\/[^/]+\.xml$/i.test(entry.name)) {
      largestWorksheetXmlBytes = Math.max(largestWorksheetXmlBytes, entry.uncompressedSize);
      totalWorksheetXmlBytes += entry.uncompressedSize;
      continue;
    }

    if (entry.name === "xl/sharedStrings.xml") {
      sharedStringsBytes = entry.uncompressedSize;
    }
  }

  const tooLarge =
    largestWorksheetXmlBytes > MAX_INTERACTIVE_WORKSHEET_XML_BYTES ||
    sharedStringsBytes > MAX_INTERACTIVE_SHARED_STRINGS_BYTES ||
    totalWorksheetXmlBytes + sharedStringsBytes > MAX_INTERACTIVE_TOTAL_XML_BYTES;

  return {
    largestWorksheetXmlBytes,
    sharedStringsBytes,
    tooLarge,
    totalWorksheetXmlBytes
  };
}

export class XlsxFileSizeLimitExceededError extends Error {
  fileSizeBytes: number;
  maxFileSizeBytes: number;

  constructor(fileSizeBytes: number, maxFileSizeBytes: number) {
    super(
      `XLSX file size ${formatBinaryBytes(fileSizeBytes)} exceeds the configured limit of ${formatBinaryBytes(maxFileSizeBytes)}.`
    );
    this.name = "XlsxFileSizeLimitExceededError";
    this.fileSizeBytes = fileSizeBytes;
    this.maxFileSizeBytes = maxFileSizeBytes;
  }
}

export function fileStem(fileName: string): string {
  const normalized = fileName.trim();
  const lastDot = normalized.lastIndexOf(".");
  return lastDot > 0 ? normalized.slice(0, lastDot) : normalized;
}

export function pxToSheetRowHeight(heightPx: number): number {
  return Math.max(heightPx, 16) / 1.33;
}

export function cssColor(color: Record<string, unknown> | undefined): string | null {
  if (!color?.hex) {
    return null;
  }

  const hex = String(color.hex);
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  return `#${rgb}`;
}

export function mapBorder(edge: { style: string; color?: { hex?: string } }): string {
  const color = cssColor(edge.color as Record<string, unknown> | undefined) ?? "#000000";
  const widthMap: Record<string, string> = {
    dashed: "1px",
    dotted: "1px",
    double: "3px",
    hair: "1px",
    medium: "2px",
    thick: "3px",
    thin: "1px"
  };
  const styleMap: Record<string, string> = {
    dashDot: "dashed",
    dashDotDot: "dotted",
    dashed: "dashed",
    dotted: "dotted",
    double: "double",
    hair: "solid",
    medium: "solid",
    mediumDashDot: "dashed",
    mediumDashDotDot: "dotted",
    mediumDashed: "dashed",
    slantDashDot: "dashed",
    thick: "solid",
    thin: "solid"
  };

  return `${widthMap[edge.style] ?? "1px"} ${styleMap[edge.style] ?? "solid"} ${color}`;
}

export function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&#34;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

export function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function cloneBytes(bytes: Uint8Array): Uint8Array {
  const nextBytes = new Uint8Array(bytes.byteLength);
  nextBytes.set(bytes);
  return nextBytes;
}

export function sanitizeSavedWorkbookBytes(bytes: Uint8Array): Uint8Array {
  try {
    const archive = unzipSync(bytes);
    const stylesEntry = archive["xl/styles.xml"];
    if (stylesEntry) {
      const stylesXml = strFromU8(stylesEntry)
        .replace(/&amp;quot;/g, "&quot;")
        .replace(/&amp;apos;/g, "&apos;");
      archive["xl/styles.xml"] = strToU8(stylesXml);
    }

    return zipSync(archive, { level: 6 });
  } catch {
    return cloneBytes(bytes);
  }
}

export function resolveDisplayFileName(src?: string, fileName?: string): string {
  if (typeof fileName === "string" && fileName.trim().length > 0) {
    return fileName.trim();
  }

  if (!src) {
    return "Workbook.xlsx";
  }

  const pathWithoutQuery = src.split("?")[0] ?? "";
  const pathSegments = pathWithoutQuery.split("/");
  const lastSegment = pathSegments[pathSegments.length - 1] ?? "";

  if (!lastSegment) {
    return "Workbook.xlsx";
  }

  try {
    return decodeURIComponent(lastSegment);
  } catch {
    return lastSegment;
  }
}

// Re-exported for type-only consumers that need the cell-range shape alongside
// formatting helpers (e.g. clipboard HTML builders).
export type { XlsxCellRange };

export function resolveInheritedCellStyle(sheet: XlsxSheetData | null | undefined, row: number, col: number): XlsxResolvedCellStyle | null {
  if (!sheet) {
    return null;
  }

  const rowStyleId = sheet.rowStyleIds[row];
  if (rowStyleId !== undefined) {
    return sheet.styleById[rowStyleId] ?? null;
  }

  const colStyleId = sheet.colStyleIds[col];
  if (colStyleId !== undefined) {
    return sheet.styleById[colStyleId] ?? null;
  }

  return null;
}
