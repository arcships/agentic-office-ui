import {
  assertOfficeInputBytes,
  resolveOfficeLimits,
  validateOfficeArchive,
} from "@arcships/office-runtime";
import { unzipSync } from "fflate";
import { validateXlsxImageAssets } from "./images/image-budget";

const MIB = 1024 * 1024;

/** Public, instance-owned XLSX resource limits. */
export interface XlsxRuntimeLimits {
  maxInputBytes?: number;
  maxArchiveEntries?: number;
  maxUncompressedBytes?: number;
  maxSingleEntryBytes?: number;
  maxCompressionRatio?: number;
  maxArchivePathLength?: number;
  maxXmlBytes?: number;
  maxSingleXmlBytes?: number;
  maxXmlDepth?: number;
  maxXmlAttributeBytes?: number;
  maxTextNodeBytes?: number;
  maxRelationships?: number;
  maxWorksheetXmlBytes?: number;
  maxSharedStringsBytes?: number;
  maxWorksheetRows?: number;
  maxWorksheetColumns?: number;
  maxWorksheets?: number;
  maxSharedStrings?: number;
  maxFormulaCount?: number;
  maxSingleImageBytes?: number;
  maxTotalImageBytes?: number;
  maxImageWidth?: number;
  maxImageHeight?: number;
  maxSingleImagePixels?: number;
  maxTotalImagePixels?: number;
  maxConcurrentImageDecodes?: number;
  maxParseMs?: number;
}

/** Public archive entry summary returned by XLSX preflight validation. */
export interface XlsxArchiveEntry {
  name: string;
  compressedBytes: number;
  uncompressedBytes: number;
  compression: number;
}

/** Public archive summary returned before an XLSX workbook is parsed. */
export interface XlsxArchiveValidationResult {
  sourceFormat: "xlsx" | "xlsb" | "xlsm" | "xltx" | "xltm";
  entryCount: number;
  compressedBytes: number;
  uncompressedBytes: number;
  maxCompressionRatio: number;
  entries: readonly XlsxArchiveEntry[];
  xmlBytes: number;
  relationships: number;
  worksheets: number;
  sharedStrings: number;
  formulas: number;
}

/** Public image-budget snapshot with no dependency on the private runtime. */
export interface XlsxImageBudgetSnapshot {
  compressedBytes: number;
  pixels: number;
  activeDecodes: number;
  pendingDecodes: number;
  disposed: boolean;
}

export const DEFAULT_XLSX_RUNTIME_LIMITS: Readonly<XlsxRuntimeLimits> = Object.freeze({
  maxInputBytes: 25 * MIB,
  maxArchiveEntries: 10_000,
  maxUncompressedBytes: 256 * MIB,
  maxSingleEntryBytes: 200 * MIB,
  maxCompressionRatio: 100,
  maxArchivePathLength: 1_024,
  maxXmlBytes: 256 * MIB,
  maxSingleXmlBytes: 200 * MIB,
  maxXmlDepth: 256,
  maxXmlAttributeBytes: 1 * MIB,
  maxTextNodeBytes: 16 * MIB,
  maxRelationships: 100_000,
  maxWorksheetXmlBytes: 200 * MIB,
  maxSharedStringsBytes: 50 * MIB,
  maxWorksheetRows: 1_048_576,
  maxWorksheetColumns: 16_384,
  maxWorksheets: 200,
  maxSharedStrings: 1_000_000,
  maxFormulaCount: 1_000_000,
  maxSingleImageBytes: 25 * MIB,
  maxTotalImageBytes: 100 * MIB,
  maxImageWidth: 32_768,
  maxImageHeight: 32_768,
  maxSingleImagePixels: 40_000_000,
  maxTotalImagePixels: 100_000_000,
  maxConcurrentImageDecodes: 4,
  maxParseMs: 30_000,
});

const XLSX_RUNTIME_HARD_LIMITS: Readonly<XlsxRuntimeLimits> = Object.freeze({
  maxInputBytes: 100 * MIB,
  maxArchiveEntries: 50_000,
  maxUncompressedBytes: 512 * MIB,
  maxSingleEntryBytes: 256 * MIB,
  maxCompressionRatio: 200,
  maxArchivePathLength: 4_096,
  maxXmlBytes: 512 * MIB,
  maxSingleXmlBytes: 256 * MIB,
  maxXmlDepth: 512,
  maxXmlAttributeBytes: 4 * MIB,
  maxTextNodeBytes: 64 * MIB,
  maxRelationships: 500_000,
  maxWorksheetXmlBytes: 256 * MIB,
  maxSharedStringsBytes: 100 * MIB,
  maxWorksheetRows: 1_048_576,
  maxWorksheetColumns: 16_384,
  maxWorksheets: 10_000,
  maxSharedStrings: 5_000_000,
  maxFormulaCount: 5_000_000,
  maxSingleImageBytes: 100 * MIB,
  maxTotalImageBytes: 512 * MIB,
  maxImageWidth: 65_535,
  maxImageHeight: 65_535,
  maxSingleImagePixels: 100_000_000,
  maxTotalImagePixels: 250_000_000,
  maxConcurrentImageDecodes: 16,
  maxParseMs: 120_000,
});

export function resolveXlsxRuntimeLimits(
  limits: XlsxRuntimeLimits | undefined,
): Readonly<XlsxRuntimeLimits> {
  return resolveOfficeLimits(
    DEFAULT_XLSX_RUNTIME_LIMITS,
    limits,
    XLSX_RUNTIME_HARD_LIMITS,
  ) as Readonly<XlsxRuntimeLimits>;
}

export function isLegacyXlsBytes(input: ArrayBuffer | Uint8Array): boolean {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
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

export function isBinaryXlsbBytes(input: ArrayBuffer | Uint8Array): boolean {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const marker = "xl/workbook.bin";
  if (bytes.byteLength < marker.length) return false;
  for (let offset = 0; offset <= bytes.byteLength - marker.length; offset += 1) {
    let matched = true;
    for (let index = 0; index < marker.length; index += 1) {
      if (bytes[offset + index] !== marker.charCodeAt(index)) {
        matched = false;
        break;
      }
    }
    if (matched) return true;
  }
  return false;
}

export function validateXlsxArchive(
  input: ArrayBuffer | Uint8Array,
  limits?: XlsxRuntimeLimits,
  signal?: AbortSignal,
): XlsxArchiveValidationResult | null {
  const resolvedLimits = resolveXlsxRuntimeLimits(limits);
  assertOfficeInputBytes(input.byteLength, resolvedLimits);
  if (isLegacyXlsBytes(input)) return null;
  const archiveResult = validateOfficeArchive(input, resolvedLimits, {
    format: "xlsx",
    signal,
  });
  if (signal?.aborted) {
    throw new DOMException("XLSX 图片验证已取消。", "AbortError");
  }
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  validateXlsxImageAssets(unzipSync(bytes), resolvedLimits);
  return archiveResult as XlsxArchiveValidationResult;
}
