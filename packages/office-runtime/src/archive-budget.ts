import { Unzip, UnzipInflate } from "fflate";
import { OfficeLoadError } from "./errors";
import type { OfficeFormat } from "./errors";
import type { OfficeLimits } from "./limits";
import { detectOoxmlSourceFormat } from "./source-format";
import type { OoxmlSourceFormat } from "./source-format";

const END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const CENTRAL_DIRECTORY_ENTRY = 0x02014b50;
const MAX_ZIP_COMMENT_BYTES = 0xffff;

export interface OfficeArchiveEntry {
  name: string;
  compressedBytes: number;
  uncompressedBytes: number;
  compression: number;
}

export interface OfficeArchiveSummary {
  entryCount: number;
  compressedBytes: number;
  uncompressedBytes: number;
  maxCompressionRatio: number;
  entries: readonly OfficeArchiveEntry[];
}

export interface OfficeXmlEntrySummary {
  relationships: number;
  rows: number;
  maxRow: number;
  maxColumn: number;
  worksheets: number;
  sharedStrings: number;
  formulas: number;
}

export interface OfficeArchiveValidationOptions {
  format: Exclude<OfficeFormat, "pdf">;
  signal?: AbortSignal;
}

export interface OfficeArchiveValidationResult extends OfficeArchiveSummary {
  sourceFormat: OoxmlSourceFormat;
  xmlBytes: number;
  relationships: number;
  worksheets: number;
  sharedStrings: number;
  formulas: number;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  throw new OfficeLoadError({
    code: "ABORTED",
    message: "文档解析已取消。",
    phase: "archive",
  });
}

function assertLimit(
  actual: number,
  allowed: number | undefined,
  limit: keyof OfficeLimits,
  phase: "input" | "archive" | "xml",
): void {
  if (allowed === undefined || actual <= allowed) return;
  throw new OfficeLoadError({
    code: "LIMIT_EXCEEDED",
    message: `${String(limit)} 的实际值 ${actual} 超过允许值 ${allowed}。`,
    phase,
    limit,
    actual,
    allowed,
  });
}

function invalidArchive(message: string, cause?: unknown): OfficeLoadError {
  return new OfficeLoadError({
    code: "INVALID_SOURCE",
    message,
    phase: "archive",
    cause,
  });
}

function invalidXml(message: string, cause?: unknown): OfficeLoadError {
  return new OfficeLoadError({
    code: "INVALID_SOURCE",
    message,
    phase: "xml",
    cause,
  });
}

function findEndOfCentralDirectory(bytes: Uint8Array): number {
  const minimum = Math.max(0, bytes.byteLength - 22 - MAX_ZIP_COMMENT_BYTES);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = bytes.byteLength - 22; offset >= minimum; offset -= 1) {
    if (view.getUint32(offset, true) === END_OF_CENTRAL_DIRECTORY) return offset;
  }
  throw invalidArchive("文件不是完整的 OOXML 压缩包：找不到目录结束记录。");
}

function decodeEntryName(bytes: Uint8Array): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes).normalize("NFC");
  } catch (error) {
    throw invalidArchive("压缩包中的文件名不是有效的 UTF-8。", error);
  }
}

function validateEntryName(name: string, limits: Readonly<OfficeLimits>): string {
  assertLimit(new TextEncoder().encode(name).byteLength, limits.maxArchivePathLength, "maxArchivePathLength", "archive");
  if (
    !name ||
    name.includes("\0") ||
    name.includes("\\") ||
    name.startsWith("/") ||
    /^[A-Za-z]:/.test(name)
  ) {
    throw invalidArchive(`压缩包包含不安全的文件路径：${name || "<empty>"}。`);
  }
  const segments = name.split("/");
  if (segments.some((segment, index) => segment === ".." || segment === "." || (segment === "" && index < segments.length - 1))) {
    throw invalidArchive(`压缩包包含不安全的文件路径：${name}。`);
  }
  return name;
}

function compressionRatio(uncompressedBytes: number, compressedBytes: number): number {
  if (uncompressedBytes === 0) return 0;
  return uncompressedBytes / Math.max(1, compressedBytes);
}

export function inspectOfficeArchive(
  input: ArrayBuffer | Uint8Array,
  limits: Readonly<OfficeLimits>,
  signal?: AbortSignal,
): OfficeArchiveSummary {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  throwIfAborted(signal);
  if (bytes.byteLength < 22) throw invalidArchive("文件不是完整的 OOXML 压缩包。");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = findEndOfCentralDirectory(bytes);
  const disk = view.getUint16(eocd + 4, true);
  const directoryDisk = view.getUint16(eocd + 6, true);
  const entriesOnDisk = view.getUint16(eocd + 8, true);
  const entryCount = view.getUint16(eocd + 10, true);
  const directoryBytes = view.getUint32(eocd + 12, true);
  const directoryOffset = view.getUint32(eocd + 16, true);
  const commentBytes = view.getUint16(eocd + 20, true);
  if (disk !== 0 || directoryDisk !== 0 || entriesOnDisk !== entryCount) {
    throw invalidArchive("不支持分卷压缩包。");
  }
  if (entryCount === 0xffff || directoryBytes === 0xffffffff || directoryOffset === 0xffffffff) {
    throw invalidArchive("不支持 ZIP64 压缩包。");
  }
  if (eocd + 22 + commentBytes !== bytes.byteLength || directoryOffset + directoryBytes !== eocd) {
    throw invalidArchive("压缩包目录已损坏或被截断。");
  }
  assertLimit(entryCount, limits.maxArchiveEntries, "maxArchiveEntries", "archive");

  const seenNames = new Set<string>();
  const entries: OfficeArchiveEntry[] = [];
  let offset = directoryOffset;
  let totalCompressed = 0;
  let totalUncompressed = 0;
  let largestRatio = 0;
  for (let index = 0; index < entryCount; index += 1) {
    throwIfAborted(signal);
    if (offset + 46 > eocd || view.getUint32(offset, true) !== CENTRAL_DIRECTORY_ENTRY) {
      throw invalidArchive("压缩包目录项已损坏或数量不一致。");
    }
    const flags = view.getUint16(offset + 8, true);
    const compression = view.getUint16(offset + 10, true);
    const compressedBytes = view.getUint32(offset + 20, true);
    const uncompressedBytes = view.getUint32(offset + 24, true);
    const nameBytes = view.getUint16(offset + 28, true);
    const extraBytes = view.getUint16(offset + 30, true);
    const entryCommentBytes = view.getUint16(offset + 32, true);
    const nextOffset = offset + 46 + nameBytes + extraBytes + entryCommentBytes;
    if (nextOffset > eocd) throw invalidArchive("压缩包目录项已被截断。");
    if ((flags & 0x1) !== 0) throw invalidArchive("不支持加密的 OOXML 压缩包。");
    if (compression !== 0 && compression !== 8) throw invalidArchive(`不支持压缩算法 ${compression}。`);
    if (compressedBytes === 0xffffffff || uncompressedBytes === 0xffffffff) {
      throw invalidArchive("不支持 ZIP64 目录项。");
    }
    const name = validateEntryName(
      decodeEntryName(bytes.subarray(offset + 46, offset + 46 + nameBytes)),
      limits,
    );
    if (seenNames.has(name)) throw invalidArchive(`压缩包包含重复文件路径：${name}。`);
    seenNames.add(name);
    assertLimit(uncompressedBytes, limits.maxSingleEntryBytes, "maxSingleEntryBytes", "archive");
    const ratio = compressionRatio(uncompressedBytes, compressedBytes);
    assertLimit(ratio, limits.maxCompressionRatio, "maxCompressionRatio", "archive");
    totalCompressed += compressedBytes;
    totalUncompressed += uncompressedBytes;
    assertLimit(totalUncompressed, limits.maxUncompressedBytes, "maxUncompressedBytes", "archive");
    largestRatio = Math.max(largestRatio, ratio);
    entries.push({ name, compressedBytes, uncompressedBytes, compression });
    offset = nextOffset;
  }
  if (offset !== directoryOffset + directoryBytes) {
    throw invalidArchive("压缩包目录长度与实际内容不一致。");
  }
  return Object.freeze({
    entryCount,
    compressedBytes: totalCompressed,
    uncompressedBytes: totalUncompressed,
    maxCompressionRatio: largestRatio,
    entries: Object.freeze(entries),
  });
}

function decodeXml(bytes: Uint8Array): string {
  let encoding = "utf-8";
  if (bytes[0] === 0xff && bytes[1] === 0xfe) encoding = "utf-16le";
  if (bytes[0] === 0xfe && bytes[1] === 0xff) encoding = "utf-16be";
  try {
    const text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
    const declared = /^\uFEFF?<\?xml[^>]*\bencoding=["']([^"']+)["']/i.exec(text)?.[1]?.toLowerCase();
    if (declared && !["utf-8", "utf8", "utf-16", "utf-16le", "utf-16be"].includes(declared)) {
      throw invalidXml(`不支持 XML 编码 ${declared}。`);
    }
    return text;
  } catch (error) {
    if (error instanceof OfficeLoadError) throw error;
    throw invalidXml("XML 内容使用了无效的字符编码。", error);
  }
}

function findTagEnd(xml: string, start: number): number {
  let quote = "";
  for (let index = start + 1; index < xml.length; index += 1) {
    const char = xml[index] ?? "";
    if (quote) {
      if (char === quote) quote = "";
    } else if (char === "\"" || char === "'") {
      quote = char;
    } else if (char === ">") {
      return index;
    }
  }
  return -1;
}

function columnNumber(label: string): number {
  let value = 0;
  for (const char of label.toUpperCase()) value = value * 26 + char.charCodeAt(0) - 64;
  return value;
}

function xmlLocalName(qualifiedName: string): string {
  const separator = qualifiedName.lastIndexOf(":");
  return separator >= 0 ? qualifiedName.slice(separator + 1) : qualifiedName;
}

function numericXmlAttribute(tag: string, localName: string): number | undefined {
  const escapedName = localName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = new RegExp(
    `(?:^|\\s)(?:[^\\s=/>]+:)?${escapedName}\\s*=\\s*(["'])(\\d+)\\1`,
  ).exec(tag);
  if (!match?.[2]) return undefined;
  const value = Number(match[2]);
  return Number.isSafeInteger(value) ? value : undefined;
}

function cellReferenceXmlAttribute(tag: string): { row: number; column: number } | undefined {
  const match = /(?:^|\s)(?:[^\s=/>]+:)?r\s*=\s*(["'])([A-Za-z]+)(\d+)\1/.exec(tag);
  if (!match?.[2] || !match[3]) return undefined;
  const row = Number(match[3]);
  if (!Number.isSafeInteger(row)) return undefined;
  return { row, column: columnNumber(match[2]) };
}

export function validateOfficeXmlEntry(
  name: string,
  bytes: Uint8Array,
  limits: Readonly<OfficeLimits>,
): OfficeXmlEntrySummary {
  assertLimit(bytes.byteLength, limits.maxSingleXmlBytes, "maxSingleXmlBytes", "xml");
  if (/^xl\/worksheets\/[^/]+\.xml$/i.test(name)) {
    assertLimit(bytes.byteLength, limits.maxWorksheetXmlBytes, "maxWorksheetXmlBytes", "xml");
  }
  if (/^xl\/sharedStrings\.xml$/i.test(name)) {
    assertLimit(bytes.byteLength, limits.maxSharedStringsBytes, "maxSharedStringsBytes", "xml");
  }
  const xml = decodeXml(bytes);
  if (/<!DOCTYPE\b/i.test(xml) || /<!ENTITY\b/i.test(xml)) {
    throw invalidXml(`XML 文件 ${name} 包含被禁止的文档类型或实体声明。`);
  }

  const stack: string[] = [];
  let roots = 0;
  let cursor = 0;
  let relationships = 0;
  let worksheets = 0;
  let sharedStrings = 0;
  let formulas = 0;
  let rows = 0;
  let maxRow = 0;
  let maxColumn = 0;
  const workbookEntry = /^xl\/workbook\.xml$/i.test(name);
  const sharedStringsEntry = /^xl\/sharedStrings\.xml$/i.test(name);
  const formulaEntry = /^xl\/(?:worksheets|chartsheets)\/[^/]+\.xml$/i.test(name);
  const worksheetEntry = /^xl\/worksheets\/[^/]+\.xml$/i.test(name);
  while (cursor < xml.length) {
    const start = xml.indexOf("<", cursor);
    if (start < 0) {
      assertLimit(
        new TextEncoder().encode(xml.slice(cursor)).byteLength,
        limits.maxTextNodeBytes,
        "maxTextNodeBytes",
        "xml",
      );
      break;
    }
    assertLimit(
      new TextEncoder().encode(xml.slice(cursor, start)).byteLength,
      limits.maxTextNodeBytes,
      "maxTextNodeBytes",
      "xml",
    );
    if (xml.startsWith("<!--", start)) {
      const end = xml.indexOf("-->", start + 4);
      if (end < 0) throw invalidXml(`XML 文件 ${name} 的注释未结束。`);
      cursor = end + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", start)) {
      const end = xml.indexOf("]]>", start + 9);
      if (end < 0) throw invalidXml(`XML 文件 ${name} 的 CDATA 未结束。`);
      assertLimit(
        new TextEncoder().encode(xml.slice(start + 9, end)).byteLength,
        limits.maxTextNodeBytes,
        "maxTextNodeBytes",
        "xml",
      );
      cursor = end + 3;
      continue;
    }
    const end = findTagEnd(xml, start);
    if (end < 0) throw invalidXml(`XML 文件 ${name} 的标签未结束。`);
    const tag = xml.slice(start, end + 1);
    assertLimit(
      new TextEncoder().encode(tag).byteLength,
      limits.maxXmlAttributeBytes,
      "maxXmlAttributeBytes",
      "xml",
    );
    cursor = end + 1;
    if (tag.startsWith("<?") || tag.startsWith("<!")) continue;
    const closing = /^<\s*\/\s*([^\s>]+)/.exec(tag)?.[1];
    if (closing) {
      if (stack.pop() !== closing) throw invalidXml(`XML 文件 ${name} 的标签嵌套不正确。`);
      continue;
    }
    const opening = /^<\s*([^\s/>]+)/.exec(tag)?.[1];
    if (!opening) throw invalidXml(`XML 文件 ${name} 包含无效标签。`);
    const localName = xmlLocalName(opening);
    if (localName === "Relationship") relationships += 1;
    if (workbookEntry && localName === "sheet") worksheets += 1;
    if (sharedStringsEntry && localName === "si") sharedStrings += 1;
    if (formulaEntry && localName === "f") formulas += 1;
    if (worksheetEntry && localName === "row") {
      rows += 1;
      maxRow = Math.max(maxRow, numericXmlAttribute(tag, "r") ?? rows);
    }
    if (worksheetEntry && localName === "c") {
      const reference = cellReferenceXmlAttribute(tag);
      if (reference) {
        maxColumn = Math.max(maxColumn, reference.column);
        maxRow = Math.max(maxRow, reference.row);
      }
    }
    if (stack.length === 0) roots += 1;
    if (!/\/\s*>$/.test(tag)) {
      stack.push(opening);
      assertLimit(stack.length, limits.maxXmlDepth, "maxXmlDepth", "xml");
    }
  }
  if (stack.length > 0 || roots !== 1) throw invalidXml(`XML 文件 ${name} 不是完整的单根文档。`);
  if (worksheetEntry) {
    assertLimit(Math.max(rows, maxRow), limits.maxWorksheetRows, "maxWorksheetRows", "xml");
    assertLimit(maxColumn, limits.maxWorksheetColumns, "maxWorksheetColumns", "xml");
  }
  return { relationships, rows, maxRow, maxColumn, worksheets, sharedStrings, formulas };
}

function concatChunks(chunks: readonly Uint8Array[], length: number): Uint8Array {
  if (chunks.length === 1 && chunks[0]?.byteLength === length) return chunks[0];
  const result = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return result;
}

function isXmlEntry(name: string): boolean {
  return /(?:\.xml|\.rels)$/i.test(name) || name === "[Content_Types].xml";
}

export function validateOfficeArchive(
  input: ArrayBuffer | Uint8Array,
  limits: Readonly<OfficeLimits>,
  options: OfficeArchiveValidationOptions,
): OfficeArchiveValidationResult {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  assertLimit(bytes.byteLength, limits.maxInputBytes, "maxInputBytes", "input");
  const central = inspectOfficeArchive(bytes, limits, options.signal);
  let entryCount = 0;
  let totalUncompressed = 0;
  let totalXmlBytes = 0;
  let relationships = 0;
  let worksheets = 0;
  let sharedStrings = 0;
  let formulas = 0;
  let completedEntries = 0;
  let contentTypes = "";
  let failure: unknown;
  const seenNames = new Set<string>();
  const unzip = new Unzip((file) => {
    if (failure) return;
    try {
      throwIfAborted(options.signal);
      const name = validateEntryName(file.name.normalize("NFC"), limits);
      if (seenNames.has(name)) throw invalidArchive(`压缩包包含重复文件路径：${name}。`);
      seenNames.add(name);
      entryCount += 1;
      assertLimit(entryCount, limits.maxArchiveEntries, "maxArchiveEntries", "archive");
      if (file.originalSize !== undefined) {
        assertLimit(file.originalSize, limits.maxSingleEntryBytes, "maxSingleEntryBytes", "archive");
      }
      if (file.size !== undefined && file.originalSize !== undefined) {
        assertLimit(compressionRatio(file.originalSize, file.size), limits.maxCompressionRatio, "maxCompressionRatio", "archive");
      }
      const xmlEntry = isXmlEntry(name);
      const xmlChunks: Uint8Array[] = [];
      let entryBytes = 0;
      file.ondata = (error, chunk, final) => {
        if (failure) return;
        try {
          if (error) throw error;
          throwIfAborted(options.signal);
          entryBytes += chunk.byteLength;
          totalUncompressed += chunk.byteLength;
          assertLimit(entryBytes, limits.maxSingleEntryBytes, "maxSingleEntryBytes", "archive");
          assertLimit(totalUncompressed, limits.maxUncompressedBytes, "maxUncompressedBytes", "archive");
          if (xmlEntry) {
            totalXmlBytes += chunk.byteLength;
            assertLimit(totalXmlBytes, limits.maxXmlBytes, "maxXmlBytes", "xml");
            xmlChunks.push(chunk);
          }
          if (final) {
            if (file.originalSize !== undefined && entryBytes !== file.originalSize) {
              throw invalidArchive(`压缩包文件 ${name} 的实际大小与目录记录不一致。`);
            }
            if (xmlEntry) {
              const xmlBytes = concatChunks(xmlChunks, entryBytes);
              const summary = validateOfficeXmlEntry(name, xmlBytes, limits);
              relationships += summary.relationships;
              assertLimit(relationships, limits.maxRelationships, "maxRelationships", "xml");
              worksheets += summary.worksheets;
              sharedStrings += summary.sharedStrings;
              formulas += summary.formulas;
              assertLimit(worksheets, limits.maxWorksheets, "maxWorksheets", "xml");
              assertLimit(sharedStrings, limits.maxSharedStrings, "maxSharedStrings", "xml");
              assertLimit(formulas, limits.maxFormulaCount, "maxFormulaCount", "xml");
              if (name === "[Content_Types].xml") contentTypes = decodeXml(xmlBytes);
            }
            completedEntries += 1;
          }
        } catch (error) {
          failure = error;
          try { file.terminate(); } catch { /* best effort */ }
        }
      };
      file.start();
    } catch (error) {
      failure = error;
      try { file.terminate(); } catch { /* best effort */ }
    }
  });
  unzip.register(UnzipInflate);
  try {
    unzip.push(bytes, true);
  } catch (error) {
    if (!failure) failure = invalidArchive("压缩包解压失败。", error);
  }
  if (failure) throw failure;
  if (entryCount !== central.entryCount || completedEntries !== central.entryCount) {
    throw invalidArchive("压缩包目录与实际文件数量不一致。");
  }
  for (const entry of central.entries) {
    if (!seenNames.has(entry.name)) {
      throw invalidArchive("压缩包目录与实际文件路径不一致。");
    }
  }
  const detected = contentTypes ? detectOoxmlSourceFormat(contentTypes) : undefined;
  if (!detected || detected.family !== options.format) {
    throw invalidArchive(`文件不是有效的 ${options.format.toUpperCase()} 文档。`);
  }
  return Object.freeze({
    ...central,
    sourceFormat: detected.format as OoxmlSourceFormat,
    uncompressedBytes: totalUncompressed,
    xmlBytes: totalXmlBytes,
    relationships,
    worksheets,
    sharedStrings,
    formulas,
  });
}
