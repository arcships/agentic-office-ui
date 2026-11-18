import {
  parseZip,
  parseZipLazyMedia,
  type PptxFiles,
} from "@aiden0z/pptx-renderer"
import {
  inspectOfficeArchive,
  validateOfficeArchiveParts,
  type OfficeArchivePart,
  type OfficeLimits,
} from "@arcships/office-runtime"

import type { PptxPreviewLimits } from "./types"

/**
 * Zip budgets forwarded to the archive reader. It mirrors the renderer's
 * `ZipParseLimits` so consumer projects never need the pinned renderer's types.
 */
export interface PptxArchiveZipLimits {
  maxEntries?: number
  maxEntryUncompressedBytes?: number
  maxTotalUncompressedBytes?: number
  maxMediaBytes?: number
  maxConcurrency?: number
}

/** Central-directory summary, read without decompressing a single entry. */
export interface PptxArchiveSummary {
  readonly entryCount: number
  readonly compressedBytes: number
  readonly uncompressedBytes: number
  readonly maxCompressionRatio: number
}

/** OOXML formats the PPTX family accepts. */
export type PptxArchiveSourceFormat = "pptx" | "pptm" | "ppsx" | "ppsm" | "potx" | "potm"

/** Budgets validated against the parts the parser really decoded. */
export interface PptxArchiveValidation {
  readonly sourceFormat: PptxArchiveSourceFormat
  readonly entryCount: number
  readonly uncompressedBytes: number
  readonly xmlBytes: number
  readonly relationships: number
}

/**
 * Reads the container bytes exactly once. Injectable so hosts can move parsing
 * into a worker and so tests can instrument or replace the unzip layer.
 */
export interface PptxArchiveReader {
  parse(buffer: ArrayBuffer, limits?: PptxArchiveZipLimits): Promise<unknown>
  parseLazyMedia(buffer: ArrayBuffer, limits?: PptxArchiveZipLimits): Promise<unknown>
}

export const PPTX_ARCHIVE_READER: PptxArchiveReader = Object.freeze({
  parse: (buffer: ArrayBuffer, limits?: PptxArchiveZipLimits) => parseZip(buffer, limits),
  parseLazyMedia: (buffer: ArrayBuffer, limits?: PptxArchiveZipLimits) =>
    parseZipLazyMedia(buffer, limits),
})

export interface LoadPptxArchiveOptions {
  /** Resolved preview budgets enforced on the decoded parts. */
  limits: PptxPreviewLimits
  /** Renderer-side unzip budgets, mapped from the same preview limits. */
  zipLimits?: PptxArchiveZipLimits
  /** Index media for on-demand decoding instead of inflating it during load. */
  lazyMedia?: boolean
  signal?: AbortSignal
  reader?: PptxArchiveReader
}

/**
 * Result of the single archive pass. `files` is the renderer's decoded container
 * (`PptxFiles`); it stays `unknown` in the public surface so consumer projects
 * never depend on the pinned renderer's declarations.
 */
export interface LoadedPptxArchive {
  readonly files: unknown
  readonly summary: PptxArchiveSummary
  readonly validation: PptxArchiveValidation
}

/** Map every documented PPTX budget onto the shared office limits. */
export function toOfficeLimits(limits: PptxPreviewLimits | undefined): OfficeLimits {
  return {
    maxInputBytes: limits?.maxInputBytes,
    maxArchiveEntries: limits?.maxArchiveEntries,
    maxSingleEntryBytes: limits?.maxSingleEntryBytes,
    maxUncompressedBytes: limits?.maxUncompressedBytes,
    maxTotalImageBytes: limits?.maxMediaBytes,
    maxCompressionRatio: limits?.maxCompressionRatio,
    maxArchivePathLength: limits?.maxArchivePathLength,
    maxXmlBytes: limits?.maxXmlBytes,
    maxSingleXmlBytes: limits?.maxSingleXmlBytes,
    maxXmlDepth: limits?.maxXmlDepth,
    maxXmlAttributeBytes: limits?.maxXmlAttributeBytes,
    maxTextNodeBytes: limits?.maxTextNodeBytes,
    maxRelationships: limits?.maxRelationships,
  }
}

/**
 * Flatten the parts a renderer parse produced. Only decoded parts are listed;
 * media a lazy resolver has not loaded yet contributes no size.
 */
function pptxArchiveParts(files: PptxFiles): readonly OfficeArchivePart[] {
  const parts: OfficeArchivePart[] = []
  const textPart = (name: string, text: string | undefined): void => {
    if (text !== undefined) parts.push({ name, text })
  }
  const mapParts = (entries: ReadonlyMap<string, string> | undefined): void => {
    if (!entries) return
    for (const [name, text] of entries) parts.push({ name, text })
  }
  textPart("[Content_Types].xml", files.contentTypes)
  textPart("ppt/presentation.xml", files.presentation)
  textPart("ppt/_rels/presentation.xml.rels", files.presentationRels)
  textPart("ppt/tableStyles.xml", files.tableStyles)
  mapParts(files.slides)
  mapParts(files.slideRels)
  mapParts(files.slideLayouts)
  mapParts(files.slideLayoutRels)
  mapParts(files.slideMasters)
  mapParts(files.slideMasterRels)
  mapParts(files.themes)
  mapParts(files.themeOverrides)
  mapParts(files.charts)
  mapParts(files.chartRels)
  mapParts(files.chartStyles)
  mapParts(files.chartColors)
  mapParts(files.diagramDrawings)
  for (const [name, data] of files.media) parts.push({ name, bytes: data.byteLength })
  return Object.freeze(parts)
}

/**
 * Inspect the central directory, decompress the container exactly once with the
 * reader, then validate the decoded parts. Inspecting the directory neither
 * inflates nor scans entries, so archive-level budgets still fire before any
 * decompression and no XML part is decoded twice.
 */
export async function loadPptxArchive(
  buffer: ArrayBuffer,
  options: LoadPptxArchiveOptions,
): Promise<LoadedPptxArchive> {
  const reader = options.reader ?? PPTX_ARCHIVE_READER
  const limits: Readonly<OfficeLimits> = toOfficeLimits(options.limits)
  const summary = inspectOfficeArchive(buffer, limits, options.signal)
  const files = options.lazyMedia
    ? await reader.parseLazyMedia(buffer, options.zipLimits)
    : await reader.parse(buffer, options.zipLimits)
  const validation = validateOfficeArchiveParts(pptxArchiveParts(files as PptxFiles), limits, {
    format: "pptx",
    signal: options.signal,
  })
  return Object.freeze({
    files,
    summary: Object.freeze({
      entryCount: summary.entryCount,
      compressedBytes: summary.compressedBytes,
      uncompressedBytes: summary.uncompressedBytes,
      maxCompressionRatio: summary.maxCompressionRatio,
    }),
    validation: Object.freeze({
      sourceFormat: validation.sourceFormat as PptxArchiveSourceFormat,
      entryCount: validation.entryCount,
      uncompressedBytes: validation.uncompressedBytes,
      xmlBytes: validation.xmlBytes,
      relationships: validation.relationships,
    }),
  })
}
