import type { OfficeLimits } from "@arcships/office-runtime/limits"

import type { PptxPreviewLimits } from "./types"

/**
 * Map every documented PPTX budget onto the shared office limits.
 *
 * Kept free of any archive/renderer import so parsers that only validate one
 * part (speaker notes, for example) can reuse the mapping without loading the
 * browser bundle.
 */
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
