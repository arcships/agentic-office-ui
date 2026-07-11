import type { PptxCapabilityReport, PptxFeatureRecord } from "./types"

export function createPptxCapabilityReport(
  features: readonly PptxFeatureRecord[] = [],
): PptxCapabilityReport {
  const records = Object.freeze([...features])
  return Object.freeze({
    discovered: records.length,
    strict: records.filter((item) => item.disposition === "strict").length,
    approximate: records.filter((item) => item.disposition === "approximate").length,
    static: records.filter((item) => item.disposition === "static").length,
    unparsed: records.filter((item) => item.disposition === "unparsed").length,
    features: records,
  })
}

