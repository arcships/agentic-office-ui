import { OfficeLoadError } from "./errors";

export interface OfficeLimits {
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
  maxDocxNodes?: number;
  maxDocxPages?: number;
  maxParseMs?: number;
}

export function snapshotOfficeLimits(limits: OfficeLimits | undefined): Readonly<OfficeLimits> {
  return Object.freeze({ ...(limits ?? {}) });
}

/**
 * Resolve one runtime instance's limits without retaining the caller's object.
 * A requested value can only tighten a hard cap, never disable or raise it.
 */
export function resolveOfficeLimits(
  defaults: Readonly<OfficeLimits>,
  requested: OfficeLimits | undefined,
  hardCaps: Readonly<OfficeLimits>,
): Readonly<OfficeLimits> {
  const resolved: OfficeLimits = {};
  for (const name of Object.keys(defaults) as Array<keyof OfficeLimits>) {
    const fallback = defaults[name];
    const hardCap = hardCaps[name];
    const candidate = requested?.[name];
    const positiveCandidate = typeof candidate === "number" && Number.isFinite(candidate) && candidate > 0
      ? candidate
      : fallback;
    resolved[name] = hardCap !== undefined && positiveCandidate !== undefined
      ? Math.min(positiveCandidate, hardCap)
      : positiveCandidate;
  }
  return Object.freeze(resolved);
}

export function assertOfficeInputBytes(actual: number, limits: OfficeLimits | undefined): void {
  const allowed = limits?.maxInputBytes;
  if (allowed === undefined || allowed <= 0 || actual <= allowed) return;
  throw new OfficeLoadError({
    code: "LIMIT_EXCEEDED",
    message: `输入大小 ${actual} 字节超过允许值 ${allowed} 字节。`,
    phase: "input",
    limit: "maxInputBytes",
    actual,
    allowed,
  });
}
