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

function isPositiveFinite(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value) && value > 0;
}

function invalidLimit(name: keyof OfficeLimits, value: unknown): never {
  throw new OfficeLoadError({
    code: "INVALID_ARGUMENT",
    message: `${String(name)} 必须是正的有限数值，收到 ${String(value)}。`,
    phase: "config",
    limit: name,
    actual: typeof value === "number" ? value : undefined,
  });
}

/**
 * Resolve one runtime instance's limits without retaining the caller's object.
 * A requested value can only tighten a hard cap, never disable or raise it.
 * Keys present in `hardCaps` but absent from `defaults` are still resolved to
 * the hard cap so a cap can never be dropped by omission. Non-positive or
 * non-finite requested values are rejected instead of silently disabling the
 * guard.
 */
export function resolveOfficeLimits(
  defaults: Readonly<OfficeLimits>,
  requested: OfficeLimits | undefined,
  hardCaps: Readonly<OfficeLimits>,
): Readonly<OfficeLimits> {
  const resolved: OfficeLimits = {};
  const names = new Set<keyof OfficeLimits>([
    ...(Object.keys(defaults) as Array<keyof OfficeLimits>),
    ...(Object.keys(hardCaps) as Array<keyof OfficeLimits>),
  ]);
  for (const name of names) {
    const hardCap = hardCaps[name];
    const fallback = defaults[name] ?? hardCap;
    const candidate = requested?.[name];
    let positiveCandidate: number | undefined;
    if (candidate === undefined) {
      positiveCandidate = fallback;
    } else if (isPositiveFinite(candidate)) {
      positiveCandidate = candidate;
    } else {
      invalidLimit(name, candidate);
    }
    resolved[name] = hardCap !== undefined && positiveCandidate !== undefined
      ? Math.min(positiveCandidate, hardCap)
      : positiveCandidate;
  }
  return Object.freeze(resolved);
}

/**
 * Read the configured input byte cap.
 *
 * `undefined` means the host did not configure a cap. Any other value must be
 * a positive finite number: `0`, negatives, `NaN` and `Infinity` are rejected
 * as configuration errors instead of being treated as "unlimited".
 */
export function resolveOfficeInputByteLimit(limits: OfficeLimits | undefined): number | undefined {
  const allowed = limits?.maxInputBytes;
  if (allowed === undefined) return undefined;
  if (!isPositiveFinite(allowed)) invalidLimit("maxInputBytes", allowed);
  return allowed;
}

export function assertOfficeInputBytes(actual: number, limits: OfficeLimits | undefined): void {
  const allowed = resolveOfficeInputByteLimit(limits);
  if (allowed === undefined || actual <= allowed) return;
  throw new OfficeLoadError({
    code: "LIMIT_EXCEEDED",
    message: `输入大小 ${actual} 字节超过允许值 ${allowed} 字节。`,
    phase: "input",
    limit: "maxInputBytes",
    actual,
    allowed,
  });
}
