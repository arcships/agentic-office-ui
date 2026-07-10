import { OfficeLoadError } from "./errors";

export interface OfficeLimits {
  maxInputBytes?: number;
}

export function snapshotOfficeLimits(limits: OfficeLimits | undefined): Readonly<OfficeLimits> {
  return Object.freeze({ ...(limits ?? {}) });
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
