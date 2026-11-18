import type { OfficeSourceKind } from "./errors";
import type { OfficeSource } from "./source";
import { resolveOfficeLimits, snapshotOfficeLimits, type OfficeLimits } from "./limits";
import { snapshotOfficeUrlPolicy, type OfficeUrlPolicy } from "./url-policy";

export interface OfficeResourceConfig {
  wasmUrl?: string;
  workerUrl?: string;
}

export interface CreateOfficeLoadContextOptions {
  runtimeId: string;
  taskId: string;
  source: OfficeSource;
  resources?: OfficeResourceConfig;
  limits?: OfficeLimits;
  /** When present, `limits` is resolved against these format defaults. */
  limitDefaults?: OfficeLimits;
  /** Hard caps applied after the defaults when `limitDefaults` is present. */
  limitHardCaps?: OfficeLimits;
  urlPolicy?: OfficeUrlPolicy;
  signal: AbortSignal;
}

export interface OfficeLoadContext {
  readonly runtimeId: string;
  readonly taskId: string;
  readonly source: Readonly<OfficeSource>;
  readonly sourceKind: OfficeSourceKind;
  readonly resources: Readonly<OfficeResourceConfig>;
  readonly limits: Readonly<OfficeLimits>;
  readonly urlPolicy?: Readonly<OfficeUrlPolicy>;
  readonly signal: AbortSignal;
}

export function createOfficeLoadContext(
  options: CreateOfficeLoadContextOptions,
): Readonly<OfficeLoadContext> {
  const source = Object.freeze({ ...options.source }) as Readonly<OfficeSource>;
  return Object.freeze({
    runtimeId: options.runtimeId,
    taskId: options.taskId,
    source,
    sourceKind: source.kind,
    resources: Object.freeze({ ...(options.resources ?? {}) }),
    limits: options.limitDefaults
      ? resolveOfficeLimits(options.limitDefaults, options.limits, options.limitHardCaps ?? {})
      : snapshotOfficeLimits(options.limits),
    urlPolicy: snapshotOfficeUrlPolicy(options.urlPolicy),
    signal: options.signal,
  });
}
