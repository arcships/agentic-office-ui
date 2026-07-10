import type { OfficeSourceKind } from "./errors";
import type { OfficeSource } from "./source";
import { snapshotOfficeLimits, type OfficeLimits } from "./limits";
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
    limits: snapshotOfficeLimits(options.limits),
    urlPolicy: snapshotOfficeUrlPolicy(options.urlPolicy),
    signal: options.signal,
  });
}
