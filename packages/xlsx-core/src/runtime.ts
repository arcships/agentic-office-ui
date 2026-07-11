/** Explicit XLSX browser/runtime adapter entry. */
export { bundledXlsxWasmUrl } from "./wasm-asset";
export { XlsxWorkerClient, XlsxWorkerError, createBundledXlsxWorker } from "./worker-client";
export type { XlsxWorkerClientOptions, XlsxWorkerErrorCode } from "./worker-client";
export { XlsxRuntimeError, createXlsxRuntime } from "./runtime/xlsx-runtime";
export {
  DEFAULT_XLSX_RUNTIME_LIMITS,
  resolveXlsxRuntimeLimits,
  validateXlsxArchive,
} from "./resource-limits";
export type {
  XlsxRuntime,
  XlsxRuntimeConfig,
  XlsxRuntimeDiagnostic,
  XlsxRuntimeErrorCode,
  XlsxRuntimeParseOptions,
} from "./runtime/xlsx-runtime";
export type {
  XlsxArchiveEntry,
  XlsxArchiveValidationResult,
  XlsxImageBudgetSnapshot,
  XlsxRuntimeLimits,
} from "./resource-limits";
export {
  XlsxSourceError,
  loadVerifiedXlsxSource,
  resolveAllowedXlsxUrl,
  toXlsxLoadError,
} from "./runtime/xlsx-url-policy";
export type { ResolvedXlsxSource } from "./runtime/xlsx-url-policy";
export type {
  XlsxDiagnostic,
  XlsxLoadError,
  XlsxLoadErrorCode,
  XlsxSourceKind,
  XlsxSourceState,
  XlsxUrlPolicy,
} from "./types/worksheet-types";
export type { WorkerWasmSource } from "./wasm";
