/** Explicit XLSX browser/runtime adapter entry. */
export { bundledXlsxWasmUrl } from "./wasm-asset";
export { XlsxWorkerClient, createBundledXlsxWorker } from "./worker-client";
export type { XlsxWorkerClientOptions } from "./worker-client";
export { XlsxRuntimeError, createXlsxRuntime } from "./runtime/xlsx-runtime";
export type {
  XlsxRuntime,
  XlsxRuntimeConfig,
  XlsxRuntimeDiagnostic,
  XlsxRuntimeErrorCode,
  XlsxRuntimeParseOptions,
} from "./runtime/xlsx-runtime";
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
