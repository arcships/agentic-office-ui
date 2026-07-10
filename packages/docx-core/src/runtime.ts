import {
  OfficeLoadError,
  assertOfficeInputBytes,
  createLatestTaskCoordinator,
  createOfficeTaskSequence,
  loadOfficeSource,
  snapshotOfficeLimits,
  snapshotOfficeUrlPolicy,
  type LatestTaskCoordinator,
  type OfficeFetch,
  type OfficeLimits,
  type OfficeSource,
  type OfficeUrlPolicy,
} from "@extend-ai/office-runtime";
import {
  DocxImportError,
  createBundledDocxWorker,
  importDocxBuffer,
  type DocxImportDiagnostic,
  type DocxImportErrorCode,
  type DocxImportOptions,
  type DocxImportResult,
} from "./viewer/docx-import";
import { bundledDocxWasmUrl } from "./viewer/bundled-assets";
import type { WorkerWasmSource } from "./viewer/wasm-source";

export { bundledDocxWasmUrl } from "./viewer/bundled-assets";
export { DocxImportError } from "./viewer/docx-import";
export type {
  DocxImportDiagnostic,
  DocxImportDiagnosticType,
  DocxImportErrorCode,
  DocxImportOptions,
  DocxImportResult,
  DocxImportSource,
  DocxImportWorkerTimings,
} from "./viewer/docx-import";
export type { WorkerWasmSource } from "./viewer/wasm-source";

export interface DocxFileLike {
  readonly name?: string;
  readonly type?: string;
  readonly size?: number;
  arrayBuffer(): Promise<ArrayBuffer>;
}

export type DocxSource =
  | { kind: "file"; file: DocxFileLike; name?: string }
  | { kind: "bytes"; bytes: ArrayBuffer; name?: string }
  | { kind: "url"; url: string; name?: string };

export interface DocxUrlPolicy {
  enabled?: boolean;
  baseUrl?: string;
  allowRelativeUrl?: boolean;
  allowedProtocols?: readonly string[];
  allowedOrigins?: readonly string[];
  allowHttpOnLocalhost?: boolean;
  fetch?: typeof fetch;
}

export interface DocxRuntimeLimits {
  maxInputBytes?: number;
}

export interface DocxRuntimeConfig {
  /** A URL or transferable binary used by both the main runtime and Worker. */
  wasmSource?: WorkerWasmSource;
  /** Convenience alias for the common URL case. */
  wasmUrl?: string;
  /** Use this public URL when the host serves its own Worker. */
  workerUrl?: string;
  /** Override Worker creation without relying on module-level configuration. */
  createWorker?: () => Worker;
  /** Disabled by default; fallback also requires a positive byte limit. */
  allowMainThreadFallback?: boolean;
  mainThreadFallbackMaxBytes?: number;
  /** Per-instance source policy; URL loading is disabled when omitted. */
  urlPolicy?: DocxUrlPolicy;
  /** Current shared input limit. Archive/model budgets are added by P3. */
  limits?: DocxRuntimeLimits;
  onDiagnostic?: (event: DocxImportDiagnostic) => void;
}

export interface DocxRuntimeLoader {
  load(
    source: ArrayBuffer | DocxSource,
    options?: Omit<DocxImportOptions, "requestId">,
  ): Promise<DocxImportResult>;
  cancel(): void;
  dispose(): void;
}

export interface DocxRuntime {
  readonly id: string;
  importDocxBuffer(
    buffer: ArrayBuffer,
    options?: Omit<DocxImportOptions, "requestId">,
  ): Promise<DocxImportResult>;
  loadSource(
    source: DocxSource,
    options?: Omit<DocxImportOptions, "requestId">,
  ): Promise<DocxImportResult>;
  createLoader(): DocxRuntimeLoader;
  dispose(): void;
}

function createRuntimeId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return `docx-runtime-${crypto.randomUUID()}`;
  }
  return `docx-runtime-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function asOfficeSource(source: ArrayBuffer | DocxSource): OfficeSource {
  if (source instanceof ArrayBuffer) return { kind: "bytes", bytes: source };
  return source as OfficeSource;
}

function asOfficePolicy(policy: DocxUrlPolicy | undefined): OfficeUrlPolicy | undefined {
  if (!policy) return undefined;
  return { ...policy, fetch: policy.fetch as OfficeFetch | undefined };
}

function mapOfficeError(error: OfficeLoadError): DocxImportError {
  const compatibleCodes = new Set<DocxImportErrorCode>([
    "ABORTED",
    "STALE_RESULT",
    "SOURCE_NOT_ALLOWED",
    "FETCH_FAILED",
    "LIMIT_EXCEEDED",
    "WORKER_UNAVAILABLE",
    "WASM_LOAD_FAILED",
    "PARSE_FAILED",
    "RUNTIME_DISPOSED",
  ]);
  const code = error.code === "STALE_RESULT"
    ? "ABORTED"
    : compatibleCodes.has(error.code as DocxImportErrorCode)
    ? error.code as DocxImportErrorCode
    : "PARSE_FAILED";
  const mapped = new DocxImportError(code, error.message);
  if (code === "ABORTED") mapped.name = "AbortError";
  (mapped as Error & { cause?: unknown }).cause = error;
  return mapped;
}

/**
 * Creates an isolated DOCX execution configuration. Shared office-runtime
 * utilities provide source policy, input limits and task ownership, while the
 * format runtime continues to own Worker/WASM and parsing state.
 */
export function createDocxRuntime(config: DocxRuntimeConfig = {}): DocxRuntime {
  const id = createRuntimeId();
  const taskSequence = createOfficeTaskSequence(id);
  const runtimeConfig = Object.freeze({
    ...config,
    limits: snapshotOfficeLimits(config.limits),
    urlPolicy: snapshotOfficeUrlPolicy(asOfficePolicy(config.urlPolicy)),
  });
  const wasmSource = runtimeConfig.wasmSource ?? runtimeConfig.wasmUrl ?? bundledDocxWasmUrl;
  const createWorker = runtimeConfig.createWorker ?? (
    runtimeConfig.workerUrl
      ? () => new Worker(runtimeConfig.workerUrl as string, {
        type: "module",
        name: "@extend-ai/docx-core-import",
      })
      : createBundledDocxWorker
  );
  const coordinators = new Set<LatestTaskCoordinator>();
  let disposed = false;

  const performImport = async (
    buffer: ArrayBuffer,
    taskId: string,
    signal: AbortSignal,
    options: Omit<DocxImportOptions, "requestId">,
  ): Promise<DocxImportResult> => {
    const notify = (event: DocxImportDiagnostic): void => {
      const runtimeEvent = { ...event, runtimeId: id };
      try { runtimeConfig.onDiagnostic?.(runtimeEvent); } catch { /* diagnostics are isolated */ }
      try { options.onDiagnostic?.(runtimeEvent); } catch { /* diagnostics are isolated */ }
    };
    return importDocxBuffer(buffer, {
      ...options,
      signal,
      requestId: taskId,
      wasmSource: options.wasmSource ?? wasmSource,
      workerFactory: options.workerFactory ?? createWorker,
      workerUrl: options.workerUrl ?? runtimeConfig.workerUrl,
      allowMainThreadFallback:
        options.allowMainThreadFallback ?? runtimeConfig.allowMainThreadFallback,
      mainThreadFallbackMaxBytes:
        options.mainThreadFallbackMaxBytes ?? runtimeConfig.mainThreadFallbackMaxBytes,
      onDiagnostic: notify,
    });
  };

  const runWithCoordinator = async (
    coordinator: LatestTaskCoordinator,
    source: ArrayBuffer | DocxSource,
    options: Omit<DocxImportOptions, "requestId"> = {},
  ): Promise<DocxImportResult> => {
    if (disposed) {
      throw new DocxImportError("RUNTIME_DISPOSED", "This DOCX runtime has been disposed.");
    }
    const officeSource = asOfficeSource(source);
    const task = coordinator.start(officeSource, {
      signal: options.signal,
      limits: runtimeConfig.limits as OfficeLimits,
      resources: {
        wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
        workerUrl: options.workerUrl ?? runtimeConfig.workerUrl,
      },
      urlPolicy: runtimeConfig.urlPolicy,
    });
    try {
      const buffer = officeSource.kind === "bytes"
        ? (() => {
            assertOfficeInputBytes(officeSource.bytes.byteLength, runtimeConfig.limits);
            return officeSource.bytes;
          })()
        : (await loadOfficeSource(officeSource, {
            signal: task.signal,
            limits: runtimeConfig.limits as OfficeLimits,
            urlPolicy: runtimeConfig.urlPolicy,
          })).buffer;
      task.assertCurrent();
      const result = await performImport(buffer, task.context.taskId, task.signal, options);
      task.assertCurrent();
      return result;
    } catch (error) {
      if (error instanceof OfficeLoadError) throw mapOfficeError(error);
      throw error;
    } finally {
      task.finish();
    }
  };

  const createTrackedCoordinator = (): LatestTaskCoordinator => {
    const coordinator = createLatestTaskCoordinator(taskSequence);
    coordinators.add(coordinator);
    return coordinator;
  };

  const runOne = async (
    source: ArrayBuffer | DocxSource,
    options: Omit<DocxImportOptions, "requestId"> = {},
  ): Promise<DocxImportResult> => {
    const coordinator = createTrackedCoordinator();
    try {
      return await runWithCoordinator(coordinator, source, options);
    } finally {
      coordinator.dispose();
      coordinators.delete(coordinator);
    }
  };

  return {
    id,
    importDocxBuffer: (buffer, options) => runOne(buffer, options),
    loadSource: (source, options) => runOne(source, options),
    createLoader(): DocxRuntimeLoader {
      const coordinator = createTrackedCoordinator();
      let loaderDisposed = false;
      return {
        load(source, options = {}): Promise<DocxImportResult> {
          if (loaderDisposed) {
            throw new DocxImportError("RUNTIME_DISPOSED", "This DOCX loader has been disposed.");
          }
          return runWithCoordinator(coordinator, source, options);
        },
        cancel(): void {
          coordinator.cancel();
        },
        dispose(): void {
          if (loaderDisposed) return;
          loaderDisposed = true;
          coordinator.dispose();
          coordinators.delete(coordinator);
        },
      };
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      for (const coordinator of coordinators) coordinator.dispose();
      coordinators.clear();
      taskSequence.dispose();
    },
  };
}
