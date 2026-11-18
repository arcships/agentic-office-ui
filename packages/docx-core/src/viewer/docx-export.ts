import type { DocModel } from "../engine/types";
import type { OoxmlPackage } from "../engine/ooxml-core";
import { serializeDocx } from "../engine/serializer";
import { DocxSerializationError } from "../engine/prepare-model";
import {
  resolveDocxRuntimeLimits,
  type DocxRuntimeLimits,
} from "../resource-limits";
import { isUiEnvironment } from "./runtime-environment";
import { isOpenDocumentModel, OPEN_DOCUMENT_VIEW_ONLY_REASON } from "./open-document";
import type { WorkerWasmSource } from "./wasm-source";

export type DocxExportErrorCode =
  | "ABORTED"
  | "TIMEOUT"
  | "WORKER_UNAVAILABLE"
  | "EXPORT_FAILED"
  /** The model is not serializable (an illegal collapsed table grid, for example). */
  | "INVALID_MODEL";

/** A stable, user-visible error shape for DOCX export failures. */
export class DocxExportError extends Error {
  readonly code: DocxExportErrorCode;
  readonly phase?: string;
  readonly limit?: string;
  readonly actual?: number;
  readonly allowed?: number;

  constructor(
    code: DocxExportErrorCode,
    message: string,
    details: {
      phase?: string;
      limit?: string;
      actual?: number;
      allowed?: number;
    } = {},
  ) {
    super(message);
    this.name = "DocxExportError";
    this.code = code;
    this.phase = details.phase;
    this.limit = details.limit;
    this.actual = details.actual;
    this.allowed = details.allowed;
  }
}

export type DocxExportProgressPhase =
  | "starting"
  | "serializing"
  | "finalizing"
  | "completed";

export interface DocxExportProgress {
  phase: DocxExportProgressPhase;
  percent: number;
}

export interface DocxExportOptions {
  signal?: AbortSignal;
  /** Per-call Worker factory; the runtime uses this for instance isolation. */
  workerFactory?: () => Worker;
  /** Diagnostic-only URL for a custom Worker factory. */
  workerUrl?: string;
  /** A worker-transferable WASM source shared by the caller and Worker. */
  wasmSource?: WorkerWasmSource;
  limits?: DocxRuntimeLimits;
  onProgress?: (progress: DocxExportProgress) => void;
}

export interface DocxExportWorkerRequest {
  id: string;
  type: "export-docx";
  model: DocModel;
  basePackage?: OoxmlPackage;
  wasmSource?: WorkerWasmSource;
}

export interface DocxExportWorkerProgressResponse {
  id: string;
  type: "progress";
  phase: Exclude<DocxExportProgressPhase, "completed">;
  percent: number;
}

export interface DocxExportWorkerSuccessResponse {
  id: string;
  type: "success";
  bytes: ArrayBuffer;
}

export interface DocxExportWorkerErrorResponse {
  id: string;
  type: "error";
  error: {
    code?: DocxExportErrorCode;
    name?: string;
    message: string;
    stack?: string;
  };
}

export type DocxExportWorkerResponse =
  | DocxExportWorkerProgressResponse
  | DocxExportWorkerSuccessResponse
  | DocxExportWorkerErrorResponse;

function createRequestId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `docx-export-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function createAbortError(): DocxExportError {
  const error = new DocxExportError("ABORTED", "DOCX export was aborted");
  error.name = "AbortError";
  return error;
}

function createTimeoutError(
  limits: Readonly<DocxRuntimeLimits>,
  actual: number,
): DocxExportError {
  const allowed = limits.maxExportMs ?? 0;
  return new DocxExportError(
    "TIMEOUT",
    `DOCX 导出耗时超过 ${allowed} 毫秒。`,
    {
      phase: "export",
      limit: "maxExportMs",
      actual: Math.max(actual, allowed + 1),
      allowed,
    },
  );
}

function createWorkerUnavailableError(cause?: unknown): DocxExportError {
  const detail = cause instanceof Error && cause.message
    ? `: ${cause.message}`
    : ".";
  return new DocxExportError(
    "WORKER_UNAVAILABLE",
    "DOCX export requires a Worker in UI environments" + detail,
  );
}

function createExportError(message: string, cause?: unknown): DocxExportError {
  const error = new DocxExportError("EXPORT_FAILED", message);
  if (cause instanceof Error && cause.stack) error.stack = cause.stack;
  return error;
}

/**
 * A model the serializer refused to write keeps its typed failure: callers can
 * branch on `code: "INVALID_MODEL"` and show the model error message instead of
 * a generic export failure.
 */
function createModelError(error: DocxSerializationError): DocxExportError {
  const exportError = new DocxExportError("INVALID_MODEL", error.message);
  if (error.stack) exportError.stack = error.stack;
  return exportError;
}

function performanceNow(): number {
  return typeof performance !== "undefined" &&
    typeof performance.now === "function"
    ? performance.now()
    : Date.now();
}

function reportProgress(
  options: DocxExportOptions,
  phase: DocxExportProgressPhase,
  percent: number,
): void {
  try {
    options.onProgress?.({ phase, percent });
  } catch {
    // Progress observers must not change an export result.
  }
}

/**
 * Create the DOCX export Worker bundled with this package. The static URL is
 * intentional: Vite can discover it for applications, and tsup emits the
 * matching file for direct package consumers.
 */
export function createBundledDocxExportWorker(): Worker {
  return new Worker(new URL("./docx-export-worker.js", import.meta.url), {
    type: "module",
    name: "@arcships/docx-core-export",
  });
}

function createWorkerFromUrl(url: string): Worker {
  return new Worker(url, {
    type: "module",
    name: "@arcships/docx-core-export",
  });
}

async function serializeInline(
  model: DocModel,
  basePackage: OoxmlPackage | undefined,
  limits: Readonly<DocxRuntimeLimits>,
  options: DocxExportOptions,
): Promise<ArrayBuffer> {
  const startedAt = performanceNow();
  reportProgress(options, "serializing", 10);
  let bytes: ArrayBuffer;
  try {
    bytes = await serializeDocx(model, basePackage);
  } catch (error) {
    if (error instanceof DocxSerializationError) throw createModelError(error);
    throw error;
  }
  if (options.signal?.aborted) {
    throw createAbortError();
  }
  const elapsed = performanceNow() - startedAt;
  if (limits.maxExportMs !== undefined && elapsed > limits.maxExportMs) {
    throw createTimeoutError(limits, elapsed);
  }
  reportProgress(options, "finalizing", 90);
  reportProgress(options, "completed", 100);
  return bytes;
}

/**
 * Serialize a DOCX model off the UI thread.
 *
 * In a UI environment this always uses a dedicated Worker, so the synchronous
 * WASM serializer, its JSON bridge and the ZIP post-patches never run on the
 * main thread. A headless host without `Worker` uses the engine serializer
 * inline because it has no UI to block.
 */
export async function exportDocxBuffer(
  model: DocModel,
  basePackage?: OoxmlPackage,
  options: DocxExportOptions = {},
): Promise<ArrayBuffer> {
  if (isOpenDocumentModel(model)) {
    throw new DocxExportError("INVALID_MODEL", OPEN_DOCUMENT_VIEW_ONLY_REASON);
  }
  const limits = resolveDocxRuntimeLimits(options.limits);
  if (options.signal?.aborted) {
    throw createAbortError();
  }

  if (typeof Worker === "undefined" && options.workerFactory === undefined) {
    if (isUiEnvironment()) {
      throw createWorkerUnavailableError();
    }
    return serializeInline(model, basePackage, limits, options);
  }

  let worker: Worker;
  try {
    worker = options.workerFactory
      ? options.workerFactory()
      : options.workerUrl
        ? createWorkerFromUrl(options.workerUrl)
        : createBundledDocxExportWorker();
  } catch (error) {
    throw createWorkerUnavailableError(error);
  }

  const requestId = createRequestId();
  return new Promise<ArrayBuffer>((resolve, reject) => {
    let settled = false;
    let timeout: ReturnType<typeof setTimeout> | undefined;

    const cleanup = (): void => {
      if (timeout !== undefined) clearTimeout(timeout);
      worker.removeEventListener("message", handleMessage);
      worker.removeEventListener("error", handleError);
      worker.removeEventListener("messageerror", handleMessageError);
      options.signal?.removeEventListener("abort", handleAbort);
      worker.terminate();
    };

    const settle = (resolver: () => void): void => {
      if (settled) return;
      settled = true;
      cleanup();
      resolver();
    };

    const handleAbort = (): void => {
      settle(() => reject(createAbortError()));
    };

    const fail = (error: DocxExportError): void => {
      settle(() => reject(error));
    };

    const handleError = (event: ErrorEvent): void => {
      fail(createExportError(event.message || "DOCX export Worker failed"));
    };

    const handleMessageError = (): void => {
      fail(createExportError("DOCX export Worker returned an unreadable response"));
    };

    const allowed = limits.maxExportMs;
    if (allowed !== undefined) {
      timeout = setTimeout(() => {
        fail(createTimeoutError(limits, allowed + 1));
      }, allowed);
    }

    const handleMessage = (event: MessageEvent<DocxExportWorkerResponse>): void => {
      const response = event.data;
      if (!response || response.id !== requestId) return;

      if (response.type === "progress") {
        reportProgress(options, response.phase, response.percent);
        return;
      }

      if (response.type === "error") {
        const error = new DocxExportError(
          response.error.code ?? "EXPORT_FAILED",
          response.error.message,
        );
        if (response.error.stack) error.stack = response.error.stack;
        fail(error);
        return;
      }

      reportProgress(options, "completed", 100);
      settle(() => resolve(response.bytes));
    };

    worker.addEventListener("message", handleMessage);
    worker.addEventListener("error", handleError);
    worker.addEventListener("messageerror", handleMessageError);
    options.signal?.addEventListener("abort", handleAbort, { once: true });

    reportProgress(options, "starting", 0);
    try {
      const request: DocxExportWorkerRequest = {
        id: requestId,
        type: "export-docx",
        model,
        basePackage,
        wasmSource: options.wasmSource,
      };
      // Structured clone keeps Uint8Array image bytes binary. There is no JSON
      // stringification and no per-byte number array crossing the boundary.
      worker.postMessage(request);
    } catch (error) {
      fail(createExportError(
        error instanceof Error ? error.message : "DOCX export Worker request failed",
        error,
      ));
    }
  });
}
