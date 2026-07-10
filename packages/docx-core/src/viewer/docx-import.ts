import type { DocModel } from "../engine/types";
import type { OoxmlPackage } from "../engine/ooxml-core";
import { initWasm } from "../engine/wasm";
import {
  canUseConfiguredWasmSourceInWorker,
  getConfiguredWorkerWasmSource,
  type WorkerWasmSource,
} from "./wasm-source";
import { bundledDocxWasmUrl } from "./bundled-assets";

export type DocxImportSource = "worker" | "main-thread";

export type DocxImportErrorCode =
  | "ABORTED"
  | "STALE_RESULT"
  | "SOURCE_NOT_ALLOWED"
  | "FETCH_FAILED"
  | "LIMIT_EXCEEDED"
  | "WORKER_UNAVAILABLE"
  | "WASM_LOAD_FAILED"
  | "PARSE_FAILED"
  | "RUNTIME_DISPOSED";

/** A stable, user-visible error shape for DOCX loading failures. */
export class DocxImportError extends Error {
  readonly code: DocxImportErrorCode;

  constructor(code: DocxImportErrorCode, message: string) {
    super(message);
    this.name = "DocxImportError";
    this.code = code;
  }
}

export interface DocxImportResult {
  package: OoxmlPackage;
  model: DocModel;
  source: DocxImportSource;
  timings?: DocxImportWorkerTimings;
}

export type DocxImportDiagnosticType =
  | "load-start"
  | "worker-created"
  | "worker-success"
  | "worker-error"
  | "main-thread-start"
  | "main-thread-success"
  | "main-thread-fallback"
  | "aborted";

/**
 * This event deliberately contains execution metadata only. It never includes
 * document text, package contents, or an unredacted input URL.
 */
export interface DocxImportDiagnostic {
  type: DocxImportDiagnosticType;
  runtimeId?: string;
  requestId: string;
  source?: DocxImportSource;
  wasmUrl?: string;
  workerUrl?: string;
  code?: DocxImportErrorCode;
  message?: string;
  timings?: DocxImportWorkerTimings;
}

export interface DocxImportOptions {
  signal?: AbortSignal;
  transferBuffer?: boolean;
  /** Set to false only when the caller explicitly wants main-thread parsing. */
  useWorker?: boolean;
  /** Per-call worker factory; the runtime uses this for instance isolation. */
  workerFactory?: () => Worker;
  /** Diagnostic-only URL for a custom worker factory. */
  workerUrl?: string;
  /** A worker-transferable WASM source shared by the caller and Worker. */
  wasmSource?: WorkerWasmSource;
  /**
   * An opt-in, bounded fallback for environments where a Worker cannot be
   * created. Worker parse or WASM errors never use this fallback.
   */
  allowMainThreadFallback?: boolean;
  mainThreadFallbackMaxBytes?: number;
  onDiagnostic?: (event: DocxImportDiagnostic) => void;
  /** Supplied by an instance runtime; direct calls receive a unique id. */
  requestId?: string;
}

export interface DocxImportWorkerTimings {
  totalMs: number;
  parseMs: number;
  buildModelMs: number;
}

export interface DocxImportWorkerRequest {
  id: string;
  type: "import-docx";
  buffer: ArrayBuffer;
  wasmSource: WorkerWasmSource;
}

export interface DocxImportWorkerSuccessResponse {
  id: string;
  type: "success";
  package: OoxmlPackage;
  model: DocModel;
  timings: DocxImportWorkerTimings;
}

export interface DocxImportWorkerErrorResponse {
  id: string;
  type: "error";
  error: {
    code?: DocxImportErrorCode;
    name?: string;
    message: string;
    stack?: string;
  };
}

export type DocxImportWorkerResponse =
  | DocxImportWorkerSuccessResponse
  | DocxImportWorkerErrorResponse;

function createRequestId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `docx-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function createAbortError(): DocxImportError {
  const error = new DocxImportError("ABORTED", "DOCX import was aborted");
  error.name = "AbortError";
  return error;
}

export function classifyDocxImportErrorCode(error: unknown): DocxImportErrorCode {
  if (error instanceof DocxImportError) {
    return error.code;
  }

  if (error instanceof WebAssembly.CompileError) {
    return "WASM_LOAD_FAILED";
  }

  const message = error instanceof Error ? error.message : String(error);
  return /webassembly|wasm|instantiate|magic word/i.test(message)
    ? "WASM_LOAD_FAILED"
    : "PARSE_FAILED";
}

function toDocxImportError(
  error: unknown,
  fallbackCode: DocxImportErrorCode = "PARSE_FAILED",
): DocxImportError {
  if (error instanceof DocxImportError) {
    return error;
  }

  const message = error instanceof Error ? error.message : String(error);
  const wrapped = new DocxImportError(
    error instanceof Error ? classifyDocxImportErrorCode(error) : fallbackCode,
    message,
  );
  if (error instanceof Error && error.stack) {
    wrapped.stack = error.stack;
  }
  return wrapped;
}

function errorFromWorkerResponse(
  response: DocxImportWorkerErrorResponse,
): DocxImportError {
  const error = new DocxImportError(
    response.error.code ?? "PARSE_FAILED",
    response.error.message,
  );
  if (response.error.stack) {
    error.stack = response.error.stack;
  }
  return error;
}

function reportDiagnostic(options: DocxImportOptions, event: DocxImportDiagnostic): void {
  try {
    options.onDiagnostic?.(event);
  } catch {
    // Diagnostics must not change a document loading result.
  }
}

function isFallbackAllowed(buffer: ArrayBuffer, options: DocxImportOptions): boolean {
  const maxBytes = options.mainThreadFallbackMaxBytes ?? 0;
  return (
    options.allowMainThreadFallback === true &&
    Number.isFinite(maxBytes) &&
    maxBytes > 0 &&
    buffer.byteLength <= maxBytes
  );
}

function resolveWorkerWasmSource(options: DocxImportOptions): WorkerWasmSource {
  if (options.wasmSource !== undefined) {
    return options.wasmSource;
  }

  if (!canUseConfiguredWasmSourceInWorker()) {
    throw new DocxImportError(
      "WASM_LOAD_FAILED",
      "The configured DOCX WASM source cannot be transferred to the Worker.",
    );
  }

  return getConfiguredWorkerWasmSource() ?? bundledDocxWasmUrl;
}

/**
 * Create the Worker bundled with this package. The static URL is intentional:
 * Vite can discover it for applications, and tsup emits the matching file for
 * direct package consumers.
 */
export function createBundledDocxWorker(): Worker {
  return new Worker(new URL("./docx-import-worker.js", import.meta.url), {
    type: "module",
    name: "@extend-ai/docx-core-import",
  });
}

function createWorkerFromUrl(url: string): Worker {
  return new Worker(url, {
    type: "module",
    name: "@extend-ai/docx-core-import",
  });
}

async function importDocxOnMainThread(
  buffer: ArrayBuffer,
  signal: AbortSignal | undefined,
  wasmSource: WorkerWasmSource | undefined,
  options: DocxImportOptions,
  requestId: string,
): Promise<DocxImportResult> {
  if (signal?.aborted) {
    throw createAbortError();
  }

  reportDiagnostic(options, {
    type: "main-thread-start",
    requestId,
    source: "main-thread",
    wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
  });

  const startedAt = performanceNow();
  try {
    const [{ parseDocx }, { buildDocModel }] = await Promise.all([
      import("../engine/ooxml-core"),
      import("../engine/doc-model"),
    ]);
    if (wasmSource !== undefined) {
      await initWasm(wasmSource);
    }
    const pkg = await parseDocx(buffer);
    const parsedAt = performanceNow();
    if (signal?.aborted) {
      throw createAbortError();
    }

    const model = await buildDocModel(pkg);
    const finishedAt = performanceNow();
    if (signal?.aborted) {
      throw createAbortError();
    }

    const result: DocxImportResult = {
      package: pkg,
      model,
      source: "main-thread",
      timings: {
        totalMs: finishedAt - startedAt,
        parseMs: parsedAt - startedAt,
        buildModelMs: finishedAt - parsedAt,
      },
    };
    reportDiagnostic(options, {
      type: "main-thread-success",
      requestId,
      source: "main-thread",
      wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
      timings: result.timings,
    });
    return result;
  } catch (error) {
    const normalized = toDocxImportError(error);
    if (normalized.code === "ABORTED") {
      reportDiagnostic(options, {
        type: "aborted",
        requestId,
        source: "main-thread",
        code: normalized.code,
      });
    }
    throw normalized;
  }
}

function performanceNow(): number {
  return typeof performance !== "undefined" &&
    typeof performance.now === "function"
    ? performance.now()
    : Date.now();
}

function workerUnavailableError(error?: unknown): DocxImportError {
  const message = error instanceof Error && error.message
    ? `DOCX import Worker could not start: ${error.message}`
    : "DOCX import Worker is unavailable in this environment.";
  return new DocxImportError("WORKER_UNAVAILABLE", message);
}

async function handleWorkerUnavailable(
  buffer: ArrayBuffer,
  options: DocxImportOptions,
  requestId: string,
  wasmSource: WorkerWasmSource | undefined,
  error?: unknown,
): Promise<DocxImportResult> {
  const unavailable = workerUnavailableError(error);
  reportDiagnostic(options, {
    type: "worker-error",
    requestId,
    source: "worker",
    wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
    workerUrl: options.workerUrl,
    code: unavailable.code,
    message: unavailable.message,
  });

  if (!isFallbackAllowed(buffer, options)) {
    throw unavailable;
  }

  reportDiagnostic(options, {
    type: "main-thread-fallback",
    requestId,
    source: "main-thread",
    wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
    code: unavailable.code,
    message: unavailable.message,
  });
  return importDocxOnMainThread(buffer, options.signal, wasmSource, options, requestId);
}

export async function importDocxBuffer(
  buffer: ArrayBuffer,
  options: DocxImportOptions = {},
): Promise<DocxImportResult> {
  const requestId = options.requestId ?? createRequestId();
  if (options.signal?.aborted) {
    const error = createAbortError();
    reportDiagnostic(options, { type: "aborted", requestId, code: error.code });
    throw error;
  }

  if (options.useWorker === false) {
    reportDiagnostic(options, {
      type: "load-start",
      requestId,
      source: "main-thread",
    });
    const wasmSource = options.wasmSource ?? getConfiguredWorkerWasmSource();
    return importDocxOnMainThread(buffer, options.signal, wasmSource, options, requestId);
  }

  let wasmSource: WorkerWasmSource;
  try {
    wasmSource = resolveWorkerWasmSource(options);
  } catch (error) {
    const normalized = toDocxImportError(error, "WASM_LOAD_FAILED");
    reportDiagnostic(options, {
      type: "worker-error",
      requestId,
      source: "worker",
      code: normalized.code,
      message: normalized.message,
    });
    throw normalized;
  }

  reportDiagnostic(options, {
    type: "load-start",
    requestId,
    source: "worker",
    wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
    workerUrl: options.workerUrl,
  });

  if (typeof Worker === "undefined") {
    return handleWorkerUnavailable(buffer, options, requestId, wasmSource);
  }

  let worker: Worker;
  try {
    worker = options.workerFactory
      ? options.workerFactory()
      : options.workerUrl
        ? createWorkerFromUrl(options.workerUrl)
        : createBundledDocxWorker();
  } catch (error) {
    return handleWorkerUnavailable(buffer, options, requestId, wasmSource, error);
  }

  reportDiagnostic(options, {
    type: "worker-created",
    requestId,
    source: "worker",
    wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
    workerUrl: options.workerUrl,
  });

  return new Promise<DocxImportResult>((resolve, reject) => {
    let settled = false;

    const cleanup = (): void => {
      worker.removeEventListener("message", handleMessage);
      worker.removeEventListener("error", handleError);
      worker.removeEventListener("messageerror", handleMessageError);
      options.signal?.removeEventListener("abort", handleAbort);
      worker.terminate();
    };

    const settle = (resolver: () => void): void => {
      if (settled) {
        return;
      }
      settled = true;
      cleanup();
      resolver();
    };

    const handleAbort = (): void => {
      const error = createAbortError();
      reportDiagnostic(options, {
        type: "aborted",
        requestId,
        source: "worker",
        code: error.code,
      });
      settle(() => reject(error));
    };

    const rejectWorkerError = (error: DocxImportError): void => {
      reportDiagnostic(options, {
        type: "worker-error",
        requestId,
        source: "worker",
        wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
        workerUrl: options.workerUrl,
        code: error.code,
        message: error.message,
      });
      settle(() => reject(error));
    };

    const handleError = (event: ErrorEvent): void => {
      rejectWorkerError(
        new DocxImportError(
          "WORKER_UNAVAILABLE",
          event.message || "DOCX import Worker failed",
        ),
      );
    };

    const handleMessageError = (): void => {
      rejectWorkerError(
        new DocxImportError(
          "PARSE_FAILED",
          "DOCX import Worker returned an unreadable response",
        ),
      );
    };

    const handleMessage = (event: MessageEvent<DocxImportWorkerResponse>): void => {
      const response = event.data;
      if (!response || response.id !== requestId) {
        return;
      }

      if (response.type === "error") {
        rejectWorkerError(errorFromWorkerResponse(response));
        return;
      }

      reportDiagnostic(options, {
        type: "worker-success",
        requestId,
        source: "worker",
        wasmUrl: typeof wasmSource === "string" ? wasmSource : undefined,
        workerUrl: options.workerUrl,
        timings: response.timings,
      });
      settle(() =>
        resolve({
          package: response.package,
          model: response.model,
          source: "worker",
          timings: response.timings,
        }),
      );
    };

    worker.addEventListener("message", handleMessage);
    worker.addEventListener("error", handleError);
    worker.addEventListener("messageerror", handleMessageError);
    options.signal?.addEventListener("abort", handleAbort, { once: true });

    try {
      const request: DocxImportWorkerRequest = {
        id: requestId,
        type: "import-docx",
        buffer,
        wasmSource,
      };
      const transfer = options.transferBuffer ? [buffer] : [];
      worker.postMessage(request, transfer);
    } catch (error) {
      rejectWorkerError(
        toDocxImportError(error, "WORKER_UNAVAILABLE"),
      );
    }
  });
}
