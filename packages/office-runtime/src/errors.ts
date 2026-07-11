export type OfficeFormat = "docx" | "xlsx" | "pdf";

export type OfficeSourceKind = "file" | "bytes" | "url";

export type OfficeLoadErrorCode =
  | "SOURCE_NOT_ALLOWED"
  | "FETCH_FAILED"
  | "INVALID_SOURCE"
  | "INVALID_ARGUMENT"
  | "LIMIT_EXCEEDED"
  | "IMAGE_LIMIT_EXCEEDED"
  | "INVALID_IMAGE"
  | "IMAGE_DECODE_FAILED"
  | "ABORTED"
  | "STALE_RESULT"
  | "RUNTIME_DISPOSED"
  | "WORKER_UNAVAILABLE"
  | "WORKER_FAILED"
  | "WASM_LOAD_FAILED"
  | "TIMEOUT";

export interface OfficeLoadErrorInit {
  code: OfficeLoadErrorCode;
  message: string;
  format?: OfficeFormat;
  sourceKind?: OfficeSourceKind;
  runtimeId?: string;
  taskId?: string;
  phase?: string;
  url?: string;
  limit?: string;
  actual?: number;
  allowed?: number;
  cause?: unknown;
}

export interface OfficeErrorFallback {
  fallbackCode: OfficeLoadErrorCode;
  message?: string;
  format?: OfficeFormat;
  sourceKind?: OfficeSourceKind;
  runtimeId?: string;
  taskId?: string;
  phase?: string;
  url?: string;
}

export function sanitizeOfficeUrl(rawUrl: string | undefined): string | undefined {
  if (!rawUrl) return undefined;
  try {
    const url = new URL(rawUrl);
    url.username = "";
    url.password = "";
    url.search = "";
    url.hash = "";
    return url.href;
  } catch {
    const withoutFragment = rawUrl.split("#", 1)[0] ?? "";
    const withoutQuery = withoutFragment.split("?", 1)[0] ?? "";
    return withoutQuery || undefined;
  }
}

export function isOfficeAbortError(error: unknown): boolean {
  return (
    (typeof DOMException !== "undefined" && error instanceof DOMException && error.name === "AbortError") ||
    (error instanceof Error && error.name === "AbortError") ||
    (typeof error === "object" && error !== null && "code" in error && error.code === "ABORTED")
  );
}

export class OfficeLoadError extends Error {
  readonly code: OfficeLoadErrorCode;
  readonly format?: OfficeFormat;
  readonly sourceKind?: OfficeSourceKind;
  readonly runtimeId?: string;
  readonly taskId?: string;
  readonly phase?: string;
  readonly url?: string;
  readonly limit?: string;
  readonly actual?: number;
  readonly allowed?: number;

  constructor(init: OfficeLoadErrorInit) {
    super(init.message);
    if (init.cause !== undefined) {
      (this as Error & { cause?: unknown }).cause = init.cause;
    }
    this.name = "OfficeLoadError";
    this.code = init.code;
    this.format = init.format;
    this.sourceKind = init.sourceKind;
    this.runtimeId = init.runtimeId;
    this.taskId = init.taskId;
    this.phase = init.phase;
    this.url = sanitizeOfficeUrl(init.url);
    this.limit = init.limit;
    this.actual = init.actual;
    this.allowed = init.allowed;
  }

  toJSON(): Record<string, unknown> {
    return {
      name: this.name,
      code: this.code,
      message: this.message,
      ...(this.format ? { format: this.format } : {}),
      ...(this.sourceKind ? { sourceKind: this.sourceKind } : {}),
      ...(this.runtimeId ? { runtimeId: this.runtimeId } : {}),
      ...(this.taskId ? { taskId: this.taskId } : {}),
      ...(this.phase ? { phase: this.phase } : {}),
      ...(this.url ? { url: this.url } : {}),
      ...(this.limit ? { limit: this.limit } : {}),
      ...(this.actual !== undefined ? { actual: this.actual } : {}),
      ...(this.allowed !== undefined ? { allowed: this.allowed } : {}),
    };
  }
}

export function toOfficeLoadError(error: unknown, fallback: OfficeErrorFallback): OfficeLoadError {
  if (error instanceof OfficeLoadError) return error;
  const aborted = isOfficeAbortError(error);
  return new OfficeLoadError({
    code: aborted ? "ABORTED" : fallback.fallbackCode,
    message: aborted ? "加载已取消。" : (fallback.message ?? "无法加载文档。"),
    format: fallback.format,
    sourceKind: fallback.sourceKind,
    runtimeId: fallback.runtimeId,
    taskId: fallback.taskId,
    phase: fallback.phase,
    url: fallback.url,
    cause: error,
  });
}
