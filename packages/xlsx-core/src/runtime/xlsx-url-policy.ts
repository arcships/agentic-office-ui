import {
  OfficeLoadError,
  loadOfficeSource,
  resolveOfficeUrl,
  sanitizeOfficeUrl,
  toOfficeLoadError,
  type OfficeFetch,
  type OfficeFetchResponse,
  type OfficeUrlPolicy,
} from "@arcships/office-runtime";
import type { XlsxRuntimeLimits } from "../resource-limits";
import type {
  XlsxLoadError,
  XlsxLoadErrorCode,
  XlsxSourceKind,
  XlsxUrlPolicy,
} from "../types";

export interface ResolvedXlsxSource {
  buffer: ArrayBuffer;
  fileName: string;
  resolvedUrl: string;
  sourceKind: "url";
}

type XlsxInputLimits = Pick<XlsxRuntimeLimits, "maxInputBytes">;

function inputSizeLimitError(actual: number, allowed: number): OfficeLoadError {
  return new OfficeLoadError({
    code: "LIMIT_EXCEEDED",
    message: `输入大小 ${actual} 字节超过允许值 ${allowed} 字节。`,
    format: "xlsx",
    sourceKind: "url",
    phase: "input",
    limit: "maxInputBytes",
    actual,
    allowed,
  });
}

async function cancelReader(reader: ReadableStreamDefaultReader<Uint8Array>): Promise<void> {
  try {
    await reader.cancel();
  } catch {
    // The size or abort error below remains the useful public failure.
  }
}

async function readBoundedResponse(
  response: Response,
  allowed: number,
  signal: AbortSignal | undefined,
): Promise<ArrayBuffer> {
  const reader = response.body?.getReader();
  if (!reader) return response.arrayBuffer();

  const chunks: Uint8Array[] = [];
  let actual = 0;
  try {
    while (true) {
      if (signal?.aborted) {
        await cancelReader(reader);
        throw new OfficeLoadError({
          code: "ABORTED",
          message: "XLSX 加载已取消。",
          format: "xlsx",
          sourceKind: "url",
          phase: "read",
        });
      }
      const { done, value } = await reader.read();
      if (done) break;
      if (!value || value.byteLength === 0) continue;
      actual += value.byteLength;
      if (actual > allowed) {
        await cancelReader(reader);
        throw inputSizeLimitError(actual, allowed);
      }
      chunks.push(value);
    }
  } finally {
    reader.releaseLock();
  }

  const bytes = new Uint8Array(actual);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes.buffer;
}

function boundedFetch(
  fetchImpl: typeof fetch,
  allowed: number,
): OfficeFetch {
  return async (url, init): Promise<OfficeFetchResponse> => {
    const response = await fetchImpl(url, init as RequestInit);
    const rawContentLength = response.headers.get("content-length");
    const contentLength = rawContentLength === null ? undefined : Number(rawContentLength);
    if (contentLength !== undefined && Number.isFinite(contentLength) && contentLength > allowed) {
      try {
        await response.body?.cancel();
      } catch {
        // Returning the stable limit error is more useful than a cancellation failure.
      }
      throw inputSizeLimitError(contentLength, allowed);
    }
    return {
      ok: response.ok,
      status: response.status,
      url: response.url,
      headers: response.headers,
      arrayBuffer: () => readBoundedResponse(response, allowed, init.signal),
    };
  };
}

function toOfficePolicy(
  policy: XlsxUrlPolicy | undefined,
  limits?: XlsxInputLimits,
): OfficeUrlPolicy | undefined {
  if (!policy) return undefined;
  const allowed = limits?.maxInputBytes;
  const fetchImpl = policy.fetch ?? (
    typeof globalThis.fetch === "function" ? globalThis.fetch.bind(globalThis) : undefined
  );
  return {
    ...policy,
    fetch: allowed !== undefined && allowed > 0 && fetchImpl
      ? boundedFetch(fetchImpl, allowed)
      : policy.fetch as OfficeUrlPolicy["fetch"],
  };
}

function nestedLimitError(error: unknown): OfficeLoadError | undefined {
  let candidate: unknown = error;
  const seen = new Set<unknown>();
  while (typeof candidate === "object" && candidate !== null && !seen.has(candidate)) {
    seen.add(candidate);
    if (candidate instanceof OfficeLoadError && candidate.code === "LIMIT_EXCEEDED") {
      return candidate;
    }
    candidate = "cause" in candidate ? candidate.cause : undefined;
  }
  return undefined;
}

function mapOfficeCode(error: OfficeLoadError): XlsxLoadErrorCode {
  switch (error.code) {
    case "SOURCE_NOT_ALLOWED":
      return "SOURCE_NOT_ALLOWED";
    case "FETCH_FAILED":
      return "FETCH_FAILED";
    case "LIMIT_EXCEEDED":
      return "LIMIT_EXCEEDED";
    case "IMAGE_LIMIT_EXCEEDED":
      return "IMAGE_LIMIT_EXCEEDED";
    case "INVALID_IMAGE":
      return "INVALID_IMAGE";
    case "IMAGE_DECODE_FAILED":
      return "IMAGE_DECODE_FAILED";
    case "TIMEOUT":
      return "TIMEOUT";
    case "ABORTED":
    case "STALE_RESULT":
      return "ABORTED";
    default:
      return "INVALID_WORKBOOK";
  }
}

function xlsxMessage(code: XlsxLoadErrorCode): string {
  switch (code) {
    case "SOURCE_NOT_ALLOWED":
      return "XLSX 地址不在允许范围内。";
    case "FETCH_FAILED":
      return "无法获取 XLSX 文件。";
    case "ABORTED":
      return "XLSX 加载已取消。";
    case "LIMIT_EXCEEDED":
      return "XLSX 文件超过资源限制。";
    case "IMAGE_LIMIT_EXCEEDED":
      return "XLSX 图片超过资源限制。";
    case "INVALID_IMAGE":
      return "XLSX 包含无效图片。";
    case "IMAGE_DECODE_FAILED":
      return "XLSX 图片无法解码。";
    case "TIMEOUT":
      return "XLSX 解析超时。";
    case "WORKER_UNAVAILABLE":
      return "XLSX Worker 无法完成解析。";
    default:
      return "无法打开 XLSX 文件。";
  }
}

export class XlsxSourceError extends Error implements XlsxLoadError {
  readonly code: XlsxLoadErrorCode;
  readonly sourceKind: XlsxSourceKind;
  readonly url?: string;
  readonly phase?: string;
  readonly limit?: string;
  readonly actual?: number;
  readonly allowed?: number;

  constructor(error: XlsxLoadError, options?: { cause?: unknown }) {
    super(error.message);
    if (options?.cause !== undefined) {
      (this as Error & { cause?: unknown }).cause = options.cause;
    }
    this.name = "XlsxSourceError";
    this.code = error.code;
    this.sourceKind = error.sourceKind;
    this.url = sanitizeOfficeUrl(error.url);
    this.phase = error.phase;
    this.limit = error.limit;
    this.actual = error.actual;
    this.allowed = error.allowed;
  }

  toJSON(): XlsxLoadError {
    return {
      code: this.code,
      message: this.message,
      sourceKind: this.sourceKind,
      ...(this.url ? { url: this.url } : {}),
      ...(this.phase ? { phase: this.phase } : {}),
      ...(this.limit ? { limit: this.limit } : {}),
      ...(this.actual !== undefined ? { actual: this.actual } : {}),
      ...(this.allowed !== undefined ? { allowed: this.allowed } : {}),
    };
  }
}

function fromOfficeError(error: unknown, sourceKind: XlsxSourceKind, url?: string): XlsxSourceError {
  const officeError = toOfficeLoadError(error, {
    fallbackCode: "INVALID_SOURCE",
    message: "无法打开 XLSX 文件。",
    format: "xlsx",
    sourceKind: sourceKind === "file" ? "bytes" : "url",
    url,
  });
  const code = mapOfficeCode(officeError);
  return new XlsxSourceError({
    code,
    message: code === "LIMIT_EXCEEDED"
      || code === "IMAGE_LIMIT_EXCEEDED"
      || code === "INVALID_IMAGE"
      || code === "IMAGE_DECODE_FAILED"
      || code === "TIMEOUT"
      ? officeError.message
      : xlsxMessage(code),
    sourceKind,
    ...(officeError.url ? { url: officeError.url } : {}),
    ...(officeError.phase ? { phase: officeError.phase } : {}),
    ...(officeError.limit ? { limit: officeError.limit } : {}),
    ...(officeError.actual !== undefined ? { actual: officeError.actual } : {}),
    ...(officeError.allowed !== undefined ? { allowed: officeError.allowed } : {}),
  }, { cause: officeError });
}

export function resolveAllowedXlsxUrl(rawUrl: string, policy: XlsxUrlPolicy | undefined): URL {
  try {
    return resolveOfficeUrl(rawUrl, toOfficePolicy(policy));
  } catch (error) {
    throw fromOfficeError(error, "url", rawUrl);
  }
}

function verifyWorkbookBytes(buffer: ArrayBuffer, url: string): void {
  const bytes = new Uint8Array(buffer);
  const hasZipHeader =
    bytes.length >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    ((bytes[2] === 0x03 && bytes[3] === 0x04) || (bytes[2] === 0x05 && bytes[3] === 0x06));
  const hasLegacyXlsHeader =
    bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1;
  if (!hasZipHeader && !hasLegacyXlsHeader) {
    throw new XlsxSourceError({
      code: "INVALID_WORKBOOK",
      message: "来源不是可验证的 XLSX 文件。",
      sourceKind: "url",
      url,
    });
  }
}

function fileNameFromUrl(rawUrl: string): string {
  const url = new URL(rawUrl);
  const segment = url.pathname.split("/").filter(Boolean).pop() || "workbook.xlsx";
  let decoded = segment;
  try {
    decoded = decodeURIComponent(segment);
  } catch {
    // Keep the escaped segment as a safe display name.
  }
  return /\.xlsx?$/i.test(decoded) ? decoded : `${decoded}.xlsx`;
}

export function toXlsxLoadError(error: unknown, sourceKind: XlsxSourceKind, url?: string): XlsxLoadError {
  if (error instanceof XlsxSourceError) return error.toJSON();
  if (typeof error === "object" && error !== null && "code" in error) {
    const candidate = error as {
      code?: unknown;
      message?: unknown;
      phase?: unknown;
      limit?: unknown;
      actual?: unknown;
      allowed?: unknown;
    };
    if (
      candidate.code === "LIMIT_EXCEEDED" ||
      candidate.code === "IMAGE_LIMIT_EXCEEDED" ||
      candidate.code === "INVALID_IMAGE" ||
      candidate.code === "IMAGE_DECODE_FAILED" ||
      candidate.code === "TIMEOUT" ||
      candidate.code === "WORKER_FAILED" ||
      candidate.code === "WORKER_UNAVAILABLE"
    ) {
      const code: XlsxLoadErrorCode = candidate.code === "WORKER_FAILED"
        ? "WORKER_UNAVAILABLE"
        : candidate.code;
      return {
        code,
        message: typeof candidate.message === "string" ? candidate.message : xlsxMessage(code),
        sourceKind,
        ...(url ? { url: sanitizeOfficeUrl(url) } : {}),
        ...(typeof candidate.phase === "string" ? { phase: candidate.phase } : {}),
        ...(typeof candidate.limit === "string" ? { limit: candidate.limit } : {}),
        ...(typeof candidate.actual === "number" ? { actual: candidate.actual } : {}),
        ...(typeof candidate.allowed === "number" ? { allowed: candidate.allowed } : {}),
      };
    }
  }
  return fromOfficeError(error, sourceKind, url).toJSON();
}

/**
 * Compatibility adapter over the private shared loading rules. Public XLSX
 * error names and codes remain stable while URL, fetch and redaction behavior
 * is owned by office-runtime.
 */
export async function loadVerifiedXlsxSource(
  rawUrl: string,
  policy: XlsxUrlPolicy | undefined,
  signal?: AbortSignal,
  limits?: XlsxInputLimits,
): Promise<ResolvedXlsxSource> {
  let resolved;
  try {
    resolved = await loadOfficeSource(
      { kind: "url", url: rawUrl },
      { urlPolicy: toOfficePolicy(policy, limits), signal, limits },
    );
  } catch (error) {
    throw fromOfficeError(nestedLimitError(error) ?? error, "url", rawUrl);
  }
  const resolvedUrl = resolved.resolvedUrl;
  if (!resolvedUrl) {
    throw fromOfficeError(new OfficeLoadError({
      code: "FETCH_FAILED",
      message: "XLSX 响应缺少最终地址。",
      sourceKind: "url",
    }), "url", rawUrl);
  }
  verifyWorkbookBytes(resolved.buffer, resolvedUrl);
  return {
    buffer: resolved.buffer,
    fileName: fileNameFromUrl(resolvedUrl),
    resolvedUrl,
    sourceKind: "url",
  };
}
