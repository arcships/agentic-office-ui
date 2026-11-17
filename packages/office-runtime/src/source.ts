import {
  OfficeLoadError,
  isOfficeAbortError,
  sanitizeOfficeUrl,
  type OfficeSourceKind,
} from "./errors";
import {
  assertOfficeInputBytes,
  resolveOfficeInputByteLimit,
  type OfficeLimits,
} from "./limits";
import {
  resolveOfficeUrl,
  type OfficeFetch,
  type OfficeFetchResponse,
  type OfficeUrlPolicy,
} from "./url-policy";

export interface OfficeFileLike {
  readonly name?: string;
  readonly type?: string;
  readonly size?: number;
  arrayBuffer(): Promise<ArrayBuffer>;
}

export type OfficeSource =
  | { kind: "file"; file: OfficeFileLike; name?: string }
  /**
   * Byte sources are read in place: `loadOfficeSource` returns the caller's
   * buffer instead of copying it, so hosts must not reuse the buffer for a
   * differently sized archive while a load is running.
   */
  | { kind: "bytes"; bytes: ArrayBuffer; name?: string }
  | { kind: "url"; url: string; name?: string };

export interface ResolvedOfficeSource {
  buffer: ArrayBuffer;
  sourceKind: OfficeSourceKind;
  fileName?: string;
  resolvedUrl?: string;
  contentType?: string;
}

export interface LoadOfficeSourceOptions {
  signal?: AbortSignal;
  limits?: OfficeLimits;
  urlPolicy?: OfficeUrlPolicy;
}

function safeFileNameFromUrl(url: URL): string | undefined {
  const segment = url.pathname.split("/").filter(Boolean).pop();
  if (!segment) return undefined;
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

function failFetch(message: string, url: string, cause?: unknown): never {
  throw new OfficeLoadError({
    code: "FETCH_FAILED",
    message,
    sourceKind: "url",
    phase: "fetch",
    url,
    cause,
  });
}

function failRead(error: unknown, url: string): never {
  if (isOfficeAbortError(error)) {
    throw new OfficeLoadError({
      code: "ABORTED",
      message: "加载已取消。",
      sourceKind: "url",
      phase: "read",
      url,
    });
  }
  failFetch("无法读取文档响应。", url, error);
}

function inputLimitError(url: string, actual: number, allowed: number): OfficeLoadError {
  return new OfficeLoadError({
    code: "LIMIT_EXCEEDED",
    message: `输入大小 ${actual} 字节超过允许值 ${allowed} 字节。`,
    sourceKind: "url",
    phase: "input",
    url,
    limit: "maxInputBytes",
    actual,
    allowed,
  });
}

function parseContentLength(response: OfficeFetchResponse): number | undefined {
  const raw = response.headers?.get("content-length");
  if (!raw) return undefined;
  const parsed = Number(raw);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : undefined;
}

function resolveFetch(policy: OfficeUrlPolicy): OfficeFetch | undefined {
  if (policy.fetch) return policy.fetch;
  const platformFetch = globalThis.fetch;
  if (typeof platformFetch !== "function") return undefined;
  return async (url, init) => platformFetch(url, init) as unknown as OfficeFetchResponse;
}

async function cancelResponseBody(response: OfficeFetchResponse): Promise<void> {
  const body = response.body;
  if (!body || typeof body.cancel !== "function") return;
  try {
    await body.cancel();
  } catch {
    // The limit error above is the useful public failure.
  }
}

/**
 * Read a response body without ever allocating more than `allowed` bytes.
 *
 * The body is consumed chunk by chunk; as soon as the running total crosses
 * the cap the reader is cancelled and a LIMIT_EXCEEDED error is thrown, so a
 * chunked response without Content-Length cannot force a full allocation.
 * Custom fetchers that only expose `arrayBuffer()` keep working, but their
 * size can only be checked after the read.
 */
async function readResponseBody(
  response: OfficeFetchResponse,
  allowed: number | undefined,
  signal: AbortSignal | undefined,
  url: string,
): Promise<ArrayBuffer> {
  const body = response.body;
  if (!body || typeof body.getReader !== "function") {
    try {
      return await response.arrayBuffer();
    } catch (error) {
      failRead(error, url);
    }
  }

  const reader = body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  const cancelReader = async (): Promise<void> => {
    try {
      await reader.cancel();
    } catch {
      // The size or abort error below remains the useful public failure.
    }
  };
  try {
    for (;;) {
      if (signal?.aborted) {
        await cancelReader();
        throw new OfficeLoadError({
          code: "ABORTED",
          message: "加载已取消。",
          sourceKind: "url",
          phase: "read",
          url,
        });
      }
      const { done, value } = await reader.read();
      if (done) break;
      if (!value || value.byteLength === 0) continue;
      total += value.byteLength;
      if (allowed !== undefined && total > allowed) {
        await cancelReader();
        throw inputLimitError(url, total, allowed);
      }
      chunks.push(value);
    }
  } catch (error) {
    if (error instanceof OfficeLoadError) throw error;
    failRead(error, url);
  } finally {
    reader.releaseLock();
  }

  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes.buffer;
}

export async function loadOfficeSource(
  source: OfficeSource,
  options: LoadOfficeSourceOptions = {},
): Promise<ResolvedOfficeSource> {
  if (options.signal?.aborted) {
    throw new OfficeLoadError({ code: "ABORTED", message: "加载已取消。", sourceKind: source.kind });
  }
  // Validate the configured cap up front: a non-positive or non-finite value
  // is a configuration error, never an implicit "unlimited".
  const allowed = resolveOfficeInputByteLimit(options.limits);

  if (source.kind === "bytes") {
    assertOfficeInputBytes(source.bytes.byteLength, options.limits);
    // Byte sources are handed over, not copied: callers own their buffer and a
    // slice would double peak memory for large documents.
    return {
      buffer: source.bytes,
      sourceKind: "bytes",
      fileName: source.name,
    };
  }

  if (source.kind === "file") {
    if (source.file.size !== undefined) assertOfficeInputBytes(source.file.size, options.limits);
    let buffer: ArrayBuffer;
    try {
      buffer = await source.file.arrayBuffer();
    } catch (error) {
      if (isOfficeAbortError(error)) {
        throw new OfficeLoadError({ code: "ABORTED", message: "加载已取消。", sourceKind: "file" });
      }
      throw new OfficeLoadError({
        code: "INVALID_SOURCE",
        message: "无法读取文件。",
        sourceKind: "file",
        phase: "read",
        cause: error,
      });
    }
    assertOfficeInputBytes(buffer.byteLength, options.limits);
    return {
      buffer,
      sourceKind: "file",
      fileName: source.name ?? source.file.name,
      contentType: source.file.type,
    };
  }

  const policy = options.urlPolicy;
  const requestedUrl = resolveOfficeUrl(source.url, policy);
  const fetchImpl = policy ? resolveFetch(policy) : undefined;
  if (!fetchImpl) failFetch("当前环境没有可用的 fetch。", requestedUrl.href);

  let response: OfficeFetchResponse;
  try {
    response = await fetchImpl(requestedUrl.href, {
      credentials: "omit",
      // Follow redirects, then revalidate the final URL below before reading.
      redirect: "follow",
      signal: options.signal,
    });
  } catch (error) {
    if (isOfficeAbortError(error)) {
      throw new OfficeLoadError({
        code: "ABORTED",
        message: "加载已取消。",
        sourceKind: "url",
        phase: "fetch",
        url: requestedUrl.href,
      });
    }
    failFetch("无法获取文档。", requestedUrl.href, error);
  }
  if (!response.ok) failFetch(`文档地址返回失败状态 ${response.status}。`, requestedUrl.href);
  if (!response.url) failFetch("文档响应缺少最终地址。", requestedUrl.href);

  // A followed redirect is only trusted after the final address passes the
  // same protocol/origin policy as the originally requested address.
  const finalUrl = resolveOfficeUrl(response.url, policy);
  const contentLength = parseContentLength(response);
  if (contentLength !== undefined && allowed !== undefined && contentLength > allowed) {
    await cancelResponseBody(response);
    throw inputLimitError(finalUrl.href, contentLength, allowed);
  }

  const buffer = await readResponseBody(response, allowed, options.signal, finalUrl.href);
  assertOfficeInputBytes(buffer.byteLength, options.limits);
  return {
    buffer,
    sourceKind: "url",
    fileName: source.name ?? safeFileNameFromUrl(finalUrl),
    resolvedUrl: sanitizeOfficeUrl(finalUrl.href),
    contentType: response.headers?.get("content-type") ?? undefined,
  };
}
