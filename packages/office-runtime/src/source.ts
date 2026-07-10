import {
  OfficeLoadError,
  isOfficeAbortError,
  sanitizeOfficeUrl,
  type OfficeSourceKind,
} from "./errors";
import { assertOfficeInputBytes, type OfficeLimits } from "./limits";
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

export async function loadOfficeSource(
  source: OfficeSource,
  options: LoadOfficeSourceOptions = {},
): Promise<ResolvedOfficeSource> {
  if (options.signal?.aborted) {
    throw new OfficeLoadError({ code: "ABORTED", message: "加载已取消。", sourceKind: source.kind });
  }

  if (source.kind === "bytes") {
    assertOfficeInputBytes(source.bytes.byteLength, options.limits);
    return {
      buffer: source.bytes.slice(0),
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
      redirect: "error",
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

  const finalUrl = resolveOfficeUrl(response.url, policy);
  const contentLength = parseContentLength(response);
  if (contentLength !== undefined) assertOfficeInputBytes(contentLength, options.limits);

  let buffer: ArrayBuffer;
  try {
    buffer = await response.arrayBuffer();
  } catch (error) {
    if (isOfficeAbortError(error)) {
      throw new OfficeLoadError({
        code: "ABORTED",
        message: "加载已取消。",
        sourceKind: "url",
        phase: "read",
        url: finalUrl.href,
      });
    }
    failFetch("无法读取文档响应。", finalUrl.href, error);
  }
  assertOfficeInputBytes(buffer.byteLength, options.limits);
  return {
    buffer,
    sourceKind: "url",
    fileName: source.name ?? safeFileNameFromUrl(finalUrl),
    resolvedUrl: sanitizeOfficeUrl(finalUrl.href),
    contentType: response.headers?.get("content-type") ?? undefined,
  };
}
