import { OfficeLoadError } from "./errors";

export interface OfficeFetchHeaders {
  get(name: string): string | null;
}

export interface OfficeFetchBodyReader {
  read(): Promise<{ done: boolean; value?: Uint8Array }>;
  cancel(reason?: unknown): Promise<void> | void;
  releaseLock(): void;
}

export interface OfficeFetchBody {
  getReader(): OfficeFetchBodyReader;
  cancel?(reason?: unknown): Promise<void> | void;
}

export interface OfficeFetchResponse {
  readonly ok: boolean;
  readonly status: number;
  readonly url: string;
  readonly headers: OfficeFetchHeaders;
  /** Optional streamed body; when present the loader reads it under the byte cap. */
  readonly body?: OfficeFetchBody | null;
  arrayBuffer(): Promise<ArrayBuffer>;
}

export interface OfficeFetchInit {
  readonly credentials: "omit";
  /**
   * Redirects are followed by the fetcher. The loader revalidates the final
   * `response.url` against the caller's URL policy before accepting any bytes,
   * so a redirect can never smuggle in a disallowed origin.
   */
  readonly redirect: "follow";
  readonly signal?: AbortSignal;
}

export type OfficeFetch = (
  url: string,
  init: OfficeFetchInit,
) => Promise<OfficeFetchResponse>;

export interface OfficeUrlPolicy {
  enabled?: boolean;
  baseUrl?: string;
  allowRelativeUrl?: boolean;
  allowedProtocols?: readonly string[];
  allowedOrigins?: readonly string[];
  allowHttpOnLocalhost?: boolean;
  fetch?: OfficeFetch;
}

function fail(message: string, url?: string, cause?: unknown): never {
  throw new OfficeLoadError({
    code: "SOURCE_NOT_ALLOWED",
    message,
    sourceKind: "url",
    phase: "source",
    url,
    cause,
  });
}

function isLocalhost(hostname: string): boolean {
  const normalized = hostname.trim().toLowerCase().replace(/^\[/, "").replace(/\]$/, "");
  return normalized === "localhost" || normalized === "::1" || /^127(?:\.\d{1,3}){3}$/.test(normalized);
}

function hasUrlScheme(value: string): boolean {
  return /^[A-Za-z][A-Za-z\d+.-]*:/.test(value) || value.startsWith("//");
}

export function resolveOfficeUrl(rawUrl: string, policy: OfficeUrlPolicy | undefined): URL {
  if (!policy || policy.enabled === false) {
    fail("必须由宿主提供地址策略。");
  }
  if (rawUrl.trim() !== rawUrl || /[\u0000-\u001F\u007F]/.test(rawUrl)) {
    fail("地址包含不允许的空白或控制字符。");
  }
  if (!policy.baseUrl) {
    fail("地址策略缺少 baseUrl。");
  }
  if (!hasUrlScheme(rawUrl) && policy.allowRelativeUrl !== true) {
    fail("地址策略不允许相对地址。");
  }

  let resolved: URL;
  try {
    resolved = new URL(rawUrl, new URL(policy.baseUrl));
  } catch (error) {
    fail("地址或 baseUrl 无效。", undefined, error);
  }

  const protocols = new Set((policy.allowedProtocols ?? ["https:"]).map((value) => value.toLowerCase()));
  const localHttpAllowed =
    resolved.protocol === "http:" && policy.allowHttpOnLocalhost === true && isLocalhost(resolved.hostname);
  if (!protocols.has(resolved.protocol.toLowerCase()) && !localHttpAllowed) {
    fail("地址协议不在允许范围内。", resolved.href);
  }

  const origins = new Set<string>();
  for (const origin of policy.allowedOrigins ?? []) {
    try {
      origins.add(new URL(origin).origin);
    } catch (error) {
      fail("允许来源中包含无效地址。", undefined, error);
    }
  }
  if (!origins.has(resolved.origin)) {
    fail("地址来源不在允许范围内。", resolved.href);
  }
  return resolved;
}

export function snapshotOfficeUrlPolicy(
  policy: OfficeUrlPolicy | undefined,
): Readonly<OfficeUrlPolicy> | undefined {
  if (!policy) return undefined;
  return Object.freeze({
    ...policy,
    ...(policy.allowedProtocols ? { allowedProtocols: Object.freeze([...policy.allowedProtocols]) } : {}),
    ...(policy.allowedOrigins ? { allowedOrigins: Object.freeze([...policy.allowedOrigins]) } : {}),
  });
}
