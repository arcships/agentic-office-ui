import {
  OfficeLoadError,
  loadOfficeSource,
  resolveOfficeUrl,
  sanitizeOfficeUrl,
  toOfficeLoadError,
  type OfficeUrlPolicy,
} from "@extend-ai/office-runtime";
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

function toOfficePolicy(policy: XlsxUrlPolicy | undefined): OfficeUrlPolicy | undefined {
  if (!policy) return undefined;
  return {
    ...policy,
    fetch: policy.fetch as OfficeUrlPolicy["fetch"],
  };
}

function mapOfficeCode(error: OfficeLoadError): XlsxLoadErrorCode {
  switch (error.code) {
    case "SOURCE_NOT_ALLOWED":
      return "SOURCE_NOT_ALLOWED";
    case "FETCH_FAILED":
      return "FETCH_FAILED";
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
    default:
      return "无法打开 XLSX 文件。";
  }
}

export class XlsxSourceError extends Error implements XlsxLoadError {
  readonly code: XlsxLoadErrorCode;
  readonly sourceKind: XlsxSourceKind;
  readonly url?: string;

  constructor(error: XlsxLoadError, options?: { cause?: unknown }) {
    super(error.message);
    if (options?.cause !== undefined) {
      (this as Error & { cause?: unknown }).cause = options.cause;
    }
    this.name = "XlsxSourceError";
    this.code = error.code;
    this.sourceKind = error.sourceKind;
    this.url = sanitizeOfficeUrl(error.url);
  }

  toJSON(): XlsxLoadError {
    return {
      code: this.code,
      message: this.message,
      sourceKind: this.sourceKind,
      ...(this.url ? { url: this.url } : {}),
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
    message: xlsxMessage(code),
    sourceKind,
    ...(officeError.url ? { url: officeError.url } : {}),
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
): Promise<ResolvedXlsxSource> {
  let resolved;
  try {
    resolved = await loadOfficeSource(
      { kind: "url", url: rawUrl },
      { urlPolicy: toOfficePolicy(policy), signal },
    );
  } catch (error) {
    throw fromOfficeError(error, "url", rawUrl);
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
