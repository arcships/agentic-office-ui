import {
  loadOfficeSource,
  sanitizeOfficeUrl,
  toOfficeLoadError,
  type OfficeLoadError,
  type OfficeUrlPolicy,
} from "@arcships/office-runtime"

export type PdfSource =
  | { kind: "url"; url: string }
  | { kind: "blob"; blob: Blob; fileName?: string }

export interface PdfUrlPolicy {
  enabled?: boolean
  baseUrl?: string
  allowRelativeUrl?: boolean
  allowedProtocols?: readonly string[]
  allowedOrigins?: readonly string[]
  allowHttpOnLocalhost?: boolean
  fetch?: typeof fetch
}

export const DEFAULT_PDF_MAX_FILE_SIZE = 50 * 1024 * 1024

export interface PdfLoadOptions {
  signal?: AbortSignal
  /** Maximum accepted source size in bytes. Defaults to 50 MiB. */
  maxFileSize?: number
}

export type PdfLoadErrorCode =
  | "SOURCE_NOT_ALLOWED"
  | "FETCH_FAILED"
  | "INVALID_PDF"
  | "PDF_TOO_LARGE"
  | "ABORTED"

export interface PdfLoadError {
  code: PdfLoadErrorCode
  message: string
  sourceKind: PdfSource["kind"]
  url?: string
  actual?: number
  allowed?: number
}

export interface PdfVerifiedDocument {
  bytes: Uint8Array
  blob: Blob
  fileName: string
  sourceKind: PdfSource["kind"]
  resolvedUrl?: string
}

export interface PdfDiagnostic {
  type:
    | "load-start"
    | "load-success"
    | "load-error"
    | "load-cancelled"
    | "render-start"
    | "render-success"
    | "render-error"
    | "search-start"
    | "search-success"
    | "search-error"
    | "object-url-revoked"
    | "download"
  requestId: number
  taskId?: string
  runtimeId?: string
  sourceKind?: PdfSource["kind"]
  url?: string
  bytes?: number
  pageNumber?: number
  pageCount?: number
  zoom?: number
  rotation?: number
  matches?: number
  error?: PdfLoadError
}

export class PdfSourceError extends Error implements PdfLoadError {
  readonly code: PdfLoadErrorCode
  readonly sourceKind: PdfSource["kind"]
  readonly url?: string
  readonly actual?: number
  readonly allowed?: number

  constructor(error: PdfLoadError, options?: { cause?: unknown }) {
    super(error.message)
    if (options?.cause !== undefined) {
      (this as Error & { cause?: unknown }).cause = options.cause
    }
    this.name = "PdfSourceError"
    this.code = error.code
    this.sourceKind = error.sourceKind
    this.url = sanitizeOfficeUrl(error.url)
    this.actual = error.actual
    this.allowed = error.allowed
  }

  toJSON(): PdfLoadError {
    return {
      code: this.code,
      message: this.message,
      sourceKind: this.sourceKind,
      ...(this.url ? { url: this.url } : {}),
      ...(this.actual !== undefined ? { actual: this.actual } : {}),
      ...(this.allowed !== undefined ? { allowed: this.allowed } : {}),
    }
  }
}

function toOfficePolicy(policy: PdfUrlPolicy | undefined): OfficeUrlPolicy | undefined {
  if (!policy) return undefined
  return { ...policy, fetch: policy.fetch as OfficeUrlPolicy["fetch"] }
}

function mapOfficeCode(error: OfficeLoadError): PdfLoadErrorCode {
  switch (error.code) {
    case "SOURCE_NOT_ALLOWED": return "SOURCE_NOT_ALLOWED"
    case "FETCH_FAILED": return "FETCH_FAILED"
    case "LIMIT_EXCEEDED": return "PDF_TOO_LARGE"
    case "ABORTED":
    case "STALE_RESULT": return "ABORTED"
    default: return "INVALID_PDF"
  }
}

function pdfMessage(code: PdfLoadErrorCode, actual?: number, allowed?: number): string {
  switch (code) {
    case "SOURCE_NOT_ALLOWED": return "PDF 地址不在允许范围内。"
    case "FETCH_FAILED": return "无法获取 PDF 文件。"
    case "PDF_TOO_LARGE": return `PDF 文件大小 ${actual ?? "未知"} 字节超过允许值 ${allowed ?? "未知"} 字节。`
    case "ABORTED": return "PDF 加载已取消。"
    default: return "来源不是可验证的 PDF 文件。"
  }
}

function fromOfficeError(error: unknown, sourceKind: PdfSource["kind"], url?: string): PdfSourceError {
  const officeError = toOfficeLoadError(error, {
    fallbackCode: "INVALID_SOURCE",
    message: "无法打开 PDF 文件。",
    format: "pdf",
    sourceKind: sourceKind === "blob" ? "file" : "url",
    url,
  })
  const code = mapOfficeCode(officeError)
  return new PdfSourceError({
    code,
    message: pdfMessage(code, officeError.actual, officeError.allowed),
    sourceKind,
    ...(officeError.url ? { url: officeError.url } : {}),
    ...(officeError.actual !== undefined ? { actual: officeError.actual } : {}),
    ...(officeError.allowed !== undefined ? { allowed: officeError.allowed } : {}),
  }, { cause: officeError })
}

function normalizeMaxFileSize(value: number | undefined): number {
  if (value === undefined) return DEFAULT_PDF_MAX_FILE_SIZE
  if (!Number.isFinite(value)) return DEFAULT_PDF_MAX_FILE_SIZE
  return Math.max(1, Math.floor(value))
}

function normalizeLoadOptions(value: AbortSignal | PdfLoadOptions | undefined): Required<Pick<PdfLoadOptions, "maxFileSize">> & Pick<PdfLoadOptions, "signal"> {
  const candidate = value as AbortSignal | undefined
  const isSignal = !!candidate && typeof candidate.aborted === "boolean" && typeof candidate.addEventListener === "function"
  const options = isSignal ? { signal: value as AbortSignal } : (value as PdfLoadOptions | undefined)
  return {
    signal: options?.signal,
    maxFileSize: normalizeMaxFileSize(options?.maxFileSize),
  }
}

function verifyPdfBytes(bytes: Uint8Array, sourceKind: PdfSource["kind"], url?: string): void {
  const isPdf =
    bytes.length >= 5 &&
    bytes[0] === 0x25 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x44 &&
    bytes[3] === 0x46 &&
    bytes[4] === 0x2d
  if (!isPdf) {
    throw new PdfSourceError({ code: "INVALID_PDF", message: "来源不是可验证的 PDF 文件。", sourceKind, url })
  }
}

function verifyPdfMime(contentType: string | undefined, sourceKind: PdfSource["kind"], url?: string): void {
  const mime = contentType?.split(";", 1)[0]?.trim().toLowerCase()
  if (mime !== "application/pdf") {
    throw new PdfSourceError({ code: "INVALID_PDF", message: "来源响应不是 application/pdf。", sourceKind, url })
  }
}

function fileNameFromUrl(rawUrl: string): string {
  const url = new URL(rawUrl)
  const lastSegment = url.pathname.split("/").filter(Boolean).pop()
  let decoded = lastSegment || "document.pdf"
  try { decoded = decodeURIComponent(decoded) } catch { /* Keep safe escaped text. */ }
  return decoded.toLowerCase().endsWith(".pdf") ? decoded : `${decoded}.pdf`
}

function fileNameFromBlob(source: Extract<PdfSource, { kind: "blob" }>): string {
  if (source.fileName?.trim()) {
    return source.fileName.toLowerCase().endsWith(".pdf") ? source.fileName : `${source.fileName}.pdf`
  }
  const file = source.blob as File
  if (typeof file.name === "string" && file.name.trim()) {
    return file.name.toLowerCase().endsWith(".pdf") ? file.name : `${file.name}.pdf`
  }
  return "document.pdf"
}

export function toPdfLoadError(error: unknown, sourceKind: PdfSource["kind"], url?: string): PdfLoadError {
  if (error instanceof PdfSourceError) return error.toJSON()
  return fromOfficeError(error, sourceKind, url).toJSON()
}

/**
 * Compatibility adapter over office-runtime. PDF keeps its public Blob shape,
 * while URL validation, controlled fetch, cancellation and redaction are shared.
 */
export async function loadVerifiedPdfSource(
  source: PdfSource,
  policy: PdfUrlPolicy | undefined,
  optionsOrSignal?: AbortSignal | PdfLoadOptions,
): Promise<PdfVerifiedDocument> {
  const options = normalizeLoadOptions(optionsOrSignal)
  try {
    const resolved = source.kind === "blob"
      ? await loadOfficeSource({
          kind: "file",
          name: fileNameFromBlob(source),
          file: {
            name: fileNameFromBlob(source),
            type: source.blob.type,
            size: source.blob.size,
            arrayBuffer: () => source.blob.arrayBuffer(),
          },
        }, { signal: options.signal, limits: { maxInputBytes: options.maxFileSize } })
      : await loadOfficeSource(
          { kind: "url", url: source.url },
          {
            signal: options.signal,
            limits: { maxInputBytes: options.maxFileSize },
            urlPolicy: toOfficePolicy(policy),
          },
        )

    const sourceKind = source.kind
    if (sourceKind === "blob") {
      const mime = source.blob.type.trim().toLowerCase()
      if (mime && mime !== "application/pdf") {
        throw new PdfSourceError({ code: "INVALID_PDF", message: "文件不是 application/pdf。", sourceKind })
      }
    } else {
      verifyPdfMime(resolved.contentType, sourceKind, resolved.resolvedUrl)
    }
    const bytes = new Uint8Array(resolved.buffer)
    verifyPdfBytes(bytes, sourceKind, resolved.resolvedUrl)
    return {
      bytes,
      blob: new Blob([bytes], { type: "application/pdf" }),
      fileName: sourceKind === "blob"
        ? fileNameFromBlob(source)
        : fileNameFromUrl(resolved.resolvedUrl ?? source.url),
      sourceKind,
      ...(resolved.resolvedUrl ? { resolvedUrl: resolved.resolvedUrl } : {}),
    }
  } catch (error) {
    if (error instanceof PdfSourceError) throw error
    throw fromOfficeError(error, source.kind, source.kind === "url" ? source.url : undefined)
  }
}
