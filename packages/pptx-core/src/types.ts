export type PptxPreviewErrorCode =
  | "ABORTED"
  | "INVALID_SOURCE"
  | "LIMIT_EXCEEDED"
  | "PARSE_FAILED"
  | "RENDER_FAILED"
  | "STALE_RESULT"
  | "UNSUPPORTED_CONTENT"

export class PptxPreviewError extends Error {
  readonly code: PptxPreviewErrorCode
  readonly cause?: unknown

  constructor(code: PptxPreviewErrorCode, message: string, cause?: unknown) {
    super(message)
    this.name = "PptxPreviewError"
    this.code = code
    this.cause = cause
  }
}

export interface PptxPreviewLimits {
  maxInputBytes?: number
  maxArchiveEntries?: number
  maxSingleEntryBytes?: number
  maxUncompressedBytes?: number
  maxMediaBytes?: number
  maxConcurrency?: number
}

export interface PptxSlideInfo {
  index: number
  number: number
  hidden: boolean
}

export interface PptxPreviewWarning {
  code: "NOTES_UNSUPPORTED" | "UNSUPPORTED_CONTENT"
  message: string
  slideIndex?: number
}

export interface PptxPreviewDocument {
  fileName?: string
  width: number
  height: number
  slides: readonly PptxSlideInfo[]
  warnings: readonly PptxPreviewWarning[]
}

export interface PptxSearchResult {
  slideIndex: number
  nodeId: string
  text: string
  snippet: string
  matchStart: number
  matchEnd: number
}
