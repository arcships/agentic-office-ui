export type PptxPlaybackErrorCode =
  | "PLAYBACK_NOT_READY"
  | "PLAYBACK_DISPOSED"
  | "TARGET_NOT_FOUND"
  | "TARGET_AMBIGUOUS"
  | "TIME_CONDITION_UNSUPPORTED"
  | "TRACK_COMPILE_FAILED"
  | "MEDIA_BLOCKED"
  | "MEDIA_FAILED"
  | "FULLSCREEN_REJECTED"
  | "UNSUPPORTED_FEATURE"

export interface PptxPlaybackErrorDetails {
  slideIndex?: number
  objectKey?: string
  sourceNodeId?: string
  cause?: unknown
}

export class PptxPlaybackError extends Error {
  readonly code: PptxPlaybackErrorCode
  readonly slideIndex?: number
  readonly objectKey?: string
  readonly sourceNodeId?: string
  readonly cause?: unknown

  constructor(code: PptxPlaybackErrorCode, message: string, details: PptxPlaybackErrorDetails = {}) {
    super(message)
    this.name = "PptxPlaybackError"
    this.code = code
    this.slideIndex = details.slideIndex
    this.objectKey = details.objectKey
    this.sourceNodeId = details.sourceNodeId
    this.cause = details.cause
  }
}

