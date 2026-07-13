import type {
  PptxPreviewDocument,
  PptxPreviewLimits,
  PptxSearchResult,
} from "./types"
import type { PptxPlaybackError } from "./playback/errors"
import type {
  PptxCapabilityReport,
  PptxPlaybackDocument,
  PptxPlaybackWarning,
} from "./playback/types"

export interface PptxFileLike {
  readonly name?: string
  readonly type?: string
  readonly size?: number
  arrayBuffer(): Promise<ArrayBuffer>
}

export type PptxPreviewSource = ArrayBuffer | Uint8Array | PptxFileLike

export interface PptxDisposableHandle {
  readonly ready?: Promise<void>
  dispose(): void
}

export interface PptxPreviewOpenOptions {
  signal?: AbortSignal
  limits?: PptxPreviewLimits
  initialSlide?: number
  lazyMedia?: boolean
  lazySlides?: boolean
}

export interface PptxPreviewSessionOptions {
  limits?: PptxPreviewLimits
  fitMode?: "contain" | "none"
  zoomPercent?: number
  /** Explicit scroll owner used by list virtualization and Surface zoom anchoring. */
  scrollContainer?: HTMLElement
  lazyMedia?: boolean
  lazySlides?: boolean
  /** Static previews use a vertical page list; presentation playback uses one slide. */
  renderMode?: "list" | "slide"
  listOptions?: {
    windowed?: boolean
    batchSize?: number
    initialSlides?: number
    overscanViewport?: number
    showSlideLabels?: boolean
  }
  onSlideChange?: (index: number) => void
}

export type PptxApproximationPolicy = "off" | "safe"

export interface PptxDocumentSessionOptions extends PptxPreviewSessionOptions {
  approximation?: PptxApproximationPolicy
  externalMedia?: "disabled" | "allowed"
}

export interface PptxPreviewSession {
  readonly document: PptxPreviewDocument | null
  readonly currentSlideIndex: number
  readonly zoomPercent: number
  open(source: PptxPreviewSource, options?: PptxPreviewOpenOptions): Promise<PptxPreviewDocument>
  renderSlide(index: number): Promise<void>
  renderThumbnail(index: number, container: HTMLElement, width?: number): PptxDisposableHandle | null
  searchText(query: string): PptxSearchResult[]
  highlightSearchResult(result: PptxSearchResult): Promise<PptxDisposableHandle | null>
  clearSearchHighlights(): void
  setZoom(percent: number): Promise<void>
  cancel(): void
  dispose(): void
}

export type PptxPlaybackStatus =
  | "idle"
  | "ready"
  | "transitioning"
  | "running"
  | "waiting"
  | "paused"
  | "blocked"
  | "ended"
  | "disposed"

export interface PptxPlaybackSnapshot {
  status: PptxPlaybackStatus
  slideIndex: number
  clickBoundary: number
  positionMs: number
  activeNodeIds: readonly string[]
  blockedMediaIds: readonly string[]
  generation: number
}

export interface PptxPlaybackControllerOptions {
  initialSlide?: number
  skipHiddenSlides?: boolean
  autoplay?: boolean
  approximation?: PptxApproximationPolicy
}

export type PptxActionRequest =
  | { kind: "go-to-slide"; slideIndex: number; sourceObjectKey: string }
  | { kind: "open-url"; url: string; sourceObjectKey: string }
  | { kind: "mailto"; url: string; sourceObjectKey: string }
  | { kind: "unsupported"; action: string; sourceObjectKey: string }

export type PptxPlaybackEvent =
  | { type: "statechange"; snapshot: PptxPlaybackSnapshot }
  | { type: "slidechange"; from: number; to: number; reason: string }
  | { type: "stepchange"; slideIndex: number; boundary: number }
  | { type: "effectstart"; slideIndex: number; nodeId: string; objectKey?: string }
  | { type: "effectend"; slideIndex: number; nodeId: string; objectKey?: string }
  | { type: "warning"; warning: PptxPlaybackWarning }
  | { type: "capability"; report: PptxCapabilityReport }
  | { type: "mediarequest"; mediaId: string; reason: "autoplay-blocked" }
  | { type: "action"; action: PptxActionRequest }
  | { type: "error"; error: PptxPlaybackError }

export interface PptxPlaybackController {
  readonly snapshot: PptxPlaybackSnapshot
  next(): Promise<void>
  activateObject(objectKey: string): Promise<boolean>
  previous(): Promise<void>
  play(): Promise<void>
  pause(): void
  resume(): Promise<void>
  reset(): Promise<void>
  goToSlide(index: number, options?: { includeHidden?: boolean }): Promise<void>
  resumeBlockedMedia(mediaId?: string): Promise<void>
  subscribe(listener: (event: PptxPlaybackEvent) => void): () => void
  dispose(): void
}

export interface PptxDocumentSession extends PptxPreviewSession {
  readonly playbackDocument: PptxPlaybackDocument | null
  readonly capabilityReport: PptxCapabilityReport | null
  createPlaybackController(options?: PptxPlaybackControllerOptions): PptxPlaybackController
}

export type PptxPreviewSessionFactory = (
  container: HTMLElement,
  options?: PptxPreviewSessionOptions,
) => PptxPreviewSession

export type PptxDocumentSessionFactory = (
  container: HTMLElement,
  options?: PptxDocumentSessionOptions,
) => PptxDocumentSession
