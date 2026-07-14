import type {
  PptxCapabilityReport,
  PptxPlaybackError,
  PptxPlaybackWarning,
  PptxPreviewDocument,
  PptxPreviewError,
  PptxSearchResult,
} from "@arcships/pptx-core"
import type {
  PptxApproximationPolicy,
  PptxDocumentSession,
  PptxDocumentSessionOptions,
  PptxPlaybackController,
  PptxPlaybackEvent,
  PptxPlaybackSnapshot,
  PptxPlaybackStatus,
  PptxPreviewSource,
} from "@arcships/pptx-core/browser"
import type {
  ComputedRef,
  Ref,
  ShallowRef,
} from "vue"
import type { OfficeReferenceSurfaceExposed } from "@arcships/office-interaction"

type ValueOrRefOrGetter<T> = T | Readonly<Ref<T>> | (() => T)

export type PptxStageTarget =
  | Readonly<Ref<HTMLElement | null>>
  | (() => HTMLElement | null)

export interface UsePptxDocumentOptions {
  source?: ValueOrRefOrGetter<PptxPreviewSource | null | undefined>
  initialSlide?: ValueOrRefOrGetter<number | undefined>
  session?: ValueOrRefOrGetter<PptxDocumentSessionOptions | undefined>
}

export type PptxDocumentState =
  | "idle"
  | "waiting-for-stage"
  | "loading"
  | "ready"
  | "error"
  | "disposed"

export type PptxSearchStatus = "idle" | "searching" | "ready" | "error"

export interface PptxSearchState {
  status: PptxSearchStatus
  query: string
  matches: readonly PptxSearchResult[]
  activeIndex: number
  error?: {
    code: "SEARCH_FAILED" | "ACTIVATION_FAILED"
    message: string
  }
}

export interface UsePptxDocumentReturn {
  readonly state: Readonly<Ref<PptxDocumentState>>
  readonly error: Readonly<ShallowRef<PptxPreviewError | null>>
  readonly document: Readonly<ShallowRef<PptxPreviewDocument | null>>
  readonly capability: Readonly<ShallowRef<PptxCapabilityReport | null>>
  readonly activeIndex: Readonly<Ref<number>>
  readonly zoomPercent: Readonly<Ref<number>>
  readonly searchState: Readonly<ShallowRef<PptxSearchState>>

  open(source: PptxPreviewSource): Promise<PptxPreviewDocument>
  close(): void
  goTo(index: number): Promise<void>
  nextSlide(): Promise<void>
  previousSlide(): Promise<void>
  setZoom(percent: number): Promise<void>
  search(query: string): Promise<PptxSearchState>
  activateSearchMatch(index: number): Promise<void>
  searchNext(): Promise<void>
  searchPrevious(): Promise<void>
  clearSearch(): void
  getSearchState(): PptxSearchState
  getSession(): PptxDocumentSession | null
  dispose(): void
}

export interface UsePptxPlaybackOptions {
  enabled?: ValueOrRefOrGetter<boolean>
  autoplay?: boolean
  skipHiddenSlides?: boolean
  approximation?: PptxApproximationPolicy
  onEvent?: (event: PptxPlaybackEvent) => void
}

export interface UsePptxPlaybackReturn {
  readonly controller: Readonly<ShallowRef<PptxPlaybackController | null>>
  readonly snapshot: Readonly<ShallowRef<PptxPlaybackSnapshot | null>>
  readonly status: ComputedRef<PptxPlaybackStatus | "unavailable">
  readonly capability: Readonly<ShallowRef<PptxCapabilityReport | null>>
  readonly lastWarning: Readonly<ShallowRef<PptxPlaybackWarning | null>>
  readonly lastError: Readonly<ShallowRef<PptxPlaybackError | null>>

  next(): Promise<void>
  activateObject(objectKey: string): Promise<boolean>
  previous(): Promise<void>
  play(): Promise<void>
  pause(): void
  resume(): Promise<void>
  reset(): Promise<void>
  goToSlide(index: number, options?: { includeHidden?: boolean }): Promise<void>
  resumeBlockedMedia(mediaId?: string): Promise<void>
  dispose(): void
}

export interface PptxStageExpose extends OfficeReferenceSurfaceExposed {
  readonly element: HTMLElement | null
  readonly scrollContainer: HTMLElement | null
}

export interface PptxStageSelection {
  kind: "slide"
  slideIndex: number
}

export interface PptxStageObjectClick {
  kind: "object"
  slideIndex: number
  objectKey: string
}

interface PptxStageContextMenuPosition {
  slideIndex: number
  clientX: number
  clientY: number
  containerX: number
  containerY: number
}

export type PptxStageContextMenu = PptxStageContextMenuPosition & (
  | { kind: "slide"; objectKey?: never }
  | { kind: "object"; objectKey: string }
)
