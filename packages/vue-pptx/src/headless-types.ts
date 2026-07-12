import type {
  PptxCapabilityReport,
  PptxPlaybackError,
  PptxPlaybackWarning,
  PptxPreviewDocument,
  PptxPreviewError,
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
  MaybeRefOrGetter,
  Ref,
  ShallowRef,
} from "vue"

export type PptxStageTarget =
  | Readonly<Ref<HTMLElement | null>>
  | (() => HTMLElement | null)

export interface UsePptxDocumentOptions {
  source?: MaybeRefOrGetter<PptxPreviewSource | null | undefined>
  initialSlide?: MaybeRefOrGetter<number | undefined>
  session?: MaybeRefOrGetter<PptxDocumentSessionOptions | undefined>
}

export type PptxDocumentState =
  | "idle"
  | "waiting-for-stage"
  | "loading"
  | "ready"
  | "error"
  | "disposed"

export interface UsePptxDocumentReturn {
  readonly state: Readonly<Ref<PptxDocumentState>>
  readonly error: Readonly<ShallowRef<PptxPreviewError | null>>
  readonly document: Readonly<ShallowRef<PptxPreviewDocument | null>>
  readonly capability: Readonly<ShallowRef<PptxCapabilityReport | null>>
  readonly activeIndex: Readonly<Ref<number>>
  readonly zoomPercent: Readonly<Ref<number>>

  open(source: PptxPreviewSource): Promise<PptxPreviewDocument>
  close(): void
  goTo(index: number): Promise<void>
  nextSlide(): Promise<void>
  previousSlide(): Promise<void>
  setZoom(percent: number): Promise<void>
  getSession(): PptxDocumentSession | null
  dispose(): void
}

export interface UsePptxPlaybackOptions {
  enabled?: MaybeRefOrGetter<boolean>
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

export interface PptxStageExpose {
  readonly element: HTMLElement | null
}
