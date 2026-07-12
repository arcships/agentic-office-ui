import PptxViewerComponent from "./PptxViewer.vue"
import "./style.css"

export const PptxViewer = PptxViewerComponent
export { default as PptxThumbnail } from "./PptxThumbnail.vue"
export { default as PptxStage } from "./PptxStage.vue"
export { usePptxDocument } from "./composables/usePptxDocument"
export { usePptxPlayback } from "./composables/usePptxPlayback"

export type {
  PptxDocumentState,
  PptxStageExpose,
  PptxStageTarget,
  UsePptxDocumentOptions,
  UsePptxDocumentReturn,
  UsePptxPlaybackOptions,
  UsePptxPlaybackReturn,
} from "./headless-types"

export type {
  PptxCapabilityReport,
  PptxFeatureDisposition,
  PptxFeatureRecord,
  PptxObjectIdentity,
  PptxPlaybackDocument,
  PptxPlaybackErrorCode,
  PptxPlaybackSlide,
  PptxPlaybackWarning,
  PptxPreviewDocument,
  PptxPreviewErrorCode,
  PptxPreviewLimits,
  PptxPreviewWarning,
  PptxSearchResult,
  PptxSlideInfo,
} from "@arcships/pptx-core"
export type {
  PptxActionRequest,
  PptxApproximationPolicy,
  PptxDocumentSession,
  PptxDocumentSessionFactory,
  PptxDocumentSessionOptions,
  PptxFileLike,
  PptxPlaybackController,
  PptxPlaybackControllerOptions,
  PptxPlaybackEvent,
  PptxPlaybackSnapshot,
  PptxPlaybackStatus,
  PptxPreviewOpenOptions,
  PptxPreviewSession,
  PptxPreviewSessionFactory,
  PptxPreviewSessionOptions,
  PptxPreviewSource,
} from "@arcships/pptx-core/browser"
