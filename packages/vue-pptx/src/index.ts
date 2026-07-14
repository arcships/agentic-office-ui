import PptxViewerComponent from "./PptxViewer.vue"
import "./style.css"

export type {
  OfficeDocumentRevision,
  OfficeObjectReference,
  OfficeReferenceCandidateChange,
  OfficeReferenceConfirmEvent,
  OfficeReferenceError,
  OfficeReferenceResolveEvent,
  OfficeReferenceSurfaceExposed,
  OfficeReferenceSurfaceProps,
  OfficeRegionDraftChange,
  OfficeSelectionCancelEvent,
  OfficeSelectionMode,
} from "@arcships/office-interaction"

export const PptxViewer = PptxViewerComponent
export { default as PptxThumbnail } from "./PptxThumbnail.vue"
export { default as PptxStage } from "./PptxStage.vue"
export { usePptxDocument } from "./composables/usePptxDocument"
export { usePptxPlayback } from "./composables/usePptxPlayback"

export type {
  PptxDocumentState,
  PptxSearchState,
  PptxSearchStatus,
  PptxStageContextMenu,
  PptxStageExpose,
  PptxStageObjectClick,
  PptxStageSelection,
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
