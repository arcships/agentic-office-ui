export {
  PptxPreviewError,
} from "./types"
export type {
  PptxPreviewDocument,
  PptxPreviewErrorCode,
  PptxPreviewLimits,
  PptxPreviewWarning,
  PptxSearchResult,
  PptxSlideInfo,
} from "./types"
export {
  PptxPlaybackError,
} from "./playback/errors"
export type {
  PptxPlaybackErrorCode,
  PptxPlaybackErrorDetails,
} from "./playback/errors"
export {
  createPptxCapabilityReport,
} from "./playback/capability"
export {
  createPptxObjectKey,
  matchPptxMorphObjects,
} from "./playback/identity"
export {
  compilePptxSlideSchedule,
} from "./playback/time-tree"
export {
  isPptxTransitionExecutable,
  pptxTransitionCapability,
  pptxTransitionDirection,
  pptxTransitionProgress,
  pptxTransitionSplitMask,
  pptxTransitionWipeClip,
  pptxTransitionWipeInsets,
} from "./playback/transitions"
export type {
  PptxTransitionCapability,
  PptxTransitionDirection,
} from "./playback/transitions"
export type {
  PptxScheduledEffect,
  PptxSlideSchedule,
  PptxTriggerKey,
  PptxTriggerSchedule,
} from "./playback/time-tree"
export {
  compilePptxSlideTracks,
  evaluatePptxTrack,
  evaluatePptxTriggerGroup,
  rebuildPptxStateAtBoundary,
} from "./playback/track-compiler"
export type {
  PptxCompiledPropertyTrack,
  PptxCompiledTrackSegment,
  PptxCompiledSlide,
  PptxCompiledTriggerGroup,
  PptxObjectPropertyState,
  PptxObjectStateMap,
  PptxTrackValue,
} from "./playback/track-compiler"
export type {
  PptxObjectKeyParts,
} from "./playback/identity"
export {
  createPptxBehaviorDraftFromRecord,
  createPptxBehaviorReferenceDraft,
  createPptxObjectReferenceDraft,
  createPptxRegionReferenceDraft,
  createPptxSlideReferenceDraft,
  createPptxSubObjectReferenceDraft,
  createPptxTextReferenceDraft,
  describePptxReference,
  enumeratePptxBehaviors,
  pptxReferenceKindForObject,
  resolvePptxReference,
} from "./reference-adapter"
export type {
  PptxBehaviorTarget,
  PptxChartSnapshot,
  PptxObjectReferenceOptions,
  PptxOfficeReference,
  PptxOfficeReferenceDraft,
  PptxReferenceContext,
  PptxSubObjectTarget,
  PptxTableSnapshot,
} from "./reference-adapter"
export {
  enumeratePptxSlideBehaviors,
  PPTX_BEHAVIOR_REFERENCE_KIND,
} from "./pptx-behaviors"
export type {
  PptxBehaviorContext,
  PptxBehaviorKind,
  PptxBehaviorRecord,
} from "./pptx-behaviors"
export * from "./pptx-object-index"
export type {
  PptxCapabilityReport,
  PptxEffectKind,
  PptxFeatureDisposition,
  PptxFeatureRecord,
  PptxMediaBookmark,
  PptxMediaItem,
  PptxMorphMatch,
  PptxMorphMatchMethod,
  PptxObjectHyperlink,
  PptxObjectIdentity,
  PptxObjectSource,
  PptxPlaybackDocument,
  PptxPlaybackEffect,
  PptxPlaybackSlide,
  PptxPlaybackWarning,
  PptxPropertyTrack,
  PptxSlideAction,
  PptxSlideTransition,
  PptxTimeCondition,
  PptxTimeContainer,
  PptxTimeNode,
  PptxTimeNodeKind,
  PptxTrackKeyframe,
  PptxTrackProperty,
  PptxTriggerEvent,
} from "./playback/types"

/**
 * Hard PPTX preview budgets. Requested limits can only tighten these values;
 * they are also the defaults every session starts from.
 */
export const DEFAULT_PPTX_PREVIEW_LIMITS = Object.freeze({
  maxInputBytes: 50 * 1024 * 1024,
  maxArchiveEntries: 4_000,
  maxSingleEntryBytes: 32 * 1024 * 1024,
  maxUncompressedBytes: 256 * 1024 * 1024,
  maxMediaBytes: 192 * 1024 * 1024,
  maxConcurrency: 8,
  maxCompressionRatio: 100,
  maxArchivePathLength: 1_024,
  maxXmlBytes: 256 * 1024 * 1024,
  maxSingleXmlBytes: 32 * 1024 * 1024,
  maxXmlDepth: 256,
  maxXmlAttributeBytes: 1 * 1024 * 1024,
  maxTextNodeBytes: 16 * 1024 * 1024,
  maxRelationships: 100_000,
})
