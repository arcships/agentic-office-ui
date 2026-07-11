export type PptxObjectSource = "slide" | "layout" | "master"

export interface PptxObjectIdentity {
  key: string
  slidePath: string
  source: PptxObjectSource
  shapeId: string
  groupPath: readonly string[]
  name?: string
  explicitMorphName?: string
  creationId?: string
  nodeType: string
}

export type PptxMorphMatchMethod =
  | "explicit-name"
  | "creation-id"
  | "name-and-geometry"
  | "text-and-geometry"

export interface PptxMorphMatch {
  from: string
  to: string
  method: PptxMorphMatchMethod
  confidence: "strong" | "medium" | "weak"
  score: number
  unique: boolean
}

export type PptxTimeContainer =
  | "root"
  | "sequence"
  | "parallel"
  | "exclusive"
  | "behavior"
  | "unknown"

export type PptxTimeNodeKind = "group" | "effect" | "media" | "command" | "unknown"

export type PptxTriggerEvent =
  | "delay"
  | "on-click"
  | "with-previous"
  | "after-previous"
  | "on-shape-click"
  | "on-begin"
  | "on-end"
  | "on-next"
  | "on-previous"
  | "on-media-bookmark"
  | "on-stop-audio"
  | "unknown"

export interface PptxTimeCondition {
  source: "start" | "end" | "previous" | "next"
  event: PptxTriggerEvent
  delayMs: number
  targetObjectKey?: string
  targetNodeId?: string
  bookmarkName?: string
  rawEvent?: string
}

export type PptxEffectKind =
  | "appear"
  | "disappear"
  | "fade-in"
  | "fade-out"
  | "wipe"
  | "scale"
  | "rotate"
  | "motion-path"
  | "emphasis"
  | "set"
  | "media-command"
  | "unknown"

export interface PptxPlaybackEffect {
  id: string
  kind: PptxEffectKind
  targetObjectKey?: string
  paragraphRange?: { start: number; end: number }
  presetClass?: string
  presetId?: string
  transition?: "in" | "out"
  filter?: string
  command?: string
  motionPath?: string
  values: Readonly<Record<string, string | number | boolean>>
}

export interface PptxTimeNode {
  id: string
  parentId?: string
  container: PptxTimeContainer
  kind: PptxTimeNodeKind
  nodeType?: string
  delayMs: number
  durationMs: number | "indefinite"
  repeatCount?: number | "indefinite"
  autoReverse: boolean
  fill: "hold" | "remove" | "freeze" | "unknown"
  restart: "always" | "when-not-active" | "never" | "unknown"
  acceleration: number
  deceleration: number
  conditions: readonly PptxTimeCondition[]
  childIds: readonly string[]
  effect?: PptxPlaybackEffect
  rawSummary?: string
}

export type PptxTrackProperty =
  | "display"
  | "opacity"
  | "translate-x"
  | "translate-y"
  | "scale-x"
  | "scale-y"
  | "rotate"
  | "clip-path"
  | "filter"
  | "fill-color"
  | "line-color"
  | "text-color"
  | "media-playback"
  | "media-time"
  | "media-volume"

export interface PptxTrackKeyframe {
  timeMs: number
  value: string | number | boolean
  easing?: string
  sourceNodeId: string
}

export interface PptxPropertyTrack {
  objectKey: string
  property: PptxTrackProperty
  initialValue: string | number | boolean
  keyframes: readonly PptxTrackKeyframe[]
  endTimeMs: number | "indefinite"
  repeatStartMs?: number
  repeatPeriodMs?: number
  repeatIndefinite?: boolean
}

export interface PptxSlideTransition {
  kind: string
  direction?: string
  option?: string
  durationMs: number
  advanceOnClick: boolean
  advanceAfterMs?: number
  soundRelationId?: string
}

export interface PptxMediaBookmark {
  name: string
  timeMs: number
}

export interface PptxMediaItem {
  id: string
  objectKey?: string
  relationId: string
  sourcePath?: string
  contentType?: string
  kind: "audio" | "video"
  embedded: boolean
  trimStartMs?: number
  trimEndMs?: number
  loop: boolean
  volume: number
  bookmarks: readonly PptxMediaBookmark[]
}

export interface PptxSlideAction {
  id: string
  sourceObjectKey: string
  trigger: "click" | "hover"
  kind: "go-to-slide" | "open-url" | "mailto" | "unsupported"
  targetSlideIndex?: number
  url?: string
  rawAction?: string
}

export type PptxFeatureDisposition = "strict" | "approximate" | "static" | "unparsed"

export interface PptxFeatureRecord {
  id: string
  slideIndex: number
  objectKey?: string
  feature: string
  disposition: PptxFeatureDisposition
  reason?: string
  sourceNodeId?: string
}

export interface PptxCapabilityReport {
  discovered: number
  strict: number
  approximate: number
  static: number
  unparsed: number
  features: readonly PptxFeatureRecord[]
}

export interface PptxPlaybackSlide {
  index: number
  hidden: boolean
  objects: readonly PptxObjectIdentity[]
  morphFromPrevious: readonly PptxMorphMatch[]
  transition?: PptxSlideTransition
  rootNodeId?: string
  nodes: Readonly<Record<string, PptxTimeNode>>
  media: readonly PptxMediaItem[]
  actions: readonly PptxSlideAction[]
  capability: PptxCapabilityReport
}

export interface PptxPlaybackDocument {
  slides: readonly PptxPlaybackSlide[]
  capability: PptxCapabilityReport
}

export interface PptxPlaybackWarning {
  code: string
  message: string
  slideIndex: number
  objectKey?: string
  sourceNodeId?: string
  recoverable: boolean
}
