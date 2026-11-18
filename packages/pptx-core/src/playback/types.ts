export type PptxObjectSource = "slide" | "layout" | "master"

export interface PptxObjectGeometry {
  /** 原始幻灯片坐标，单位 EMU；仅顶层对象保留。 */
  x: number
  y: number
  width: number
  height: number
}

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
  geometry?: PptxObjectGeometry
  text?: string
  /**
   * Hyperlinks carried by this object, in document order. Hover links are
   * reported but never executed automatically; hosts use them for tooltips.
   */
  hyperlinks?: readonly PptxObjectHyperlink[]
}

/** One `a:hlinkClick` / `a:hlinkHover` entry, resolved against the slide relationships. */
export interface PptxObjectHyperlink {
  trigger: "click" | "hover"
  kind: "go-to-slide" | "open-url" | "mailto" | "unsupported"
  targetSlideIndex?: number
  url?: string
  /** Raw `tooltip` attribute; PowerPoint falls back to the URL for external links. */
  tooltip?: string
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

/**
 * Inclusive character range in a text body. Paragraph marks count as one
 * character, matching ECMA-376 `p:charRg` inside `p:txEl`.
 */
export interface PptxCharacterRange {
  start: number
  end: number
}

/** One rendered table cell, addressed with zero-based row and column. */
export interface PptxTableCellTarget {
  row: number
  column: number
}

export type PptxChartElementType =
  | "chart"
  | "series"
  | "category"
  | "point-in-series"
  | "point-in-category"

/** `p:oleChartEl` target inside a chart object. */
export interface PptxChartElementTarget {
  type: PptxChartElementType
  index?: number
  level?: number
}

/** `p:graphicEl` / `p:subSp` target inside an object such as SmartArt. */
export interface PptxGraphicElementTarget {
  /** Raw target token, for example `dgm` or `subSp`. */
  type: string
  /** Resolved child object key when the renderer exposes that identity. */
  objectKey?: string
  sourceId?: string
  index?: number
  level?: number
}

export interface PptxPlaybackEffect {
  id: string
  kind: PptxEffectKind
  targetObjectKey?: string
  /** `cNvPr/@id` of the effect target, kept for element-level target resolution. */
  targetShapeId?: string
  /** Parsed object type of the target (`shape`, `table`, `chart`, `smartart`, ...). */
  targetNodeType?: string
  paragraphRange?: { start: number; end: number }
  characterRange?: PptxCharacterRange
  /** Report label inferred from the range; execution uses `characterRange`. */
  characterGranularity?: "word" | "character"
  /** `p:charRg` exists but the target has no readable text body. */
  characterRangeUnavailable?: boolean
  tableCell?: PptxTableCellTarget
  chartElement?: PptxChartElementTarget
  graphicElement?: PptxGraphicElementTarget
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

export interface PptxSlideTransitionSound {
  relationId: string
  sourcePath?: string
  contentType?: string
  embedded: boolean
  loop: boolean
  volume: number
}

export interface PptxSlideTransition {
  kind: string
  direction?: string
  option?: string
  durationMs: number
  advanceOnClick: boolean
  advanceAfterMs?: number
  soundRelationId?: string
  sound?: PptxSlideTransitionSound
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
  autoplay: boolean
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
  /** Raw `tooltip` attribute from the same `a:hlinkClick` / `a:hlinkHover`. */
  tooltip?: string
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
