import type {
  PptxChartElementTarget,
  PptxFeatureDisposition,
  PptxGraphicElementTarget,
  PptxPlaybackEffect,
  PptxTableCellTarget,
} from "./types"

/**
 * Stable capability feature names for the playback v2 element targets.
 *
 * Hosts key their own policy and telemetry on these strings, so they must stay
 * stable even when the implementation behind a code changes. The parser and
 * the controller share the names through this module instead of repeating
 * string literals.
 */
export const PPTX_CAPABILITY_FEATURES = Object.freeze({
  /** A `p:charRg` range that covers whole whitespace-delimited words. */
  textByWord: "text:by-word",
  /** A `p:charRg` range that covers part of a word or a single character. */
  textByCharacter: "text:by-character",
  /** `p:charRg` could not be mapped onto a rendered text body. */
  textRangeUnavailable: "text:character-range-unavailable",
  /** A table cell targeted through its flattened paragraph range. */
  tableCell: "table:cell",
  /** Table build metadata that only groups explicit cell effects. */
  tableBuild: "table-build",
  /** Chart element targets the renderer paints as one canvas. */
  chartSeries: "chart:series",
  chartCategory: "chart:category",
  chartPoint: "chart:point",
  chartElement: "chart:element",
  /** SmartArt node/level targets resolved through rendered child identities. */
  smartArtNode: "smartart:node",
  smartArtLevel: "smartart:level",
  /** `p:subSp` inside an ordinary group. */
  groupSubShape: "group:sub-shape",
  /** Slide transition sound. */
  transitionSound: "transition:sound",
  /** Prefix for a 3D transition that is intentionally reported only. */
  transition3d: "transition:3d:",
} as const)

export type PptxCapabilityFeatureCode =
  | (typeof PPTX_CAPABILITY_FEATURES)[Exclude<keyof typeof PPTX_CAPABILITY_FEATURES, "transition3d">]
  | `${typeof PPTX_CAPABILITY_FEATURES.transition3d}${string}`

export interface PptxEffectTargetSupport {
  /** `strict` when the rendered slide can resolve the target, otherwise `static`. */
  disposition: Extract<PptxFeatureDisposition, "strict" | "static">
  /** Stable capability code for the element-level target, when it has one. */
  feature?: PptxCapabilityFeatureCode
  reason?: string
}

export interface PptxTextGranularityInfo {
  granularity: "word" | "character"
  feature: PptxCapabilityFeatureCode
}

const CHARACTER_TARGET = /#text:(\d+)-(\d+)$/u
const CELL_TARGET = /#cell:(\d+):(\d+)$/u

/** `${objectKey}#paragraph:start-end` */
export function pptxParagraphTarget(
  objectKey: string,
  range: { start: number; end: number },
): string {
  return `${objectKey}#paragraph:${range.start}-${range.end}`
}

/** `${objectKey}#text:start-end`, a character range in the whole text body. */
export function pptxCharacterTarget(
  objectKey: string,
  range: { start: number; end: number },
): string {
  return `${objectKey}#text:${range.start}-${range.end}`
}

/** `${objectKey}#cell:row:column` */
export function pptxTableCellKey(objectKey: string, cell: PptxTableCellTarget): string {
  return `${objectKey}#cell:${cell.row}:${cell.column}`
}

function chartTypeToken(type: PptxChartElementTarget["type"]): string {
  switch (type) {
    case "series": return "series"
    case "category": return "category"
    case "point-in-series": return "point-in-series"
    case "point-in-category": return "point-in-category"
    default: return "chart"
  }
}

/** `${objectKey}#chart:type[:index][@level]` */
export function pptxChartElementKey(objectKey: string, element: PptxChartElementTarget): string {
  const index = element.index === undefined ? "" : `:${element.index}`
  const level = element.level === undefined ? "" : `@${element.level}`
  return `${objectKey}#chart:${chartTypeToken(element.type)}${index}${level}`
}

/** `${objectKey}#graphic:type[:sourceId|index][@level]` */
export function pptxGraphicElementKey(objectKey: string, element: PptxGraphicElementTarget): string {
  if (element.objectKey) return element.objectKey
  const source = element.sourceId !== undefined
    ? `:${element.sourceId}`
    : element.index !== undefined
      ? `:${element.index}`
      : ""
  const level = element.level === undefined ? "" : `@${element.level}`
  return `${objectKey}#graphic:${element.type}${source}${level}`
}

/**
 * The state-map key an effect writes. Object targets keep their object key;
 * every finer target appends a selector that the browser track index expands.
 * `undefined` means the effect has no resolvable target at all.
 */
export function pptxEffectTargetKey(effect: PptxPlaybackEffect): string | undefined {
  const objectKey = effect.targetObjectKey
  if (!objectKey) return undefined
  if (effect.chartElement) return pptxChartElementKey(objectKey, effect.chartElement)
  if (effect.graphicElement) return pptxGraphicElementKey(objectKey, effect.graphicElement)
  if (effect.characterRange) return pptxCharacterTarget(objectKey, effect.characterRange)
  if (effect.tableCell) return pptxTableCellKey(objectKey, effect.tableCell)
  if (effect.paragraphRange) return pptxParagraphTarget(objectKey, effect.paragraphRange)
  return objectKey
}

/** Base object key for a state-map key produced by `pptxEffectTargetKey`. */
export function pptxTargetObjectKey(target: string): string {
  return target.replace(/(?:#paragraph:|#text:|#cell:|#chart:|#graphic:).*$/u, "")
}

export function parsePptxCharacterTarget(target: string): { start: number; end: number } | undefined {
  const match = CHARACTER_TARGET.exec(target)
  return match ? { start: Number(match[1]), end: Number(match[2]) } : undefined
}

export function parsePptxTableCellTarget(target: string): PptxTableCellTarget | undefined {
  const match = CELL_TARGET.exec(target)
  return match ? { row: Number(match[1]), column: Number(match[2]) } : undefined
}

/**
 * How a `p:charRg` range should be reported. Execution always applies the
 * character range; the label only distinguishes whole-word ranges from
 * partial character ranges, because PowerPoint writes the same XML for both.
 */
export function pptxTextGranularity(
  text: string | undefined,
  range: { start: number; end: number },
): PptxTextGranularityInfo {
  const value = text ?? ""
  const start = Math.max(0, range.start)
  const end = Math.min(value.length - 1, range.end)
  const slice = end >= start ? value.slice(start, end + 1) : ""
  const before = start > 0 ? value[start - 1] : undefined
  const after = end + 1 < value.length ? value[end + 1] : undefined
  const boundary = (character: string | undefined): boolean => character === undefined || /\s/u.test(character)
  const wholeWord = slice.length > 0
    && !/\s/u.test(slice)
    && boundary(before)
    && boundary(after)
  return wholeWord
    ? { granularity: "word", feature: PPTX_CAPABILITY_FEATURES.textByWord }
    : { granularity: "character", feature: PPTX_CAPABILITY_FEATURES.textByCharacter }
}

function chartElementSupport(
  element: PptxChartElementTarget,
  nodeType: string | undefined,
): PptxEffectTargetSupport {
  if (nodeType !== undefined && nodeType !== "chart") return {
    disposition: "static",
    feature: chartFeature(element),
    reason: `目标对象类型 ${nodeType} 不是图表，无法解析图表内部元素。`,
  }
  if (element.type === "chart") return { disposition: "strict", feature: PPTX_CAPABILITY_FEATURES.chartElement }
  const feature = chartFeature(element)
  return {
    disposition: "static",
    feature,
    reason: "第三方渲染器把图表绘制为单个 canvas，没有按系列、分类或数据点的元素身份；该效果保留静态。",
  }
}

function chartFeature(element: PptxChartElementTarget): PptxCapabilityFeatureCode {
  switch (element.type) {
    case "series": return PPTX_CAPABILITY_FEATURES.chartSeries
    case "category": return PPTX_CAPABILITY_FEATURES.chartCategory
    case "point-in-series":
    case "point-in-category": return PPTX_CAPABILITY_FEATURES.chartPoint
    default: return PPTX_CAPABILITY_FEATURES.chartElement
  }
}

function graphicElementSupport(
  element: PptxGraphicElementTarget,
  nodeType: string | undefined,
): PptxEffectTargetSupport {
  const smartArt = nodeType === "smartart" || nodeType === "diagram"
  const feature = smartArt
    ? PPTX_CAPABILITY_FEATURES.smartArtNode
    : PPTX_CAPABILITY_FEATURES.groupSubShape
  if (element.objectKey) return { disposition: "strict", feature }
  if (smartArt && element.level !== undefined && element.sourceId === undefined && element.index === undefined) {
    return {
      disposition: "static",
      feature: PPTX_CAPABILITY_FEATURES.smartArtLevel,
      reason: "SmartArt 按层级的动画没有可解析的子对象身份，该效果保留静态。",
    }
  }
  return {
    disposition: "static",
    feature,
    reason: "渲染器没有暴露该内部元素的唯一 DOM 身份，该效果保留静态。",
  }
}

/**
 * Target-only capability decision. The parser reports this instead of guessing
 * from the effect kind, and the track compiler skips effects the renderer
 * cannot address.
 */
export function pptxEffectTargetSupport(effect: PptxPlaybackEffect): PptxEffectTargetSupport {
  if (!effect.targetObjectKey) return { disposition: "static" }
  if (effect.chartElement) return chartElementSupport(effect.chartElement, effect.targetNodeType)
  if (effect.graphicElement) return graphicElementSupport(effect.graphicElement, effect.targetNodeType)
  if (effect.characterRangeUnavailable) return {
    disposition: "static",
    feature: PPTX_CAPABILITY_FEATURES.textRangeUnavailable,
    reason: "字符范围没有对应的可读文字体，该效果保留静态。",
  }
  if (effect.tableCell) return {
    disposition: effect.targetNodeType === undefined || effect.targetNodeType === "table"
      ? "strict"
      : "static",
    feature: PPTX_CAPABILITY_FEATURES.tableCell,
  }
  if (effect.characterRange) return {
    disposition: "strict",
    feature: effect.characterGranularity === "word"
      ? PPTX_CAPABILITY_FEATURES.textByWord
      : PPTX_CAPABILITY_FEATURES.textByCharacter,
  }
  return { disposition: "strict" }
}

/** Effects whose element target is known not to be renderable stay out of tracks. */
export function isPptxEffectTargetExecutable(effect: PptxPlaybackEffect): boolean {
  return pptxEffectTargetSupport(effect).disposition === "strict"
}

/**
 * Character offsets address the text body the same way ECMA-376 does: every
 * paragraph contributes its text followed by one paragraph mark. This is the
 * layout the parser and the browser text-run splitter must agree on.
 */
export function pptxTextBodyOffset(paragraphs: readonly string[], paragraph: number): number {
  let offset = 0
  for (let index = 0; index < paragraph; index += 1) offset += (paragraphs[index] ?? "").length + 1
  return offset
}

export function pptxTextBody(paragraphs: readonly string[]): string {
  return paragraphs.join("\r")
}
