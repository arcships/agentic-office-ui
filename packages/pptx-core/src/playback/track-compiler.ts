import { compilePptxSlideSchedule, type PptxScheduledEffect, type PptxSlideSchedule, type PptxTriggerKey } from "./time-tree"
import { isMediaPlayCommand, mediaSeekMs } from "./media-command"
import { isPptxEffectTargetExecutable, pptxEffectTargetKey } from "./targets"
import type {
  PptxChartElementTarget,
  PptxGraphicElementTarget,
  PptxMediaItem,
  PptxPlaybackEffect,
  PptxPlaybackSlide,
  PptxPropertyTrack,
  PptxTableCellTarget,
  PptxTrackKeyframe,
  PptxTrackProperty,
  PptxTimeNode,
} from "./types"

export type PptxTrackValue = string | number | boolean
export type PptxObjectPropertyState = Partial<Record<PptxTrackProperty, PptxTrackValue>>
export type PptxObjectStateMap = Readonly<Record<string, Readonly<PptxObjectPropertyState>>>

export interface PptxCompiledPropertyTrack extends PptxPropertyTrack {
  /** Exact state-map key, including paragraph, character, cell or element selectors. */
  target?: string
  paragraphRange?: { start: number; end: number }
  characterRange?: { start: number; end: number }
  tableCell?: PptxTableCellTarget
  chartElement?: PptxChartElementTarget
  graphicElement?: PptxGraphicElementTarget
  segments: readonly PptxCompiledTrackSegment[]
}

export interface PptxCompiledTrackSegment {
  startMs: number
  endMs: number
  from: PptxTrackValue
  to: PptxTrackValue
  sourceNodeId: string
  order: number
  fill: PptxTimeNode["fill"]
  easing?: string
  repeatStartMs?: number
  repeatPeriodMs?: number
  repeatIndefinite?: boolean
}

export interface PptxCompiledTriggerGroup {
  key: PptxTriggerKey
  clickBoundary: number
  durationMs: number | "indefinite"
  tracks: readonly PptxCompiledPropertyTrack[]
  effects: readonly PptxScheduledEffect[]
}

export interface PptxCompiledSlide {
  slideIndex: number
  schedule: PptxSlideSchedule
  groups: readonly PptxCompiledTriggerGroup[]
  initialState: PptxObjectStateMap
}

interface TrackFragment extends PptxCompiledTrackSegment {
  objectKey: string
  target: string
  paragraphRange?: { start: number; end: number }
  characterRange?: { start: number; end: number }
  tableCell?: PptxTableCellTarget
  chartElement?: PptxChartElementTarget
  graphicElement?: PptxGraphicElementTarget
  property: PptxTrackProperty
  initialValue: PptxTrackValue
}

const defaultValues: Record<PptxTrackProperty, PptxTrackValue> = {
  display: true,
  opacity: 1,
  "translate-x": 0,
  "translate-y": 0,
  "scale-x": 1,
  "scale-y": 1,
  rotate: 0,
  "clip-path": "none",
  filter: "none",
  "fill-color": "",
  "line-color": "",
  "text-color": "",
  "media-playback": "stopped",
  "media-time": 0,
  "media-volume": 1,
}

function initialValue(
  initialState: PptxObjectStateMap,
  target: string,
  property: PptxTrackProperty,
): PptxTrackValue {
  return initialState[target]?.[property] ?? defaultValues[property]
}

function initialNumber(
  initialState: PptxObjectStateMap,
  target: string,
  property: PptxTrackProperty,
  fallback: number,
): number {
  const value = initialValue(initialState, target, property)
  return typeof value === "number" && Number.isFinite(value) ? value : fallback
}

function numberValue(effect: PptxPlaybackEffect, name: string, fallback: number): number {
  const value = effect.values[name]
  return typeof value === "number" && Number.isFinite(value) ? value : fallback
}

function motionEnd(effect: PptxPlaybackEffect): { x: number; y: number } | undefined {
  const path = effect.motionPath
  if (!path || !/^[\s\d.,+\-mlehvz]+$/iu.test(path)) return undefined
  const numbers = [...path.matchAll(/[+-]?(?:\d+(?:\.\d+)?|\.\d+)/gu)].map((match) => Number(match[0]))
  if (numbers.length < 2) return undefined
  const rawX = numbers[numbers.length - 2]
  const rawY = numbers[numbers.length - 1]
  return {
    x: rawX * (Math.abs(rawX) <= 10 ? numberValue(effect, "motionScaleX", 1) : 1),
    y: rawY * (Math.abs(rawY) <= 10 ? numberValue(effect, "motionScaleY", 1) : 1),
  }
}

function easingValue(node: PptxTimeNode): string | undefined {
  return node.acceleration > 0 || node.deceleration > 0
    ? `pptx(${node.acceleration},${node.deceleration})`
    : undefined
}

function interpolateValue(from: PptxTrackValue, to: PptxTrackValue, progress: number): PptxTrackValue {
  return typeof from === "number" && typeof to === "number"
    ? from + (to - from) * progress
    : progress >= 1 ? to : from
}

function expandRepeats(fragments: readonly TrackFragment[], node: PptxTimeNode): TrackFragment[] {
  if (!fragments.length || node.durationMs === "indefinite" || node.durationMs <= 0) return [...fragments]
  const repeat = node.repeatCount === "indefinite" ? "indefinite" : Math.max(node.repeatCount ?? 1, 1)
  const repeats = repeat === "indefinite" ? 1 : Math.ceil(repeat)
  const period = node.durationMs * (node.autoReverse ? 2 : 1)
  const result: TrackFragment[] = []
  const effectStartMs = Math.min(...fragments.map((fragment) => fragment.startMs))
  for (const source of fragments) {
    if (source.startMs === source.endMs || ["display", "media-playback", "media-time", "media-volume"].includes(source.property)) {
      if (source.startMs === source.endMs && source.startMs === effectStartMs && node.durationMs > 0) {
        result.push({
          ...source,
          endMs: source.startMs + (repeat === "indefinite" ? period : repeat * period),
          from: source.to,
          repeatStartMs: repeat === "indefinite" ? source.startMs : undefined,
          repeatPeriodMs: repeat === "indefinite" ? period : undefined,
          repeatIndefinite: repeat === "indefinite" || undefined,
        })
      } else result.push(source)
      continue
    }
    for (let cycle = 0; cycle < repeats; cycle += 1) {
      const remaining = repeat === "indefinite" ? 1 : Math.min(1, repeat - cycle)
      if (remaining <= 0) break
      const start = source.startMs + cycle * period
      const forwardEnd = start + node.durationMs * remaining
      const forwardTo = interpolateValue(source.from, source.to, remaining)
      result.push({
        ...source,
        startMs: start,
        endMs: forwardEnd,
        to: forwardTo,
        repeatStartMs: repeat === "indefinite" ? source.startMs : undefined,
        repeatPeriodMs: repeat === "indefinite" ? period : undefined,
        repeatIndefinite: repeat === "indefinite" || undefined,
      })
      if (node.autoReverse && remaining === 1) result.push({
        ...source,
        startMs: forwardEnd,
        endMs: forwardEnd + node.durationMs,
        from: source.to,
        to: source.from,
        repeatStartMs: repeat === "indefinite" ? source.startMs : undefined,
        repeatPeriodMs: repeat === "indefinite" ? period : undefined,
        repeatIndefinite: repeat === "indefinite" || undefined,
      })
    }
  }
  return result
}

function fragmentsForEffect(
  scheduled: PptxScheduledEffect,
  initialState: PptxObjectStateMap,
  node: PptxTimeNode,
  /** State before entrance effects hide their targets, i.e. the element's static transform. */
  baselineState: PptxObjectStateMap = initialState,
): TrackFragment[] {
  const effect = scheduled.effect
  const objectKey = effect.targetObjectKey
  const target = pptxEffectTargetKey(effect)
  if (!objectKey || !target || !isPptxEffectTargetExecutable(effect)) return []
  const baseDuration = typeof node.durationMs === "number" ? node.durationMs : 0
  const baseEnd = scheduled.startMs + baseDuration
  // Exit effects must hide the element once every repeat pass has finished.
  const terminalEnd = scheduled.endMs === "indefinite" ? undefined : scheduled.endMs
  const rotateBase = initialNumber(baselineState, target, "rotate", 0)
  const translateBaseX = initialNumber(baselineState, target, "translate-x", 0)
  const translateBaseY = initialNumber(baselineState, target, "translate-y", 0)
  const scaleBaseX = initialNumber(baselineState, target, "scale-x", 1)
  const scaleBaseY = initialNumber(baselineState, target, "scale-y", 1)
  const fragment = (
    property: PptxTrackProperty,
    from: PptxTrackValue,
    to: PptxTrackValue,
    startMs = scheduled.startMs,
    endMs = baseEnd,
  ): TrackFragment => ({
    objectKey,
    target,
    paragraphRange: effect.paragraphRange,
    characterRange: effect.characterRange,
    tableCell: effect.tableCell,
    chartElement: effect.chartElement,
    graphicElement: effect.graphicElement,
    property,
    initialValue: initialValue(initialState, target, property),
    startMs,
    endMs,
    from,
    to,
    sourceNodeId: scheduled.nodeId,
    order: scheduled.order,
    fill: node.fill,
    easing: easingValue(node),
  })
  let fragments: TrackFragment[]
  switch (effect.kind) {
    case "appear":
      fragments = [fragment("display", false, true, scheduled.startMs, scheduled.startMs)]
      break
    case "disappear":
      fragments = [fragment("display", true, false, scheduled.startMs, scheduled.startMs)]
      break
    case "fade-in":
      fragments = [
        fragment("display", false, true, scheduled.startMs, scheduled.startMs),
        fragment("opacity", 0, 1),
      ]
      break
    case "fade-out":
      fragments = terminalEnd === undefined
        ? [fragment("opacity", 1, 0)]
        : [
            fragment("opacity", 1, 0),
            fragment("display", true, false, terminalEnd, terminalEnd),
          ]
      break
    case "scale":
      {
        const fallback = effect.presetClass === "emph" ? 1 : 0
        const fromX = numberValue(effect, "scaleFromX", fallback)
        const fromY = numberValue(effect, "scaleFromY", fallback)
        fragments = [
          fragment(
            "scale-x",
            scaleBaseX * fromX,
            scaleBaseX * numberValue(effect, "scaleToX", fromX * numberValue(effect, "scaleByX", 1)),
          ),
          fragment(
            "scale-y",
            scaleBaseY * fromY,
            scaleBaseY * numberValue(effect, "scaleToY", fromY * numberValue(effect, "scaleByY", 1)),
          ),
        ]
      }
      break
    case "emphasis": {
      const property = effect.values.emphasisProperty
      if (property === "opacity") {
        const from = numberValue(effect, "emphasisFrom", initialValue(initialState, target, "opacity") as number)
        fragments = [fragment("opacity", from, numberValue(effect, "emphasisTo", from))]
      } else fragments = []
      break
    }
    case "rotate":
      fragments = [fragment("rotate", rotateBase, rotateBase + numberValue(effect, "rotationBy", 0))]
      break
    case "motion-path": {
      const end = motionEnd(effect)
      fragments = end
        ? [
            fragment("translate-x", translateBaseX, translateBaseX + end.x),
            fragment("translate-y", translateBaseY, translateBaseY + end.y),
          ]
        : []
      break
    }
    case "wipe": {
      const direction = effect.filter?.match(/wipe\(([^)]+)\)/iu)?.[1]?.toLowerCase()
      if (direction !== "left" && direction !== "right") {
        fragments = []
        break
      }
      const from = direction === "right" ? "inset(0 0 0 100%)" : "inset(0 100% 0 0)"
      fragments = [fragment("clip-path", from, "inset(0 0 0 0)")]
      break
    }
    case "media-command": {
      const command = effect.command ?? "play"
      fragments = [fragment("media-playback", "stopped", command)]
      const seekMs = mediaSeekMs(effect.command)
      if (seekMs !== undefined) fragments.push(fragment("media-time", seekMs, seekMs))
      break
    }
    default:
      fragments = []
  }
  return expandRepeats(fragments, node)
}

function trackKey(fragment: TrackFragment): string {
  return `${fragment.target}|${fragment.property}`
}

function mediaFragment(
  item: PptxMediaItem,
  property: PptxTrackProperty,
  from: PptxTrackValue,
  to: PptxTrackValue,
  order: number,
): TrackFragment | undefined {
  if (!item.objectKey) return undefined
  return {
    objectKey: item.objectKey,
    target: item.objectKey,
    property,
    initialValue: from,
    startMs: 0,
    endMs: 0,
    from,
    to,
    sourceNodeId: `media:${item.id}`,
    order,
    fill: "hold",
  }
}

/**
 * `p:cMediaNode` 自身的音量与自动播放在时间零就生效，不必等 `p:cmd`。
 * 已经被媒体命令轨道覆盖的播放状态不重复生成，避免同一属性双写。
 */
function mediaFragments(
  media: readonly PptxMediaItem[],
  coveredPlaybackKeys: ReadonlySet<string>,
): TrackFragment[] {
  const result: TrackFragment[] = []
  media.forEach((item, index) => {
    const key = item.objectKey
    if (!key) return
    const base = 1_000_000 + index * 2
    const volume = Math.max(0, Math.min(1, item.volume))
    const volumeFragment = mediaFragment(item, "media-volume", volume, volume, base)
    if (volumeFragment) result.push(volumeFragment)
    if (!item.autoplay || coveredPlaybackKeys.has(key)) return
    const playback = mediaFragment(item, "media-playback", "stopped", "play", base + 1)
    if (playback) result.push(playback)
  })
  return result
}

function compileFragments(fragments: readonly TrackFragment[]): readonly PptxCompiledPropertyTrack[] {
  const groups = new Map<string, TrackFragment[]>()
  for (const fragment of fragments) {
    const key = trackKey(fragment)
    const list = groups.get(key) ?? []
    list.push(fragment)
    groups.set(key, list)
  }
  return Object.freeze([...groups.values()].map((items): PptxCompiledPropertyTrack => {
    const ordered = [...items].sort((left, right) => left.startMs - right.startMs || left.order - right.order)
    const frames = new Map<number, PptxTrackKeyframe>()
    for (const item of ordered) {
      frames.set(item.startMs, Object.freeze({
        timeMs: item.startMs,
        value: item.from,
        easing: item.easing,
        sourceNodeId: item.sourceNodeId,
      }))
      frames.set(item.endMs, Object.freeze({
        timeMs: item.endMs,
        value: item.to,
        sourceNodeId: item.sourceNodeId,
      }))
    }
    const keyframes = Object.freeze([...frames.values()].sort((left, right) => left.timeMs - right.timeMs))
    const first = ordered[0]
    return Object.freeze({
      objectKey: first.objectKey,
      target: first.target,
      paragraphRange: first.paragraphRange,
      characterRange: first.characterRange,
      tableCell: first.tableCell,
      chartElement: first.chartElement,
      graphicElement: first.graphicElement,
      property: first.property,
      initialValue: first.initialValue,
      keyframes,
      segments: Object.freeze(ordered.map((item) => Object.freeze({
        startMs: item.startMs,
        endMs: item.endMs,
        from: item.from,
        to: item.to,
        sourceNodeId: item.sourceNodeId,
        order: item.order,
        fill: item.fill,
        easing: item.easing,
        repeatStartMs: item.repeatStartMs,
        repeatPeriodMs: item.repeatPeriodMs,
        repeatIndefinite: item.repeatIndefinite,
      }))),
      endTimeMs: ordered.some((item) => item.repeatIndefinite)
        ? "indefinite"
        : Math.max(0, ...ordered.map((item) => item.endMs)),
      repeatStartMs: ordered.find((item) => item.repeatIndefinite)?.repeatStartMs,
      repeatPeriodMs: ordered.find((item) => item.repeatIndefinite)?.repeatPeriodMs,
      repeatIndefinite: ordered.some((item) => item.repeatIndefinite) || undefined,
    })
  }))
}

export function compilePptxSlideTracks(
  slide: PptxPlaybackSlide,
  initialState: PptxObjectStateMap = Object.freeze({}),
): PptxCompiledSlide {
  const schedule = compilePptxSlideSchedule(slide)
  const preparedState: Record<string, PptxObjectPropertyState> = Object.fromEntries(
    Object.entries(initialState).map(([key, value]) => [key, { ...value }]),
  )
  const sourceState = Object.freeze(Object.fromEntries(
    Object.entries(preparedState).map(([key, value]) => [key, Object.freeze({ ...value })]),
  ))
  const compilableNodeIds = new Set<string>()
  for (const scheduled of schedule.groups.flatMap((group) => group.effects)) {
    const node = slide.nodes[scheduled.nodeId]
    if (node && fragmentsForEffect(scheduled, sourceState, node).length > 0) {
      compilableNodeIds.add(scheduled.nodeId)
    }
  }
  const playbackCoveredKeys = new Set(
    schedule.groups
      .filter((group) => group.key === "auto")
      .flatMap((group) => group.effects)
      .filter((scheduled) => scheduled.effect.kind === "media-command")
      .flatMap((scheduled) => scheduled.effect.targetObjectKey
        ? [scheduled.effect.targetObjectKey]
        : []),
  )
  for (const scheduled of schedule.groups.flatMap((group) => group.effects)) {
    const effect = scheduled.effect
    const target = pptxEffectTargetKey(effect)
    if (!target || effect.presetClass !== "entr" || !compilableNodeIds.has(scheduled.nodeId)) continue
    const state = preparedState[target] ?? {}
    if (effect.kind === "appear") state.display = false
    else if (effect.kind === "fade-in") {
      state.display = false
      state.opacity = 0
    } else if (effect.kind === "wipe") state["clip-path"] = effect.filter?.includes("right")
      ? "inset(0 0 0 100%)"
      : "inset(0 100% 0 0)"
    else if (effect.kind === "scale") {
      const baseX = typeof state["scale-x"] === "number" ? state["scale-x"] : 1
      const baseY = typeof state["scale-y"] === "number" ? state["scale-y"] : 1
      state["scale-x"] = baseX * numberValue(effect, "scaleFromX", 0)
      state["scale-y"] = baseY * numberValue(effect, "scaleFromY", 0)
    }
    preparedState[target] = state
  }
  const frozenInitialState = Object.freeze(Object.fromEntries(
    Object.entries(preparedState).map(([key, value]) => [key, Object.freeze(value)]),
  ))
  const mediaTracks = mediaFragments(slide.media, playbackCoveredKeys)
  const compiledMediaTracks = compileFragments(mediaTracks)
  const autoMediaTracks = compiledMediaTracks.filter((track) => track.property === "media-playback")
  let groups = schedule.groups.map((group): PptxCompiledTriggerGroup => {
    const effects = group.effects.filter((effect) => compilableNodeIds.has(effect.nodeId))
    const durationMs = effects.some((effect) => effect.endMs === "indefinite")
      ? "indefinite" as const
      : Math.max(0, ...effects.map((effect) => effect.endMs as number))
    return Object.freeze({
      key: group.key,
      clickBoundary: group.clickBoundary,
      durationMs,
      tracks: compileFragments(effects.flatMap((effect) => fragmentsForEffect(
        effect,
        frozenInitialState,
        slide.nodes[effect.nodeId],
        sourceState,
      ))),
      effects: Object.freeze(effects),
    })
  })
  if (compiledMediaTracks.length) {
    if (groups.some((group) => group.key === "auto")) {
      groups = groups.map((group) => group.key === "auto"
        ? Object.freeze({ ...group, tracks: Object.freeze([...group.tracks, ...compiledMediaTracks]) })
        : group)
    } else if (autoMediaTracks.length) {
      // 自动播放的媒体需要时间零就生效，因此建立只携带媒体轨道的自动组。
      groups = [Object.freeze({
        key: "auto" as const,
        clickBoundary: 0,
        durationMs: 0,
        tracks: Object.freeze(compiledMediaTracks),
        effects: Object.freeze([]),
      }), ...groups]
    }
  }
  return Object.freeze({
    slideIndex: slide.index,
    schedule,
    groups: Object.freeze(groups),
    initialState: frozenInitialState,
  })
}

function segmentProgress(segment: PptxCompiledTrackSegment, positionMs: number): number {
  if (segment.endMs <= segment.startMs) return 1
  let progress = (positionMs - segment.startMs) / (segment.endMs - segment.startMs)
  const easing = segment.easing?.match(/^pptx\(([^,]+),([^)]+)\)$/u)
  if (easing) {
    let acceleration = Math.max(0, Number(easing[1]) || 0)
    let deceleration = Math.max(0, Number(easing[2]) || 0)
    const total = acceleration + deceleration
    if (total > 1) {
      acceleration /= total
      deceleration /= total
    }
    const peak = 1 / (1 - (acceleration + deceleration) / 2)
    if (acceleration > 0 && progress < acceleration) {
      progress = peak * progress * progress / (2 * acceleration)
    } else if (deceleration > 0 && progress > 1 - deceleration) {
      progress = 1 - peak * (1 - progress) * (1 - progress) / (2 * deceleration)
    } else progress = peak * (progress - acceleration / 2)
  }
  return Math.max(0, Math.min(1, progress))
}

function evaluateSegment(segment: PptxCompiledTrackSegment, positionMs: number): PptxTrackValue {
  if (positionMs >= segment.endMs) return segment.to
  const progress = segmentProgress(segment, positionMs)
  if (typeof segment.from === "number" && typeof segment.to === "number") {
    return segment.from + (segment.to - segment.from) * progress
  }
  if (typeof segment.from === "string" && typeof segment.to === "string") {
    const fromInset = [...segment.from.matchAll(/-?\d+(?:\.\d+)?/gu)].map((match) => Number(match[0]))
    const toInset = [...segment.to.matchAll(/-?\d+(?:\.\d+)?/gu)].map((match) => Number(match[0]))
    if (segment.from.startsWith("inset(") && segment.to.startsWith("inset(") && fromInset.length === 4 && toInset.length === 4) {
      return `inset(${fromInset.map((value, item) => `${value + (toInset[item] - value) * progress}%`).join(" ")})`
    }
  }
  return progress >= 1 ? segment.to : segment.from
}

/**
 * Transform channels are layered on top of the element's static transform, so behaviours
 * that run on the same channel at the same time add up (PowerPoint's additive behaviour)
 * instead of silently dropping the earlier one. Repeats of a single behaviour still
 * replace each other, and every other property stays an absolute visual state that the
 * last writer settles.
 */
const additiveProperties: ReadonlySet<PptxTrackProperty> = new Set<PptxTrackProperty>([
  "translate-x",
  "translate-y",
  "rotate",
  "scale-x",
  "scale-y",
])

function laterSegment(
  candidate: PptxCompiledTrackSegment,
  current: PptxCompiledTrackSegment | undefined,
): boolean {
  return !current
    || candidate.startMs > current.startMs
    || candidate.startMs === current.startMs && candidate.order > current.order
}

interface BehaviourSegments {
  /** Rest value the behaviour started from, which its repeats and reverse passes return to. */
  first: PptxCompiledTrackSegment
  /** Segment that currently drives the behaviour. */
  latest: PptxCompiledTrackSegment
}

function evaluateAdditiveTrack(
  initialValue: PptxTrackValue,
  segments: readonly PptxCompiledTrackSegment[],
  positionMs: number,
): PptxTrackValue {
  if (segments.length === 1) {
    const segment = segments[0]
    if (positionMs < segment.startMs) return initialValue
    if (positionMs >= segment.endMs && segment.fill === "remove") return initialValue
    const value = evaluateSegment(segment, positionMs)
    if (
      typeof value !== "number"
      || typeof segment.from !== "number"
      || typeof initialValue !== "number"
    ) return value
    return initialValue + value - segment.from
  }
  const behaviours = new Map<string, BehaviourSegments>()
  for (const segment of segments) {
    if (positionMs < segment.startMs) continue
    if (positionMs >= segment.endMs && segment.fill === "remove") continue
    const entry = behaviours.get(segment.sourceNodeId)
    if (!entry) {
      behaviours.set(segment.sourceNodeId, { first: segment, latest: segment })
      continue
    }
    if (segment.startMs < entry.first.startMs) entry.first = segment
    if (laterSegment(segment, entry.latest)) entry.latest = segment
  }
  let latest: PptxCompiledTrackSegment | undefined
  for (const entry of behaviours.values()) {
    if (laterSegment(entry.latest, latest)) latest = entry.latest
  }
  if (!latest) return initialValue
  if (typeof initialValue !== "number") return evaluateSegment(latest, positionMs)
  let total = initialValue
  for (const entry of behaviours.values()) {
    const value = evaluateSegment(entry.latest, positionMs)
    if (typeof value !== "number" || typeof entry.first.from !== "number") {
      return evaluateSegment(latest, positionMs)
    }
    total += value - entry.first.from
  }
  return total
}

function evaluateCompiledTrack(
  property: PptxTrackProperty,
  initialValue: PptxTrackValue,
  segments: readonly PptxCompiledTrackSegment[],
  positionMs: number,
): PptxTrackValue {
  if (additiveProperties.has(property)) return evaluateAdditiveTrack(initialValue, segments, positionMs)
  let selected: PptxCompiledTrackSegment | undefined
  for (const segment of segments) {
    if (positionMs < segment.startMs) continue
    if (positionMs >= segment.endMs && segment.fill === "remove") continue
    if (laterSegment(segment, selected)) selected = segment
  }
  return selected ? evaluateSegment(selected, positionMs) : initialValue
}

export function evaluatePptxTrack(track: PptxPropertyTrack, positionMs: number): PptxTrackValue {
  if (track.repeatIndefinite && track.repeatPeriodMs && track.repeatStartMs !== undefined && positionMs >= track.repeatStartMs) {
    positionMs = track.repeatStartMs + (positionMs - track.repeatStartMs) % track.repeatPeriodMs
  }
  const segments = (track as Partial<PptxCompiledPropertyTrack>).segments
  if (segments?.length) {
    return evaluateCompiledTrack(track.property, track.initialValue, segments, positionMs)
  }
  if (track.keyframes.length === 0 || positionMs < track.keyframes[0].timeMs) return track.initialValue
  let previous = track.keyframes[0]
  for (let index = 1; index < track.keyframes.length; index += 1) {
    const next = track.keyframes[index]
    if (positionMs < next.timeMs) {
      if (typeof previous.value === "number" && typeof next.value === "number" && next.timeMs > previous.timeMs) {
        let progress = (positionMs - previous.timeMs) / (next.timeMs - previous.timeMs)
        const easing = previous.easing?.match(/^pptx\(([^,]+),([^)]+)\)$/u)
        if (easing) {
          let acceleration = Math.max(0, Number(easing[1]) || 0)
          let deceleration = Math.max(0, Number(easing[2]) || 0)
          const total = acceleration + deceleration
          if (total > 1) {
            acceleration /= total
            deceleration /= total
          }
          const peak = 1 / (1 - (acceleration + deceleration) / 2)
          if (acceleration > 0 && progress < acceleration) progress = peak * progress * progress / (2 * acceleration)
          else if (deceleration > 0 && progress > 1 - deceleration) {
            progress = 1 - peak * (1 - progress) * (1 - progress) / (2 * deceleration)
          } else progress = peak * (progress - acceleration / 2)
        }
        return previous.value + (next.value - previous.value) * Math.max(0, Math.min(1, progress))
      }
      if (typeof previous.value === "string" && typeof next.value === "string") {
        const fromInset = [...previous.value.matchAll(/-?\d+(?:\.\d+)?/gu)].map((match) => Number(match[0]))
        const toInset = [...next.value.matchAll(/-?\d+(?:\.\d+)?/gu)].map((match) => Number(match[0]))
        if (previous.value.startsWith("inset(") && next.value.startsWith("inset(") && fromInset.length === 4 && toInset.length === 4) {
          const progress = Math.max(0, Math.min(1, (positionMs - previous.timeMs) / (next.timeMs - previous.timeMs)))
          return `inset(${fromInset.map((value, item) => `${value + (toInset[item] - value) * progress}%`).join(" ")})`
        }
      }
      return previous.value
    }
    previous = next
  }
  return previous.value
}

export function evaluatePptxTriggerGroup(
  group: PptxCompiledTriggerGroup,
  positionMs: number,
): PptxObjectStateMap {
  const result: Record<string, PptxObjectPropertyState> = {}
  for (const track of group.tracks) {
    const target = track.target
      ?? (track.paragraphRange
        ? `${track.objectKey}#paragraph:${track.paragraphRange.start}-${track.paragraphRange.end}`
        : track.objectKey)
    const state = result[target] ?? {}
    state[track.property] = evaluatePptxTrack(track, positionMs)
    result[target] = state
  }
  return Object.freeze(Object.fromEntries(
    Object.entries(result).map(([key, value]) => [key, Object.freeze(value)]),
  ))
}

function mergeState(target: Record<string, PptxObjectPropertyState>, source: PptxObjectStateMap): void {
  for (const [objectKey, values] of Object.entries(source)) {
    target[objectKey] = { ...(target[objectKey] ?? {}), ...values }
  }
}

export function rebuildPptxStateAtBoundary(
  compiled: PptxCompiledSlide,
  clickBoundary: number,
): PptxObjectStateMap {
  const result: Record<string, PptxObjectPropertyState> = Object.fromEntries(
    Object.entries(compiled.initialState).map(([key, value]) => [key, { ...value }]),
  )
  for (const group of compiled.groups) {
    if (group.key !== "auto" && (!group.key.startsWith("click:") || group.clickBoundary > clickBoundary)) continue
    const position = group.durationMs === "indefinite" ? 0 : group.durationMs
    mergeState(result, evaluatePptxTriggerGroup(group, position))
  }
  return Object.freeze(Object.fromEntries(
    Object.entries(result).map(([key, value]) => [key, Object.freeze(value)]),
  ))
}
