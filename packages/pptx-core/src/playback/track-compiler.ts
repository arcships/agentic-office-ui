import { compilePptxSlideSchedule, type PptxScheduledEffect, type PptxSlideSchedule, type PptxTriggerKey } from "./time-tree"
import type {
  PptxPlaybackEffect,
  PptxPlaybackSlide,
  PptxPropertyTrack,
  PptxTrackKeyframe,
  PptxTrackProperty,
  PptxTimeNode,
} from "./types"

export type PptxTrackValue = string | number | boolean
export type PptxObjectPropertyState = Partial<Record<PptxTrackProperty, PptxTrackValue>>
export type PptxObjectStateMap = Readonly<Record<string, Readonly<PptxObjectPropertyState>>>

export interface PptxCompiledPropertyTrack extends PptxPropertyTrack {
  paragraphRange?: { start: number; end: number }
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
  paragraphRange?: { start: number; end: number }
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
  objectKey: string,
  property: PptxTrackProperty,
  paragraphRange?: { start: number; end: number },
): PptxTrackValue {
  const target = paragraphRange
    ? `${objectKey}#paragraph:${paragraphRange.start}-${paragraphRange.end}`
    : objectKey
  return initialState[target]?.[property] ?? defaultValues[property]
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
): TrackFragment[] {
  const effect = scheduled.effect
  const objectKey = effect.targetObjectKey
  if (!objectKey) return []
  const baseDuration = typeof node.durationMs === "number" ? node.durationMs : 0
  const baseEnd = scheduled.startMs + baseDuration
  const fragment = (
    property: PptxTrackProperty,
    from: PptxTrackValue,
    to: PptxTrackValue,
    startMs = scheduled.startMs,
    endMs = baseEnd,
  ): TrackFragment => ({
    objectKey,
    paragraphRange: effect.paragraphRange,
    property,
    initialValue: initialValue(initialState, objectKey, property, effect.paragraphRange),
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
      fragments = [
        fragment("opacity", 1, 0),
        fragment("display", true, false, baseEnd, baseEnd),
      ]
      break
    case "scale":
      {
        const fallback = effect.presetClass === "emph" ? 1 : 0
        const fromX = numberValue(effect, "scaleFromX", fallback)
        const fromY = numberValue(effect, "scaleFromY", fallback)
        fragments = [
          fragment("scale-x", fromX, numberValue(effect, "scaleToX", fromX * numberValue(effect, "scaleByX", 1))),
          fragment("scale-y", fromY, numberValue(effect, "scaleToY", fromY * numberValue(effect, "scaleByY", 1))),
        ]
      }
      break
    case "emphasis": {
      const property = effect.values.emphasisProperty
      if (property === "opacity") {
        const from = numberValue(effect, "emphasisFrom", initialValue(initialState, objectKey, "opacity", effect.paragraphRange) as number)
        fragments = [fragment("opacity", from, numberValue(effect, "emphasisTo", from))]
      } else fragments = []
      break
    }
    case "rotate":
      fragments = [fragment("rotate", 0, numberValue(effect, "rotationBy", 0))]
      break
    case "motion-path": {
      const end = motionEnd(effect)
      fragments = end ? [fragment("translate-x", 0, end.x), fragment("translate-y", 0, end.y)] : []
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
    case "media-command":
      fragments = [fragment("media-playback", "stopped", effect.command ?? "play")]
      break
    default:
      fragments = []
  }
  return expandRepeats(fragments, node)
}

function trackKey(fragment: TrackFragment): string {
  const range = fragment.paragraphRange
  return `${fragment.objectKey}|${range ? `${range.start}:${range.end}` : "object"}|${fragment.property}`
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
      paragraphRange: first.paragraphRange,
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
  for (const scheduled of schedule.groups.flatMap((group) => group.effects)) {
    const effect = scheduled.effect
    const objectKey = effect.targetObjectKey
    if (!objectKey || effect.presetClass !== "entr" || !compilableNodeIds.has(scheduled.nodeId)) continue
    const target = effect.paragraphRange
      ? `${objectKey}#paragraph:${effect.paragraphRange.start}-${effect.paragraphRange.end}`
      : objectKey
    const state = preparedState[target] ?? {}
    if (effect.kind === "appear") state.display = false
    else if (effect.kind === "fade-in") {
      state.display = false
      state.opacity = 0
    } else if (effect.kind === "wipe") state["clip-path"] = effect.filter?.includes("right")
      ? "inset(0 0 0 100%)"
      : "inset(0 100% 0 0)"
    else if (effect.kind === "scale") {
      state["scale-x"] = numberValue(effect, "scaleFromX", 0)
      state["scale-y"] = numberValue(effect, "scaleFromY", 0)
    }
    preparedState[target] = state
  }
  const frozenInitialState = Object.freeze(Object.fromEntries(
    Object.entries(preparedState).map(([key, value]) => [key, Object.freeze(value)]),
  ))
  const groups = schedule.groups.map((group): PptxCompiledTriggerGroup => {
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
      ))),
      effects: Object.freeze(effects),
    })
  })
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

export function evaluatePptxTrack(track: PptxPropertyTrack, positionMs: number): PptxTrackValue {
  if (track.repeatIndefinite && track.repeatPeriodMs && track.repeatStartMs !== undefined && positionMs >= track.repeatStartMs) {
    positionMs = track.repeatStartMs + (positionMs - track.repeatStartMs) % track.repeatPeriodMs
  }
  const segments = (track as Partial<PptxCompiledPropertyTrack>).segments
  if (segments?.length) {
    const selected = segments.filter((segment) => {
      if (positionMs < segment.startMs) return false
      if (positionMs < segment.endMs) return true
      return segment.fill !== "remove"
    }).sort((left, right) => left.startMs - right.startMs || left.order - right.order).at(-1)
    return selected ? evaluateSegment(selected, positionMs) : track.initialValue
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
    const target = track.paragraphRange
      ? `${track.objectKey}#paragraph:${track.paragraphRange.start}-${track.paragraphRange.end}`
      : track.objectKey
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
