import type {
  PptxPlaybackEffect,
  PptxPlaybackSlide,
  PptxTimeCondition,
  PptxTimeNode,
  PptxTriggerEvent,
} from "./types"

export type PptxTriggerKey = "auto" | `click:${number}` | `shape:${string}` | `bookmark:${string}`

export interface PptxScheduledEffect {
  nodeId: string
  effect: PptxPlaybackEffect
  triggerKey: PptxTriggerKey
  triggerEvent: PptxTriggerEvent
  clickBoundary: number
  startMs: number
  endMs: number | "indefinite"
  order: number
}

export interface PptxTriggerSchedule {
  key: PptxTriggerKey
  clickBoundary: number
  durationMs: number | "indefinite"
  effects: readonly PptxScheduledEffect[]
}

export interface PptxSlideSchedule {
  groups: readonly PptxTriggerSchedule[]
  clickBoundaryCount: number
  unsupportedNodeIds: readonly string[]
}

interface VisitContext {
  triggerKey?: PptxTriggerKey
  triggerEvent?: PptxTriggerEvent
  clickBoundary: number
}

function startCondition(node: PptxTimeNode): PptxTimeCondition | undefined {
  return node.conditions.find((condition) => condition.source === "start" && condition.event !== "delay")
    ?? node.conditions.find((condition) => condition.source === "start")
}

function nodeTypeEvent(node: PptxTimeNode): PptxTriggerEvent | undefined {
  if (node.nodeType === "clickEffect") return "on-click"
  if (node.nodeType === "withEffect") return "with-previous"
  if (node.nodeType === "afterEffect") return "after-previous"
  return undefined
}

function effectiveDuration(node: PptxTimeNode): number | "indefinite" {
  if (node.durationMs === "indefinite" || node.repeatCount === "indefinite") return "indefinite"
  const repeat = typeof node.repeatCount === "number" ? node.repeatCount : 1
  return node.durationMs * Math.max(repeat, 1) * (node.autoReverse ? 2 : 1)
}

function triggerForCondition(
  condition: PptxTimeCondition | undefined,
  event: PptxTriggerEvent,
  clickBoundary: number,
): PptxTriggerKey | undefined {
  if (event === "on-click") return `click:${clickBoundary}`
  if (event === "on-shape-click" && condition?.targetObjectKey) return `shape:${condition.targetObjectKey}`
  if (event === "on-media-bookmark" && condition?.bookmarkName) return `bookmark:${condition.bookmarkName}`
  return undefined
}

function executableTrigger(event: PptxTriggerEvent): boolean {
  return [
    "delay",
    "on-click",
    "with-previous",
    "after-previous",
    "on-shape-click",
    "on-begin",
    "on-end",
    "on-media-bookmark",
  ].includes(event)
}

export function compilePptxSlideSchedule(slide: PptxPlaybackSlide): PptxSlideSchedule {
  if (!slide.rootNodeId || !slide.nodes[slide.rootNodeId]) {
    return Object.freeze({ groups: Object.freeze([]), clickBoundaryCount: 0, unsupportedNodeIds: Object.freeze([]) })
  }
  const scheduled: PptxScheduledEffect[] = []
  const unsupported = new Set<string>()
  const previousByTrigger = new Map<PptxTriggerKey, PptxScheduledEffect>()
  const scheduledByNode = new Map<string, PptxScheduledEffect>()
  let clickBoundary = 0
  let order = 0
  let lastTrigger: PptxTriggerKey = "auto"

  const visit = (nodeId: string, inherited: VisitContext): void => {
    const node = slide.nodes[nodeId]
    if (!node) {
      unsupported.add(nodeId)
      return
    }
    const condition = startCondition(node)
    const explicitEvent = condition?.event ?? nodeTypeEvent(node)
    const event = explicitEvent ?? inherited.triggerEvent ?? "delay"
    let context = inherited
    if (explicitEvent && !executableTrigger(explicitEvent) && !node.effect) {
      unsupported.add(node.id)
      context = { triggerEvent: explicitEvent, clickBoundary: inherited.clickBoundary }
    } else if (explicitEvent === "on-click" && !node.effect) {
      clickBoundary += 1
      context = { triggerKey: `click:${clickBoundary}`, triggerEvent: event, clickBoundary }
    } else if ((explicitEvent === "on-shape-click" || explicitEvent === "on-media-bookmark") && !node.effect) {
      const key = triggerForCondition(condition, event, clickBoundary)
      if (key) context = { triggerKey: key, triggerEvent: event, clickBoundary }
      else unsupported.add(node.id)
    }

    if (node.effect) {
      let boundary = context.clickBoundary
      let triggerKey = context.triggerKey
      if (event === "on-click" && context.triggerEvent !== "on-click") {
        clickBoundary += 1
        boundary = clickBoundary
        triggerKey = `click:${boundary}`
      } else if (event === "on-shape-click" || event === "on-media-bookmark") {
        triggerKey = triggerForCondition(condition, event, boundary) ?? context.triggerKey
      } else if (event === "with-previous" || event === "after-previous" || event === "on-begin" || event === "on-end") {
        triggerKey ??= lastTrigger
      } else triggerKey ??= "auto"

      if (!triggerKey || !executableTrigger(event)) {
        unsupported.add(node.id)
      } else {
        const previous = previousByTrigger.get(triggerKey)
        const delay = Math.max(0, condition?.delayMs ?? node.delayMs)
        let startMs = delay
        if (event === "with-previous" && previous) startMs = previous.startMs + delay
        else if (event === "after-previous" && previous) {
          startMs = previous.endMs === "indefinite" ? previous.startMs : previous.endMs + delay
        } else if ((event === "on-begin" || event === "on-end") && condition?.targetNodeId) {
          const referenced = scheduledByNode.get(condition.targetNodeId)
          if (referenced) {
            const referenceTime = event === "on-begin" || referenced.endMs === "indefinite"
              ? referenced.startMs
              : referenced.endMs
            startMs = referenceTime + delay
          } else unsupported.add(node.id)
        } else if (event === "delay" && node.container === "sequence" && previous) {
          startMs = previous.endMs === "indefinite" ? previous.startMs : previous.endMs + delay
        }
        const duration = effectiveDuration(node)
        const item: PptxScheduledEffect = Object.freeze({
          nodeId: node.id,
          effect: node.effect,
          triggerKey,
          triggerEvent: event,
          clickBoundary: boundary,
          startMs,
          endMs: duration === "indefinite" ? "indefinite" : startMs + duration,
          order: order++,
        })
        scheduled.push(item)
        previousByTrigger.set(triggerKey, item)
        scheduledByNode.set(node.id, item)
        lastTrigger = triggerKey
      }
    }
    for (const childId of node.childIds) visit(childId, context)
  }

  visit(slide.rootNodeId, { triggerKey: "auto", triggerEvent: "delay", clickBoundary: 0 })
  const groupMap = new Map<PptxTriggerKey, PptxScheduledEffect[]>()
  for (const effect of scheduled) {
    const group = groupMap.get(effect.triggerKey) ?? []
    group.push(effect)
    groupMap.set(effect.triggerKey, group)
  }
  const groups = [...groupMap].map(([key, effects]): PptxTriggerSchedule => {
    const indefinite = effects.some((effect) => effect.endMs === "indefinite")
    return Object.freeze({
      key,
      clickBoundary: effects[0]?.clickBoundary ?? 0,
      durationMs: indefinite
        ? "indefinite"
        : Math.max(0, ...effects.map((effect) => effect.endMs as number)),
      effects: Object.freeze([...effects].sort((left, right) => left.startMs - right.startMs || left.order - right.order)),
    })
  }).sort((left, right) => {
    if (left.key === "auto") return -1
    if (right.key === "auto") return 1
    return left.clickBoundary - right.clickBoundary
  })
  return Object.freeze({
    groups: Object.freeze(groups),
    clickBoundaryCount: clickBoundary,
    unsupportedNodeIds: Object.freeze([...unsupported]),
  })
}
