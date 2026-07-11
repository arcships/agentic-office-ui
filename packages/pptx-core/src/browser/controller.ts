import type {
  PptxPlaybackController,
  PptxPlaybackControllerOptions,
  PptxPlaybackEvent,
  PptxPlaybackSnapshot,
  PptxPlaybackStatus,
} from "../browser-types"
import { PptxPlaybackError } from "../playback/errors"
import {
  compilePptxSlideTracks,
  evaluatePptxTriggerGroup,
  rebuildPptxStateAtBoundary,
  type PptxCompiledSlide,
  type PptxCompiledTriggerGroup,
  type PptxObjectPropertyState,
  type PptxObjectStateMap,
} from "../playback/track-compiler"
import type { PptxMediaItem, PptxPlaybackDocument } from "../playback/types"
import { createPptxClock } from "./clock"
import { createPptxMediaPlayer, type PptxMediaPlayer } from "./media-player"
import { applyPptxObjectState, capturePptxStaticState } from "./track-player"
import {
  capturePptxTransition,
  mountPptxTransition,
  type PptxTransitionHandle,
} from "./transition-player"

export interface PptxPlaybackControllerHost {
  root: HTMLElement
  document: PptxPlaybackDocument
  initialSlideIndex: number
  renderSlide(index: number): Promise<void>
  resolveMediaSource?(item: PptxMediaItem): Promise<string | undefined>
  releaseMediaSource?(item: PptxMediaItem, url: string): void
  onDispose?(): void
}

interface PendingPlaybackCommand {
  run(): Promise<void>
  resolve(): void
  reject(reason: unknown): void
}

function mergeState(base: PptxObjectStateMap, update: PptxObjectStateMap): PptxObjectStateMap {
  const result: Record<string, PptxObjectPropertyState> = Object.fromEntries(
    Object.entries(base).map(([key, value]) => [key, { ...value }]),
  )
  for (const [key, value] of Object.entries(update)) result[key] = { ...(result[key] ?? {}), ...value }
  return Object.freeze(Object.fromEntries(
    Object.entries(result).map(([key, value]) => [key, Object.freeze(value)]),
  ))
}

export function createPptxPlaybackController(
  host: PptxPlaybackControllerHost,
  options: PptxPlaybackControllerOptions = {},
): PptxPlaybackController {
  if (options.initialSlide !== undefined && Math.trunc(options.initialSlide) !== host.initialSlideIndex) {
    throw new PptxPlaybackError(
      "PLAYBACK_NOT_READY",
      "请在打开文档时指定初始页面，再创建播放控制器。",
      { slideIndex: options.initialSlide },
    )
  }
  const listeners = new Set<(event: PptxPlaybackEvent) => void>()
  const retainedEvents: PptxPlaybackEvent[] = []
  const clock = createPptxClock()
  let status: PptxPlaybackStatus = "ready"
  let slideIndex = Math.min(
    Math.max(Math.trunc(host.initialSlideIndex), 0),
    Math.max(host.document.slides.length - 1, 0),
  )
  let clickBoundary = 0
  let generation = 1
  let disposed = false
  let compiled: PptxCompiledSlide | null = null
  let stableState: PptxObjectStateMap = Object.freeze({})
  let activeGroup: PptxCompiledTriggerGroup | null = null
  let activeResolve: (() => void) | null = null
  let activePromise: Promise<void> | null = null
  let activeTransition: PptxTransitionHandle | null = null
  let transitionResolve: (() => void) | null = null
  let transitionPromise: Promise<void> | null = null
  let activeAdvance: { durationMs: number; generation: number } | null = null
  let pausedActivity: "group" | "transition" | "advance" | "media" | null = null
  let blockedActivity: "group" | "transition" | "advance" | "media" | null = null
  let navigationInProgress = false
  let pendingCommand: PendingPlaybackCommand | null = null
  let operationId = 0
  let playStarted = false
  let mediaPlayer: PptxMediaPlayer | null = null
  const completedGroups = new Set<string>()
  const startedEffects = new Set<string>()
  const endedEffects = new Set<string>()

  const snapshot = (): PptxPlaybackSnapshot => Object.freeze({
    status,
    slideIndex,
    clickBoundary,
    positionMs: clock.positionMs,
    activeNodeIds: Object.freeze(activeGroup?.effects
      .filter((effect) => startedEffects.has(effect.nodeId) && !endedEffects.has(effect.nodeId))
      .map((effect) => effect.nodeId) ?? []),
    blockedMediaIds: Object.freeze([...(mediaPlayer?.blockedMediaIds ?? [])]),
    generation,
  })
  const emit = (event: PptxPlaybackEvent) => {
    if (listeners.size === 0 && (event.type === "warning" || event.type === "error")) {
      retainedEvents.push(event)
    }
    for (const listener of [...listeners]) listener(event)
  }
  const emitState = () => emit({ type: "statechange", snapshot: snapshot() })
  const assertUsable = () => {
    if (disposed) throw new PptxPlaybackError("PLAYBACK_DISPOSED", "PPTX 播放控制器已经销毁。")
  }
  const queueCommand = (run: () => Promise<void>): Promise<void> => new Promise((resolve, reject) => {
    pendingCommand?.resolve()
    pendingCommand = { run, resolve, reject }
  })
  const flushPendingCommand = () => {
    const pending = pendingCommand
    pendingCommand = null
    if (!pending) return
    if (disposed) {
      pending.resolve()
      return
    }
    queueMicrotask(() => {
      if (disposed) pending.resolve()
      else void pending.run().then(pending.resolve, pending.reject)
    })
  }
  const applyStableState = () => {
    applyPptxObjectState(host.root, stableState)
    mediaPlayer?.apply(stableState)
  }

  const cancelAdvance = () => {
    if (!activeAdvance) return
    activeAdvance = null
    if (clock.running && !activeGroup && !activeTransition) clock.pause()
  }

  const prepareCurrentSlide = () => {
    const slide = host.document.slides[slideIndex]
    if (!slide) throw new PptxPlaybackError("PLAYBACK_NOT_READY", "当前页面没有播放模型。", { slideIndex })
    compiled = compilePptxSlideTracks(slide, capturePptxStaticState(host.root))
    const targetCounts = new Map<string, number>()
    for (const element of host.root.querySelectorAll<HTMLElement>("[data-pptx-object-key]")) {
      const key = element.dataset.pptxObjectKey
      if (key) targetCounts.set(key, (targetCounts.get(key) ?? 0) + 1)
    }
    const checkedTargets = new Set<string>()
    for (const effect of compiled.groups.flatMap((group) => group.effects)) {
      const objectKey = effect.effect.targetObjectKey
      if (!objectKey || checkedTargets.has(objectKey)) continue
      checkedTargets.add(objectKey)
      const count = targetCounts.get(objectKey) ?? 0
      if (count === 1) continue
      emit({
        type: "warning",
        warning: {
          code: count === 0 ? "TARGET_NOT_FOUND" : "TARGET_AMBIGUOUS",
          message: count === 0
            ? `动画目标 ${objectKey} 没有对应页面元素。`
            : `动画目标 ${objectKey} 对应了 ${count} 个页面元素。`,
          slideIndex,
          objectKey,
          recoverable: true,
        },
      })
    }
    mediaPlayer?.prepare(slide)
    stableState = compiled.initialState
    clickBoundary = 0
    completedGroups.clear()
    activeGroup = null
    activeAdvance = null
    pausedActivity = null
    blockedActivity = null
    startedEffects.clear()
    endedEffects.clear()
    playStarted = false
    clock.seek(0)
    applyStableState()
    for (const nodeId of compiled.schedule.unsupportedNodeIds) emit({
      type: "warning",
      warning: {
        code: "TIME_CONDITION_UNSUPPORTED",
        message: `时间节点 ${nodeId} 的触发条件尚未执行。`,
        slideIndex,
        sourceNodeId: nodeId,
        recoverable: true,
      },
    })
  }

  const updateEffectEvents = (positionMs: number, forceEnd = false) => {
    const group = activeGroup
    if (!group) return
    for (const effect of group.effects) {
      if (positionMs >= effect.startMs && !startedEffects.has(effect.nodeId)) {
        startedEffects.add(effect.nodeId)
        emit({
          type: "effectstart",
          slideIndex,
          nodeId: effect.nodeId,
          objectKey: effect.effect.targetObjectKey,
        })
      }
      if (
        startedEffects.has(effect.nodeId)
        && !endedEffects.has(effect.nodeId)
        && (forceEnd || effect.endMs !== "indefinite" && positionMs >= effect.endMs)
      ) {
        endedEffects.add(effect.nodeId)
        emit({
          type: "effectend",
          slideIndex,
          nodeId: effect.nodeId,
          objectKey: effect.effect.targetObjectKey,
        })
      }
    }
  }

  const finishActiveGroup = () => {
    const group = activeGroup
    if (!group) return
    const end = group.durationMs === "indefinite" ? clock.positionMs : group.durationMs
    updateEffectEvents(end, group.durationMs === "indefinite")
    activeGroup = null
    if (clock.running) clock.pause()
    stableState = mergeState(stableState, evaluatePptxTriggerGroup(group, end))
    completedGroups.add(group.key)
    if (group.key.startsWith("click:")) {
      clickBoundary = Math.max(clickBoundary, group.clickBoundary)
      emit({ type: "stepchange", slideIndex, boundary: clickBoundary })
    }
    applyStableState()
    status = "waiting"
    const resolve = activeResolve
    activeResolve = null
    activePromise = null
    emitState()
    resolve?.()
  }

  const finishActiveTransition = () => {
    const transition = activeTransition
    if (!transition) return
    activeTransition = null
    if (clock.running) clock.pause()
    transition.apply(transition.durationMs)
    transition.finish()
    applyStableState()
    status = "ready"
    const resolve = transitionResolve
    transitionResolve = null
    transitionPromise = null
    emitState()
    resolve?.()
  }

  const runTransition = async (
    handle: PptxTransitionHandle | null,
  ): Promise<void> => {
    if (!handle) return
    activeTransition = handle
    status = "transitioning"
    clock.seek(0)
    transitionPromise = new Promise<void>((resolve) => {
      transitionResolve = resolve
    })
    emitState()
    clock.play()
    await transitionPromise
  }

  const unsubscribeClock = clock.subscribe((positionMs) => {
    if (activeTransition) {
      activeTransition.apply(positionMs)
      if (positionMs >= activeTransition.durationMs) finishActiveTransition()
      else emitState()
      return
    }
    const group = activeGroup
    if (group) {
      const currentState = mergeState(stableState, evaluatePptxTriggerGroup(group, positionMs))
      applyPptxObjectState(host.root, currentState)
      mediaPlayer?.apply(currentState)
      updateEffectEvents(positionMs)
      if (group.durationMs !== "indefinite" && positionMs >= group.durationMs) finishActiveGroup()
      else emitState()
      return
    }
    const advance = activeAdvance
    if (advance) {
      if (positionMs < advance.durationMs) {
        emitState()
        return
      }
      activeAdvance = null
      if (clock.running) clock.pause()
      if (advance.generation !== generation || disposed) return
      const next = findNextVisibleSlide(1)
      if (next === undefined) {
        status = "ended"
        emitState()
        return
      }
      void goTo(next, "automatic").catch((reason) => emit({
        type: "error",
        error: reason instanceof PptxPlaybackError
          ? reason
          : new PptxPlaybackError("TRACK_COMPILE_FAILED", "自动换页失败。", { cause: reason, slideIndex }),
      }))
    }
  })

  const runGroup = (group: PptxCompiledTriggerGroup): Promise<void> => {
    assertUsable()
    if (activeGroup) finishActiveGroup()
    if (completedGroups.has(group.key)) return Promise.resolve()
    activeGroup = group
    startedEffects.clear()
    endedEffects.clear()
    status = "running"
    clock.seek(0)
    updateEffectEvents(0)
    activePromise = new Promise<void>((resolve) => {
      activeResolve = resolve
    })
    emitState()
    if (group.durationMs === 0) finishActiveGroup()
    else clock.play()
    return activePromise ?? Promise.resolve()
  }

  const findNextVisibleSlide = (direction: 1 | -1): number | undefined => {
    const skipHidden = options.skipHiddenSlides ?? true
    for (let index = slideIndex + direction; index >= 0 && index < host.document.slides.length; index += direction) {
      if (!skipHidden || !host.document.slides[index].hidden) return index
    }
    return undefined
  }

  const maybeArmAutoAdvance = () => {
    if (!playStarted || activeGroup || activeTransition || activeAdvance || disposed) return
    const slide = host.document.slides[slideIndex]
    const delay = slide?.transition?.advanceAfterMs
    if (delay === undefined) return
    const groups = compiled?.groups ?? []
    if (groups.some((group) => group.durationMs === "indefinite")) {
      emit({
        type: "warning",
        warning: {
          code: "AUTO_ADVANCE_BLOCKED_BY_INDEFINITE_NODE",
          message: "页面包含无限播放节点，已停止自动换页。",
          slideIndex,
          recoverable: true,
        },
      })
      return
    }
    if (groups.some((group) => !completedGroups.has(group.key))) return
    if (findNextVisibleSlide(1) === undefined) return
    activeAdvance = { durationMs: Math.max(0, delay), generation }
    status = "waiting"
    clock.seek(0)
    emitState()
    if (delay === 0) {
      const next = findNextVisibleSlide(1)
      activeAdvance = null
      if (next !== undefined) void goTo(next, "automatic")
    } else clock.play()
  }

  mediaPlayer = createPptxMediaPlayer(host.root, {
    onBlocked(mediaId) {
      blockedActivity = activeGroup ? "group" : activeTransition ? "transition" : activeAdvance ? "advance" : "media"
      if (clock.running) clock.pause()
      status = "blocked"
      emit({ type: "mediarequest", mediaId, reason: "autoplay-blocked" })
      emitState()
    },
    onError(mediaId, reason) {
      emit({
        type: "warning",
        warning: {
          code: "MEDIA_FAILED",
          message: `媒体 ${mediaId} 无法挂载或播放。`,
          slideIndex,
          recoverable: true,
        },
      })
      if (reason instanceof Error && reason.message) emit({
        type: "error",
        error: new PptxPlaybackError("MEDIA_FAILED", reason.message, { cause: reason, slideIndex }),
      })
    },
    onBookmark(_mediaId, bookmarkName) {
      const group = compiled?.groups.find((candidate) => candidate.key === `bookmark:${bookmarkName}`)
      if (group) void runGroup(group).then(maybeArmAutoAdvance)
    },
    onStateChange() {
      if (!disposed) emitState()
    },
    resolveSource: host.resolveMediaSource,
    releaseSource: host.releaseMediaSource,
  })

  const goTo = async (index: number, reason: string, includeHidden = false): Promise<void> => {
    assertUsable()
    if (!Number.isInteger(index) || index < 0 || index >= host.document.slides.length) {
      throw new PptxPlaybackError("PLAYBACK_NOT_READY", "页面编号超出范围。", { slideIndex: index })
    }
    if (host.document.slides[index].hidden && !includeHidden && (options.skipHiddenSlides ?? true)) {
      throw new PptxPlaybackError("PLAYBACK_NOT_READY", "顺序播放不能直接进入隐藏页。", { slideIndex: index })
    }
    if (navigationInProgress) return queueCommand(() => goTo(index, reason, includeHidden))
    navigationInProgress = true
    try {
      const commandId = ++operationId
      cancelAdvance()
      if (activeGroup) finishActiveGroup()
      if (activeTransition) finishActiveTransition()
      const from = slideIndex
      const capture = capturePptxTransition(host.root)
      status = "transitioning"
      generation += 1
      emitState()
      await host.renderSlide(index)
      assertUsable()
      if (commandId !== operationId) return
      slideIndex = index
      prepareCurrentSlide()
      const targetSlide = host.document.slides[index]
      const transitionKind = targetSlide.transition?.kind.toLowerCase()
      const executableTransition = transitionKind && (
        ["cut", "fade", "push", "wipe"].includes(transitionKind)
        || options.approximation === "safe"
      ) ? targetSlide.transition : undefined
      const transition = mountPptxTransition(
        host.root,
        capture,
        executableTransition,
        from === index - 1 ? targetSlide.morphFromPrevious : [],
      )
      await runTransition(transition)
      if (commandId !== operationId) return
      status = "ready"
      emit({ type: "slidechange", from, to: index, reason })
      emitState()
      if (options.autoplay !== false) await controller.play()
    } finally {
      navigationInProgress = false
      flushPendingCommand()
    }
  }

  const controller: PptxPlaybackController = {
    get snapshot() {
      return snapshot()
    },
    async next() {
      assertUsable()
      if (navigationInProgress) return queueCommand(() => controller.next())
      cancelAdvance()
      if (activeTransition) {
        finishActiveTransition()
        return
      }
      if (activeGroup) {
        finishActiveGroup()
        return
      }
      const group = compiled?.groups.find((candidate) => candidate.key === `click:${clickBoundary + 1}`)
      if (group) {
        await runGroup(group)
        maybeArmAutoAdvance()
        return
      }
      const next = findNextVisibleSlide(1)
      if (next === undefined) {
        status = "ended"
        emitState()
        return
      }
      await goTo(next, "next")
    },
    async activateObject(objectKey) {
      assertUsable()
      if (navigationInProgress) return false
      const group = compiled?.groups.find((candidate) => candidate.key === `shape:${objectKey}`)
      if (group) {
        await runGroup(group)
        return true
      }
      const action = host.document.slides[slideIndex]?.actions.find(
        (candidate) => candidate.sourceObjectKey === objectKey && candidate.trigger === "click",
      )
      if (!action) return false
      if (action.kind === "go-to-slide" && action.targetSlideIndex !== undefined) {
        emit({
          type: "action",
          action: { kind: "go-to-slide", slideIndex: action.targetSlideIndex, sourceObjectKey: objectKey },
        })
        await goTo(action.targetSlideIndex, "action", true)
      } else if (action.kind === "open-url" && action.url) {
        emit({ type: "action", action: { kind: "open-url", url: action.url, sourceObjectKey: objectKey } })
      } else if (action.kind === "mailto" && action.url) {
        emit({ type: "action", action: { kind: "mailto", url: action.url, sourceObjectKey: objectKey } })
      } else emit({
        type: "action",
        action: { kind: "unsupported", action: action.rawAction ?? "unknown", sourceObjectKey: objectKey },
      })
      return true
    },
    async previous() {
      assertUsable()
      if (navigationInProgress) return queueCommand(() => controller.previous())
      cancelAdvance()
      if (activeTransition) {
        finishActiveTransition()
        return
      }
      if (activeGroup) finishActiveGroup()
      if (clickBoundary > 0 && compiled) {
        clickBoundary -= 1
        stableState = rebuildPptxStateAtBoundary(compiled, clickBoundary)
        completedGroups.clear()
        for (const group of compiled.groups) {
          if (group.key === "auto" || group.key.startsWith("click:") && group.clickBoundary <= clickBoundary) {
            completedGroups.add(group.key)
          }
        }
        applyStableState()
        status = "waiting"
        emit({ type: "stepchange", slideIndex, boundary: clickBoundary })
        emitState()
        return
      }
      const previous = findNextVisibleSlide(-1)
      if (previous !== undefined) await goTo(previous, "previous")
    },
    async play() {
      assertUsable()
      playStarted = true
      if (status === "paused") {
        await controller.resume()
        return
      }
      const automatic = compiled?.groups.find((group) => group.key === "auto")
      if (automatic && !completedGroups.has("auto")) await runGroup(automatic)
      else {
        status = "waiting"
        emitState()
      }
      maybeArmAutoAdvance()
    },
    pause() {
      assertUsable()
      if (status === "paused") return
      pausedActivity = activeGroup ? "group" : activeTransition ? "transition" : activeAdvance ? "advance" : "media"
      if (clock.running) clock.pause()
      mediaPlayer?.pause()
      status = "paused"
      emitState()
    },
    async resume() {
      assertUsable()
      if (status !== "paused" || !pausedActivity) return
      const activity = pausedActivity
      pausedActivity = null
      status = activity === "transition" ? "transitioning" : activity === "group" ? "running" : "waiting"
      if (activity !== "media") clock.play()
      await mediaPlayer?.resume()
      emitState()
      if (activity === "group") await activePromise
      else if (activity === "transition") await transitionPromise
    },
    async reset() {
      assertUsable()
      if (navigationInProgress) return queueCommand(() => controller.reset())
      operationId += 1
      cancelAdvance()
      if (activeTransition) finishActiveTransition()
      if (activeGroup) finishActiveGroup()
      prepareCurrentSlide()
      status = "ready"
      emitState()
      if (options.autoplay !== false) await controller.play()
    },
    async goToSlide(index, goToOptions = {}) {
      await goTo(index, "direct", goToOptions.includeHidden === true)
    },
    async resumeBlockedMedia(mediaId) {
      assertUsable()
      await mediaPlayer?.resumeBlocked(mediaId)
      if (mediaPlayer?.blockedMediaIds.length === 0 && status === "blocked") {
        const activity = blockedActivity
        blockedActivity = null
        status = activity === "transition" ? "transitioning" : activity === "group" ? "running" : "waiting"
        if (activity === "group" || activity === "transition" || activity === "advance") clock.play()
        emitState()
      }
    },
    subscribe(listener) {
      assertUsable()
      listeners.add(listener)
      listener({ type: "statechange", snapshot: snapshot() })
      listener({ type: "capability", report: host.document.capability })
      for (const event of retainedEvents) listener(event)
      return () => listeners.delete(listener)
    },
    dispose() {
      if (disposed) return
      disposed = true
      operationId += 1
      pendingCommand?.resolve()
      pendingCommand = null
      cancelAdvance()
      if (activeTransition) finishActiveTransition()
      if (activeGroup) finishActiveGroup()
      status = "disposed"
      unsubscribeClock()
      clock.dispose()
      mediaPlayer?.dispose()
      mediaPlayer = null
      host.onDispose?.()
      activeResolve?.()
      activeResolve = null
      activePromise = null
      emitState()
      listeners.clear()
    },
  }

  prepareCurrentSlide()
  if (options.autoplay !== false) queueMicrotask(() => {
    if (!disposed) void controller.play()
  })
  return controller
}
