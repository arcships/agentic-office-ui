#!/usr/bin/env python3

import argparse
import json
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import sync_playwright


def main() -> None:
    parser = argparse.ArgumentParser(description="验证真实 PPTX 能生成正式播放文档。")
    parser.add_argument("--app-url", default="http://127.0.0.1:5000")
    parser.add_argument("--sample-url", required=True)
    parser.add_argument("--initial-slide", type=int, default=0)
    parser.add_argument("--steps", type=int, default=0)
    parser.add_argument("--pause-check", action="store_true")
    parser.add_argument("--previous-check", action="store_true")
    parser.add_argument("--transition-check", action="store_true")
    parser.add_argument("--approximation", choices=("off", "safe"), default="off")
    parser.add_argument("--auto-advance-check", action="store_true")
    parser.add_argument("--media-check", action="store_true")
    parser.add_argument("--browser", choices=("chromium", "firefox", "webkit"), default="chromium")
    parser.add_argument("--performance-check", action="store_true")
    parser.add_argument("--resource-check", action="store_true")
    parser.add_argument("--quiet", action="store_true")
    parser.add_argument("--repeat-check", action="store_true")
    parser.add_argument("--indefinite-check", action="store_true")
    parser.add_argument("--hidden-check", action="store_true")
    parser.add_argument("--action-check", action="store_true")
    parser.add_argument("--overlap-check", action="store_true")
    parser.add_argument("--paragraph-check", action="store_true")
    parser.add_argument("--shape-trigger-check", action="store_true")
    parser.add_argument("--rapid-navigation-check", action="store_true")
    parser.add_argument("--effect-timing-check", action="store_true")
    parser.add_argument("--duplicate-target-check", action="store_true")
    parser.add_argument("--reset-check", action="store_true")
    parser.add_argument("--media-block-check", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    console_errors: list[str] = []
    page_errors: list[str] = []
    with sync_playwright() as playwright:
        browser_type = getattr(playwright, args.browser)
        launch_options = {"headless": True}
        chromium_args = []
        if args.media_check and args.browser == "chromium":
            chromium_args.append("--autoplay-policy=no-user-gesture-required")
        if chromium_args:
            launch_options["args"] = chromium_args
        browser = browser_type.launch(**launch_options)
        page = browser.new_page()
        page.on("console", lambda message: console_errors.append(message.text) if message.type == "error" else None)
        page.on("pageerror", lambda error: page_errors.append(str(error)))
        page.goto(args.sample_url.rsplit("/", 1)[0] + "/")
        page.wait_for_load_state("networkidle")
        workspace = Path(__file__).resolve().parents[2]
        module_path = workspace / "packages/pptx-core/src/browser.ts"
        core_path = workspace / "packages/pptx-core/src/index.ts"
        module_url = args.app_url + "/@fs" + quote(str(module_path), safe="/")
        core_url = args.app_url + "/@fs" + quote(str(core_path), safe="/")
        result = page.evaluate(
            """
            async ({ moduleUrl, coreUrl, sampleUrl, initialSlide, steps, pauseCheck, previousCheck, transitionCheck, approximation, autoAdvanceCheck, mediaCheck, resourceCheck, repeatCheck, indefiniteCheck, hiddenCheck, actionCheck, overlapCheck, paragraphCheck, shapeTriggerCheck, rapidNavigationCheck, effectTimingCheck, duplicateTargetCheck, resetCheck, mediaBlockCheck }) => {
              const api = await import(moduleUrl)
              const core = await import(coreUrl)
              const response = await fetch(sampleUrl)
              if (!response.ok) throw new Error(`sample fetch failed: ${response.status}`)
              const root = document.createElement("div")
              root.style.width = "1280px"
              root.style.height = "720px"
              document.body.replaceChildren(root)
              const session = api.createPptxDocumentSession(root, {
                lazyMedia: true,
                lazySlides: true,
                approximation,
              })
              try {
                const openStartedAt = performance.now()
                const preview = await session.open(await response.arrayBuffer(), { initialSlide })
                const openMs = performance.now() - openStartedAt
                const playback = session.playbackDocument
                if (!playback) throw new Error("playback document missing")
                const schedule = core.compilePptxSlideSchedule(playback.slides[initialSlide])
                let playbackResult = {}
                if (steps > 0 || autoAdvanceCheck || mediaCheck || repeatCheck || indefiniteCheck || hiddenCheck || actionCheck || overlapCheck || paragraphCheck || shapeTriggerCheck || rapidNavigationCheck || effectTimingCheck || duplicateTargetCheck || resetCheck || mediaBlockCheck) {
                  const events = []
                  const eventRecords = []
                  let duplicateTargetKey
                  let duplicateInitialStyles
                  if (duplicateTargetCheck) {
                    duplicateTargetKey = schedule.groups.flatMap((group) => group.effects)
                      .map((effect) => effect.effect.targetObjectKey)
                      .find(Boolean)
                    const target = [...root.querySelectorAll("[data-pptx-object-key]")]
                      .find((element) => element.dataset.pptxObjectKey === duplicateTargetKey)
                    if (!target) throw new Error("controlled target missing")
                    target.after(target.cloneNode(true))
                    duplicateInitialStyles = [...root.querySelectorAll("[data-pptx-object-key]")]
                      .filter((element) => element.dataset.pptxObjectKey === duplicateTargetKey)
                      .map((element) => ({ visibility: element.style.visibility, opacity: element.style.opacity }))
                  }
                  const compileStartedAt = performance.now()
                  const originalMediaPlay = HTMLMediaElement.prototype.play
                  let rejectedMediaPlay = false
                  if (mediaBlockCheck) HTMLMediaElement.prototype.play = function() {
                    if (!rejectedMediaPlay) {
                      rejectedMediaPlay = true
                      return Promise.reject(new DOMException("autoplay blocked", "NotAllowedError"))
                    }
                    return Promise.resolve()
                  }
                  const controller = session.createPlaybackController({ autoplay: false, initialSlide, approximation })
                  const controllerCompileMs = performance.now() - compileStartedAt
                  const unsubscribe = controller.subscribe((event) => {
                    events.push(event.type)
                    eventRecords.push({ ...event, observedAtMs: performance.now() })
                  })
                  try {
                    if (mediaCheck) {
                      for (const media of root.querySelectorAll("video,audio")) media.muted = true
                    }
                    const firstTargetKey = schedule.groups
                      .flatMap((group) => group.effects)
                      .map((effect) => effect.effect.targetObjectKey)
                      .find(Boolean)
                    const firstTarget = [...root.querySelectorAll("[data-pptx-object-key]")]
                      .find((element) => element.dataset.pptxObjectKey === firstTargetKey)
                    const styleState = () => firstTarget ? {
                      visibility: firstTarget.style.visibility,
                      opacity: firstTarget.style.opacity,
                      scale: firstTarget.style.scale,
                    } : null
                    const initialAnimatedState = styleState()
                    let repeatScale
                    let repeatTranslate
                    let repeatOpacity
                    let repeatRotate
                    let repeatClipPath
                    let repeatPauseDeltaMs
                    let repeatFinalScale
                    let indefiniteScaleBefore
                    let indefiniteScaleAfter
                    let indefinitePauseDeltaMs
                    let indefiniteFinalStatus
                    let indefiniteWarningCount
                    let indefiniteFinalSlideIndex
                    let hiddenDirectRejected
                    let hiddenDirectSlideIndex
                    let hiddenSkippedSlideIndex
                    let externalActionHandled
                    let externalActionSlideIndex
                    let internalActionHandled
                    let internalActionSlideIndex
                    let actionRequests
                    let windowOpenCalls
                    let overlapInitialState
                    let overlapMidState
                    let overlapFinalState
                    let paragraphStates
                    let shapeTriggerHandled
                    let shapeTriggerKey
                    let rapidNavigationSlideIndex
                    let rapidNavigationChanges
                    let effectTiming
                    let duplicateFinalStyles
                    let duplicateWarnings
                    let resetInitialState
                    let resetAfterStepState
                    let resetState
                    let resetReplayState
                    let blockedStatus
                    let blockedMediaIds
                    let resumedBlockedStatus
                    if (mediaBlockCheck) {
                      const running = controller.play()
                      const deadline = performance.now() + 3000
                      while (controller.snapshot.status !== "blocked" && performance.now() < deadline) {
                        await new Promise((resolve) => setTimeout(resolve, 20))
                      }
                      blockedStatus = controller.snapshot.status
                      blockedMediaIds = [...controller.snapshot.blockedMediaIds]
                      await controller.resumeBlockedMedia()
                      resumedBlockedStatus = controller.snapshot.status
                      await Promise.race([
                        running,
                        new Promise((_, reject) => setTimeout(() => reject(new Error("blocked media group timeout")), 3000)),
                      ])
                    } else if (resetCheck) {
                      resetInitialState = styleState()
                      await controller.next()
                      resetAfterStepState = styleState()
                      await controller.reset()
                      resetState = styleState()
                      await controller.next()
                      resetReplayState = styleState()
                    } else if (duplicateTargetCheck) {
                      await controller.next()
                      duplicateFinalStyles = [...root.querySelectorAll("[data-pptx-object-key]")]
                        .filter((element) => element.dataset.pptxObjectKey === duplicateTargetKey)
                        .map((element) => ({ visibility: element.style.visibility, opacity: element.style.opacity }))
                      duplicateWarnings = eventRecords
                        .filter((event) => event.type === "warning")
                        .map((event) => event.warning.code)
                    } else if (rapidNavigationCheck) {
                      await Promise.all([
                        controller.goToSlide(initialSlide + 1),
                        controller.goToSlide(initialSlide + 2),
                        controller.goToSlide(initialSlide + 3),
                      ])
                      rapidNavigationSlideIndex = controller.snapshot.slideIndex
                      rapidNavigationChanges = eventRecords
                        .filter((event) => event.type === "slidechange")
                        .map((event) => event.to)
                    } else if (effectTimingCheck) {
                      const group = schedule.groups.find((candidate) => candidate.key.startsWith("click:"))
                      if (!group) throw new Error("controlled click group missing")
                      const startedAt = performance.now()
                      await controller.next()
                      effectTiming = group.effects.map((effect) => {
                        const started = eventRecords.find((event) => event.type === "effectstart" && event.nodeId === effect.nodeId)
                        const ended = eventRecords.find((event) => event.type === "effectend" && event.nodeId === effect.nodeId)
                        return {
                          nodeId: effect.nodeId,
                          expectedStartMs: effect.startMs,
                          expectedEndMs: effect.endMs,
                          observedStartMs: started ? started.observedAtMs - startedAt : null,
                          observedEndMs: ended ? ended.observedAtMs - startedAt : null,
                        }
                      })
                    } else if (shapeTriggerCheck) {
                      const shapeGroup = schedule.groups.find((group) => group.key.startsWith("shape:"))
                      if (!shapeGroup) throw new Error("controlled shape trigger missing")
                      shapeTriggerKey = shapeGroup.key.slice("shape:".length)
                      shapeTriggerHandled = await controller.activateObject(shapeTriggerKey)
                    } else if (overlapCheck) {
                      overlapInitialState = styleState()
                      const running = controller.play()
                      await new Promise((resolve) => setTimeout(resolve, 250))
                      overlapMidState = styleState()
                      await running
                      overlapFinalState = styleState()
                    } else if (paragraphCheck) {
                      const readParagraphs = () => [...root.querySelectorAll("[data-pptx-paragraph-index]")].map((element) => ({
                        index: Number(element.dataset.pptxParagraphIndex),
                        visibility: element.style.visibility,
                        opacity: element.style.opacity,
                      }))
                      paragraphStates = [readParagraphs()]
                      for (let index = 0; index < 3; index += 1) {
                        await controller.next()
                        paragraphStates.push(readParagraphs())
                      }
                    } else if (actionCheck) {
                      const slideActions = playback.slides[initialSlide].actions
                      const external = slideActions.find((action) => action.kind === "open-url")
                      const internal = slideActions.find((action) => action.kind === "go-to-slide")
                      if (!external || !internal) throw new Error("controlled actions missing")
                      windowOpenCalls = 0
                      const originalOpen = window.open
                      window.open = () => { windowOpenCalls += 1; return null }
                      try {
                        externalActionHandled = await controller.activateObject(external.sourceObjectKey)
                        externalActionSlideIndex = controller.snapshot.slideIndex
                        internalActionHandled = await controller.activateObject(internal.sourceObjectKey)
                        internalActionSlideIndex = controller.snapshot.slideIndex
                      } finally {
                        window.open = originalOpen
                      }
                      actionRequests = eventRecords
                        .filter((event) => event.type === "action")
                        .map((event) => event.action.kind)
                    } else if (hiddenCheck) {
                      hiddenDirectRejected = false
                      try {
                        await controller.goToSlide(initialSlide + 1)
                      } catch {
                        hiddenDirectRejected = true
                      }
                      await controller.goToSlide(initialSlide + 1, { includeHidden: true })
                      hiddenDirectSlideIndex = controller.snapshot.slideIndex
                      await controller.goToSlide(initialSlide, { includeHidden: true })
                      await controller.next()
                      hiddenSkippedSlideIndex = controller.snapshot.slideIndex
                    } else if (indefiniteCheck) {
                      const scaleKey = schedule.groups.flatMap((group) => group.effects)
                        .find((effect) => effect.effect.kind === "scale")?.effect.targetObjectKey
                      const scaleElement = [...root.querySelectorAll("[data-pptx-object-key]")]
                        .find((element) => element.dataset.pptxObjectKey === scaleKey)
                      const running = controller.play()
                      await new Promise((resolve) => setTimeout(resolve, 220))
                      indefiniteScaleBefore = scaleElement?.style.scale
                      controller.pause()
                      const pausedAt = controller.snapshot.positionMs
                      await new Promise((resolve) => setTimeout(resolve, 300))
                      indefinitePauseDeltaMs = controller.snapshot.positionMs - pausedAt
                      const resumed = controller.resume()
                      await new Promise((resolve) => setTimeout(resolve, 350))
                      indefiniteScaleAfter = scaleElement?.style.scale
                      await controller.next()
                      await Promise.all([running, resumed])
                      indefiniteFinalStatus = controller.snapshot.status
                      indefiniteWarningCount = events.filter((event) => event === "warning").length
                      await controller.next()
                      indefiniteFinalSlideIndex = controller.snapshot.slideIndex
                    } else if (repeatCheck) {
                      const running = controller.play()
                      await new Promise((resolve) => setTimeout(resolve, 250))
                      const effects = schedule.groups.flatMap((group) => group.effects)
                      const findTarget = (kind) => {
                        const key = effects.find((effect) => effect.effect.kind === kind)?.effect.targetObjectKey
                        return [...root.querySelectorAll("[data-pptx-object-key]")]
                          .find((element) => element.dataset.pptxObjectKey === key)
                      }
                      repeatScale = findTarget("scale")?.style.scale
                      repeatTranslate = findTarget("motion-path")?.style.translate
                      repeatOpacity = findTarget("emphasis")?.style.opacity
                      repeatRotate = findTarget("rotate")?.style.rotate
                      repeatClipPath = findTarget("wipe")?.style.clipPath
                      controller.pause()
                      const pausedAt = controller.snapshot.positionMs
                      await new Promise((resolve) => setTimeout(resolve, 300))
                      repeatPauseDeltaMs = controller.snapshot.positionMs - pausedAt
                      await Promise.all([running, controller.resume()])
                      repeatFinalScale = findTarget("scale")?.style.scale
                    } else {
                      await Promise.race([
                        controller.play(),
                        new Promise((_, reject) => setTimeout(() => reject(new Error("controller.play timeout")), 5000)),
                      ])
                    }
                    let autoAdvanceSlideIndex
                    let autoAdvanceStatus
                    if (autoAdvanceCheck) {
                      await new Promise((resolve) => setTimeout(resolve, 2200))
                      autoAdvanceSlideIndex = controller.snapshot.slideIndex
                      autoAdvanceStatus = controller.snapshot.status
                    }
                    let mediaPaused
                    let mediaTimeMs
                    let mediaPauseDeltaMs
                    let bookmarkEffectStarts
                    let mediaDiagnostics
                    let mediaResumeDeltaMs
                    let mediaReleasedOnSlide
                    let mediaSourceUrl
                    let mediaVolume
                    let mediaPolicy
                    let audioTimeMs
                    let audioPauseDeltaMs
                    let audioResumeDeltaMs
                    if (mediaCheck) {
                      const media = root.querySelector("video,audio")
                      if (!media) {
                        const key = playback.slides[initialSlide].media[0]?.objectKey
                        const marker = [...root.querySelectorAll("[data-pptx-object-key]")].find((node) => node.dataset.pptxObjectKey === key)
                        throw new Error(`media element missing: key=${key} marker=${marker?.outerHTML.slice(0, 1200)}`)
                      }
                      const audio = root.querySelector("audio")
                      if (!audio) throw new Error("controlled audio element missing")
                      await new Promise((resolve) => setTimeout(resolve, 1200))
                      mediaPaused = media.paused
                      mediaTimeMs = media.currentTime * 1000
                      mediaVolume = media.volume
                      const mediaItem = playback.slides[initialSlide].media[0]
                      mediaPolicy = mediaItem ? {
                        trimStartMs: mediaItem.trimStartMs,
                        trimEndMs: mediaItem.trimEndMs,
                        loop: mediaItem.loop,
                        volume: mediaItem.volume,
                        nativeLoop: media.loop,
                      } : null
                      audioTimeMs = audio.currentTime * 1000
                      mediaDiagnostics = {
                        currentSrc: media.currentSrc,
                        duration: media.duration,
                        errorCode: media.error?.code,
                        networkState: media.networkState,
                        readyState: media.readyState,
                      }
                      bookmarkEffectStarts = events.filter((event) => event === "effectstart").length
                      controller.pause()
                      const pausedAt = media.currentTime * 1000
                      const audioPausedAt = audio.currentTime * 1000
                      await new Promise((resolve) => setTimeout(resolve, 300))
                      mediaPauseDeltaMs = media.currentTime * 1000 - pausedAt
                      audioPauseDeltaMs = audio.currentTime * 1000 - audioPausedAt
                      await Promise.race([
                        controller.resume(),
                        new Promise((_, reject) => setTimeout(() => reject(new Error("controller.resume timeout")), 3000)),
                      ])
                      const resumedAt = media.currentTime * 1000
                      const audioResumedAt = audio.currentTime * 1000
                      await new Promise((resolve) => setTimeout(resolve, 500))
                      mediaResumeDeltaMs = media.currentTime * 1000 - resumedAt
                      audioResumeDeltaMs = audio.currentTime * 1000 - audioResumedAt
                      controller.pause()
                      const oldMediaUrl = media.currentSrc
                      mediaSourceUrl = oldMediaUrl
                      await controller.goToSlide(initialSlide + 1)
                      if (resourceCheck) {
                        try {
                          await fetch(oldMediaUrl)
                          mediaReleasedOnSlide = false
                        } catch {
                          mediaReleasedOnSlide = true
                        }
                      } else {
                        mediaReleasedOnSlide = !media.isConnected
                      }
                    }
                    let pauseDeltaMs
                    let transitionDuring
                    let transitionLayersDuring
                    let transitionLayersAfter
                    let transitionResponseMs
                    if (transitionCheck) {
                      for (let index = 0; index < Math.max(0, steps - 1); index += 1) await controller.next()
                      const transitionStartedAt = performance.now()
                      const navigation = controller.next()
                      transitionResponseMs = performance.now() - transitionStartedAt
                      await new Promise((resolve) => setTimeout(resolve, 100))
                      transitionDuring = controller.snapshot.status
                      transitionLayersDuring = root.querySelectorAll("[data-pptx-transition-layer]").length
                      if (pauseCheck) {
                        controller.pause()
                        const beforePause = controller.snapshot.positionMs
                        await new Promise((resolve) => setTimeout(resolve, 300))
                        pauseDeltaMs = controller.snapshot.positionMs - beforePause
                        await controller.resume()
                      }
                      await navigation
                      transitionLayersAfter = root.querySelectorAll("[data-pptx-transition-layer]").length
                    } else if (pauseCheck && steps > 0) {
                      const firstStep = controller.next()
                      await new Promise((resolve) => setTimeout(resolve, 150))
                      controller.pause()
                      const beforePause = controller.snapshot.positionMs
                      await new Promise((resolve) => setTimeout(resolve, 300))
                      pauseDeltaMs = controller.snapshot.positionMs - beforePause
                      await Promise.all([firstStep, controller.resume()])
                      for (let index = 1; index < steps; index += 1) await controller.next()
                    } else {
                      for (let index = 0; index < steps; index += 1) await controller.next()
                    }
                    let previousBoundary
                    if (previousCheck) {
                      await controller.previous()
                      previousBoundary = controller.snapshot.clickBoundary
                    }
                    playbackResult = {
                      playbackStatus: controller.snapshot.status,
                      clickBoundary: controller.snapshot.clickBoundary,
                      effectStarts: events.filter((event) => event === "effectstart").length,
                      effectEnds: events.filter((event) => event === "effectend").length,
                      stepChanges: events.filter((event) => event === "stepchange").length,
                      pauseDeltaMs,
                      previousBoundary,
                      initialAnimatedState,
                      finalAnimatedState: styleState(),
                      transitionDuring,
                      transitionLayersDuring,
                      transitionLayersAfter,
                      transitionResponseMs,
                      autoAdvanceSlideIndex,
                      autoAdvanceStatus,
                      mediaPaused,
                      mediaTimeMs,
                      mediaPauseDeltaMs,
                      bookmarkEffectStarts,
                      mediaDiagnostics,
                      mediaResumeDeltaMs,
                      mediaReleasedOnSlide,
                      mediaSourceUrl,
                      mediaVolume,
                      mediaPolicy,
                      audioTimeMs,
                      audioPauseDeltaMs,
                      audioResumeDeltaMs,
                      controllerCompileMs,
                      repeatScale,
                      repeatTranslate,
                      repeatOpacity,
                      repeatRotate,
                      repeatClipPath,
                      repeatPauseDeltaMs,
                      repeatFinalScale,
                      indefiniteScaleBefore,
                      indefiniteScaleAfter,
                      indefinitePauseDeltaMs,
                      indefiniteFinalStatus,
                      indefiniteWarningCount,
                      indefiniteFinalSlideIndex,
                      hiddenDirectRejected,
                      hiddenDirectSlideIndex,
                      hiddenSkippedSlideIndex,
                      externalActionHandled,
                      externalActionSlideIndex,
                      internalActionHandled,
                      internalActionSlideIndex,
                      actionRequests,
                      windowOpenCalls,
                      overlapInitialState,
                      overlapMidState,
                      overlapFinalState,
                      paragraphStates,
                      shapeTriggerHandled,
                      shapeTriggerKey,
                      rapidNavigationSlideIndex,
                      rapidNavigationChanges,
                      effectTiming,
                      duplicateInitialStyles,
                      duplicateFinalStyles,
                      duplicateWarnings,
                      resetInitialState,
                      resetAfterStepState,
                      resetState,
                      resetReplayState,
                      blockedStatus,
                      blockedMediaIds,
                      resumedBlockedStatus,
                      eventSequence: eventRecords.map((event) => ({
                        type: event.type,
                        slideIndex: event.slideIndex,
                        nodeId: event.nodeId,
                        objectKey: event.objectKey,
                        boundary: event.boundary,
                        from: event.from,
                        to: event.to,
                        reason: event.reason,
                        mediaId: event.mediaId,
                        actionKind: event.action?.kind,
                        warningCode: event.warning?.code,
                        errorCode: event.error?.code,
                      })),
                    }
                  } finally {
                    HTMLMediaElement.prototype.play = originalMediaPlay
                    unsubscribe()
                    controller.dispose()
                  }
                }
                const output = {
                  openMs,
                  previewSlides: preview.slides.length,
                  playbackSlides: playback.slides.length,
                  objects: playback.slides.reduce((sum, slide) => sum + slide.objects.length, 0),
                  nodes: playback.slides.reduce((sum, slide) => sum + Object.keys(slide.nodes).length, 0),
                  media: playback.slides.reduce((sum, slide) => sum + slide.media.length, 0),
                  hiddenSlides: playback.slides.filter((slide) => slide.hidden).map((slide) => slide.index),
                  actions: playback.slides.reduce((sum, slide) => sum + slide.actions.length, 0),
                  discovered: playback.capability.discovered,
                  strict: playback.capability.strict,
                  approximate: playback.capability.approximate,
                  static: playback.capability.static,
                  unparsed: playback.capability.unparsed,
                  markedObjects: root.querySelectorAll("[data-pptx-object-key]").length,
                  markedParagraphs: root.querySelectorAll("[data-pptx-paragraph-index]").length,
                  schedule: schedule.groups.map((group) => ({
                    key: group.key,
                    durationMs: group.durationMs,
                    effects: group.effects.map((effect) => ({
                      nodeId: effect.nodeId,
                      kind: effect.effect.kind,
                      targetObjectKey: effect.effect.targetObjectKey,
                      paragraphRange: effect.effect.paragraphRange,
                      triggerKey: effect.triggerKey,
                      triggerEvent: effect.triggerEvent,
                      startMs: effect.startMs,
                      endMs: effect.endMs,
                    })),
                  })),
                  focusedSlide: {
                    index: playback.slides[initialSlide].index,
                    hidden: playback.slides[initialSlide].hidden,
                    transition: playback.slides[initialSlide].transition,
                    media: playback.slides[initialSlide].media,
                    actions: playback.slides[initialSlide].actions,
                    morphFromPrevious: playback.slides[initialSlide].morphFromPrevious,
                    nodeCount: Object.keys(playback.slides[initialSlide].nodes).length,
                  },
                  capabilityFeatures: playback.capability.features,
                  ...playbackResult,
                }
                if (resourceCheck) {
                  const mediaUrl = playbackResult.mediaSourceUrl ?? root.querySelector("video,audio")?.currentSrc
                  session.dispose()
                  await new Promise((resolve) => setTimeout(resolve, 50))
                  output.remainingChildrenAfterDispose = root.childElementCount
                  if (mediaUrl) {
                    try {
                      await fetch(mediaUrl)
                      output.mediaUrlRevoked = false
                    } catch {
                      output.mediaUrlRevoked = true
                    }
                  }
                }
                return output
              } finally {
                session.dispose()
              }
            }
            """,
            {
                "moduleUrl": module_url,
                "coreUrl": core_url,
                "sampleUrl": args.sample_url,
                "initialSlide": args.initial_slide,
                "steps": args.steps,
                "pauseCheck": args.pause_check,
                "previousCheck": args.previous_check,
                "transitionCheck": args.transition_check,
                "approximation": args.approximation,
                "autoAdvanceCheck": args.auto_advance_check,
                "mediaCheck": args.media_check,
                "resourceCheck": args.resource_check,
                "repeatCheck": args.repeat_check,
                "indefiniteCheck": args.indefinite_check,
                "hiddenCheck": args.hidden_check,
                "actionCheck": args.action_check,
                "overlapCheck": args.overlap_check,
                "paragraphCheck": args.paragraph_check,
                "shapeTriggerCheck": args.shape_trigger_check,
                "rapidNavigationCheck": args.rapid_navigation_check,
                "effectTimingCheck": args.effect_timing_check,
                "duplicateTargetCheck": args.duplicate_target_check,
                "resetCheck": args.reset_check,
                "mediaBlockCheck": args.media_block_check,
            },
        )
        browser.close()

    if args.resource_check or args.media_check:
        console_errors = [message for message in console_errors if "ERR_FILE_NOT_FOUND" not in message]
    assert not console_errors, f"浏览器控制台错误：{console_errors}"
    assert not page_errors, f"页面错误：{page_errors}"
    assert result["previewSlides"] > 0, result
    assert result["playbackSlides"] == result["previewSlides"], result
    assert result["objects"] > 0, result
    assert result["nodes"] > 0, result
    assert result["discovered"] > 0, result
    assert result["markedObjects"] > 0, result
    if args.steps > 0:
        if not args.transition_check and not args.previous_check:
            assert result["effectStarts"] > 0, result
            assert result["effectStarts"] == result["effectEnds"], result
            assert result["stepChanges"] > 0, result
            assert result["initialAnimatedState"] is not None, result
            assert (
                result["initialAnimatedState"]["visibility"] == "hidden"
                or result["initialAnimatedState"]["opacity"] == "0"
                or result["initialAnimatedState"]["scale"].startswith("0")
            ), result
            assert result["finalAnimatedState"]["visibility"] != "hidden", result
    if args.pause_check:
        assert abs(result["pauseDeltaMs"]) <= 20, result
    if args.previous_check:
        assert result["previousBoundary"] == max(0, args.steps - 1), result
    if args.transition_check:
        assert result["transitionDuring"] in ("transitioning", "paused"), result
        assert result["transitionLayersDuring"] == 1, result
        assert result["transitionLayersAfter"] == 0, result
    if args.auto_advance_check:
        assert result["autoAdvanceSlideIndex"] == args.initial_slide + 1, result
        assert result["autoAdvanceStatus"] in ("ready", "waiting"), result
    if args.media_check:
        assert result["media"] == 2, result
        assert result["mediaPaused"] is False, result
        assert 100 <= result["mediaTimeMs"] <= 900, result
        assert abs(result["mediaVolume"] - 0.7) <= 0.01, result
        assert result["mediaPolicy"] == {
            "trimStartMs": 100,
            "trimEndMs": 1100,
            "loop": True,
            "volume": 0.7,
            "nativeLoop": False,
        }, result
        assert result["audioTimeMs"] >= 500, result
        assert abs(result["audioPauseDeltaMs"]) <= 40, result
        assert result["audioResumeDeltaMs"] >= 80, result
        assert abs(result["mediaPauseDeltaMs"]) <= 40, result
        assert result["mediaResumeDeltaMs"] >= 80, result
        assert result["bookmarkEffectStarts"] >= 2, result
        assert result["mediaReleasedOnSlide"] is True, result
    if args.performance_check:
        assert result["openMs"] <= 2000, result
        assert (result.get("controllerCompileMs") or 0) <= 200, result
        assert (result.get("transitionResponseMs") or 0) <= 100, result
    if args.resource_check:
        assert result["remainingChildrenAfterDispose"] == 0, result
        assert result.get("mediaUrlRevoked") is True, result
    if args.repeat_check:
        scale = [float(value) for value in result["repeatScale"].split()]
        translate = [float(value.removesuffix("px")) for value in result["repeatTranslate"].split()]
        assert any(abs(value - 1) > 0.01 for value in scale), result
        assert any(abs(value) > 1 for value in translate), result
        assert 0.3 < float(result["repeatOpacity"]) < 1, result
        assert abs(float(result["repeatRotate"].removesuffix("deg"))) > 5, result
        assert result["repeatClipPath"].startswith("inset("), result
        assert result["repeatClipPath"] not in ("inset(0% 0% 0% 0%)", "inset(0px)"), result
        assert abs(result["repeatPauseDeltaMs"]) <= 20, result
        assert all(abs(float(value) - 1) <= 0.01 for value in result["repeatFinalScale"].split()), result
    if args.indefinite_check:
        before = [float(value) for value in result["indefiniteScaleBefore"].split()]
        after = [float(value) for value in result["indefiniteScaleAfter"].split()]
        assert any(abs(value - 1) > 0.02 for value in before), result
        assert any(abs(left - right) > 0.02 for left, right in zip(before, after)), result
        assert abs(result["indefinitePauseDeltaMs"]) <= 20, result
        assert result["indefiniteFinalStatus"] == "waiting", result
        assert result["indefiniteWarningCount"] >= 1, result
        assert result["indefiniteFinalSlideIndex"] == args.initial_slide + 2, result
    if args.hidden_check:
        assert args.initial_slide + 1 in result["hiddenSlides"], result
        assert result["hiddenDirectRejected"] is True, result
        assert result["hiddenDirectSlideIndex"] == args.initial_slide + 1, result
        assert result["hiddenSkippedSlideIndex"] == args.initial_slide + 2, result
    if args.action_check:
        assert result["externalActionHandled"] is True, result
        assert result["externalActionSlideIndex"] == args.initial_slide, result
        assert result["internalActionHandled"] is True, result
        assert result["internalActionSlideIndex"] == args.initial_slide - 1, result
        assert result["actionRequests"] == ["open-url", "go-to-slide"], result
        assert result["windowOpenCalls"] == 0, result
    if args.overlap_check:
        assert result["overlapInitialState"]["visibility"] == "hidden", result
        assert 0 < float(result["overlapMidState"]["opacity"]) < 1, result
        assert result["overlapFinalState"]["visibility"] == "hidden", result
        assert float(result["overlapFinalState"]["opacity"]) == 0, result
        assert result["effectStarts"] == 2, result
        assert result["effectEnds"] == 2, result
    if args.paragraph_check:
        states = result["paragraphStates"]
        assert len(states) == 4, result
        assert all(len(state) == 3 for state in states), result
        visible_counts = [sum(
            item["visibility"] != "hidden" and item["opacity"] != "0"
            for item in state
        ) for state in states]
        assert visible_counts == [0, 1, 2, 3], result
        assert result["effectStarts"] == 3, result
        assert result["effectEnds"] == 3, result
    if args.shape_trigger_check:
        assert result["shapeTriggerHandled"] is True, result
        assert result["shapeTriggerKey"], result
        assert result["effectStarts"] >= 1, result
        assert result["effectStarts"] == result["effectEnds"], result
    if args.rapid_navigation_check:
        assert result["rapidNavigationSlideIndex"] == args.initial_slide + 3, result
        assert result["rapidNavigationChanges"] == [args.initial_slide + 1, args.initial_slide + 3], result
    if args.effect_timing_check:
        assert result["effectTiming"], result
        for timing in result["effectTiming"]:
            assert timing["observedStartMs"] is not None, result
            assert abs(timing["observedStartMs"] - timing["expectedStartMs"]) <= 100, result
            if timing["expectedEndMs"] != "indefinite":
                assert timing["observedEndMs"] is not None, result
                assert abs(timing["observedEndMs"] - timing["expectedEndMs"]) <= 100, result
    if args.duplicate_target_check:
        assert result["duplicateWarnings"] == ["TARGET_AMBIGUOUS"], result
        assert len(result["duplicateInitialStyles"]) == 2, result
        assert result["duplicateFinalStyles"] == result["duplicateInitialStyles"], result
    if args.reset_check:
        assert result["resetAfterStepState"] != result["resetInitialState"], result
        assert result["resetState"] == result["resetInitialState"], result
        assert result["resetReplayState"] == result["resetAfterStepState"], result
    if args.media_block_check:
        assert result["blockedStatus"] == "blocked", result
        assert result["blockedMediaIds"], result
        assert result["resumedBlockedStatus"] in ("running", "waiting"), result
    if not args.quiet:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(
            json.dumps(result, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )


if __name__ == "__main__":
    main()
