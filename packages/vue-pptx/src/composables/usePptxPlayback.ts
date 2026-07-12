import {
  computed,
  isRef,
  onScopeDispose,
  shallowReadonly,
  shallowRef,
  watch,
  type Ref,
} from "vue"
import {
  PptxPlaybackError,
  type PptxCapabilityReport,
  type PptxPlaybackWarning,
} from "@arcships/pptx-core"
import type {
  PptxPlaybackController,
  PptxPlaybackEvent,
  PptxPlaybackSnapshot,
} from "@arcships/pptx-core/browser"
import type {
  UsePptxDocumentReturn,
  UsePptxPlaybackOptions,
  UsePptxPlaybackReturn,
} from "../headless-types"
import { getDocumentBinding, isDocumentSession } from "./internal"

function resolveValue<T>(value: T | Readonly<Ref<T>> | (() => T)): T {
  if (typeof value === "function") return (value as () => T)()
  return isRef(value) ? value.value : value
}

export function usePptxPlayback(
  document: UsePptxDocumentReturn,
  options: UsePptxPlaybackOptions = {},
): UsePptxPlaybackReturn {
  const foundBinding = getDocumentBinding(document)
  if (!foundBinding) throw new Error("usePptxPlayback 必须接收 usePptxDocument 的返回值。")
  const binding = foundBinding

  const controllerValue = shallowRef<PptxPlaybackController | null>(null)
  const snapshotValue = shallowRef<PptxPlaybackSnapshot | null>(null)
  const capabilityValue = shallowRef<PptxCapabilityReport | null>(null)
  const lastWarningValue = shallowRef<PptxPlaybackWarning | null>(null)
  const lastErrorValue = shallowRef<PptxPlaybackError | null>(null)
  let unsubscribe: (() => void) | null = null
  let disposed = false

  function releaseController(): void {
    unsubscribe?.()
    unsubscribe = null
    const current = controllerValue.value
    controllerValue.value = null
    current?.dispose()
    snapshotValue.value = null
    capabilityValue.value = null
  }

  function onEvent(event: PptxPlaybackEvent): void {
    if (event.type === "statechange") {
      snapshotValue.value = event.snapshot
      binding.setActiveIndex(event.snapshot.slideIndex)
    } else if (event.type === "slidechange") {
      binding.setActiveIndex(event.to)
    } else if (event.type === "capability") {
      capabilityValue.value = event.report
    } else if (event.type === "warning") {
      lastWarningValue.value = event.warning
    } else if (event.type === "error") {
      lastErrorValue.value = event.error
    }
    options.onEvent?.(event)
  }

  function createController(): void {
    releaseController()
    if (
      disposed
      || (options.enabled !== undefined && resolveValue(options.enabled) === false)
    ) return
    const session = binding.session.value
    if (!isDocumentSession(session) || document.state.value !== "ready") return
    const nextController = session.createPlaybackController({
      initialSlide: document.activeIndex.value,
      skipHiddenSlides: options.skipHiddenSlides,
      autoplay: options.autoplay,
      approximation: options.approximation,
    })
    controllerValue.value = nextController
    snapshotValue.value = nextController.snapshot
    capabilityValue.value = session.capabilityReport
    unsubscribe = nextController.subscribe(onEvent)
  }

  function requireController(): PptxPlaybackController {
    if (!controllerValue.value) {
      throw new PptxPlaybackError("PLAYBACK_NOT_READY", "PPTX 播放控制器尚未准备好。")
    }
    return controllerValue.value
  }

  function dispose(): void {
    if (disposed) return
    disposed = true
    releaseController()
  }

  watch(
    [
      binding.session,
      document.state,
      () => options.enabled === undefined ? true : resolveValue(options.enabled),
    ],
    createController,
    { flush: "sync" },
  )

  onScopeDispose(dispose)

  return {
    controller: shallowReadonly(controllerValue),
    snapshot: shallowReadonly(snapshotValue),
    status: computed(() => snapshotValue.value?.status ?? "unavailable"),
    capability: shallowReadonly(capabilityValue),
    lastWarning: shallowReadonly(lastWarningValue),
    lastError: shallowReadonly(lastErrorValue),
    next: () => requireController().next(),
    activateObject: (objectKey) => requireController().activateObject(objectKey),
    previous: () => requireController().previous(),
    play: () => requireController().play(),
    pause: () => requireController().pause(),
    resume: () => requireController().resume(),
    reset: () => requireController().reset(),
    goToSlide: (index, goToOptions) => requireController().goToSlide(index, goToOptions),
    resumeBlockedMedia: (mediaId) => requireController().resumeBlockedMedia(mediaId),
    dispose,
  }
}
