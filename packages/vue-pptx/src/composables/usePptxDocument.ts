import {
  onScopeDispose,
  readonly,
  ref,
  shallowReadonly,
  shallowRef,
  toValue,
  watch,
  type MaybeRefOrGetter,
} from "vue"
import {
  PptxPreviewError,
  type PptxCapabilityReport,
  type PptxPreviewDocument,
} from "@arcships/pptx-core"
import {
  createPptxDocumentSession,
  type PptxDocumentSession,
  type PptxDocumentSessionFactory,
  type PptxPreviewSession,
  type PptxPreviewSessionFactory,
  type PptxPreviewSource,
} from "@arcships/pptx-core/browser"
import type {
  PptxDocumentState,
  PptxStageTarget,
  UsePptxDocumentOptions,
  UsePptxDocumentReturn,
} from "../headless-types"
import { isDocumentSession, registerDocumentBinding } from "./internal"

interface InternalDocumentOptions extends UsePptxDocumentOptions {
  factory?: PptxDocumentSessionFactory | PptxPreviewSessionFactory
}

interface PendingOpen {
  source: PptxPreviewSource
  resolve(document: PptxPreviewDocument): void
  reject(reason: unknown): void
}

function resolveTarget(target: PptxStageTarget): HTMLElement | null {
  return typeof target === "function" ? target() : target.value
}

function toPreviewError(reason: unknown): PptxPreviewError {
  return reason instanceof PptxPreviewError
    ? reason
    : new PptxPreviewError(
      "PARSE_FAILED",
      reason instanceof Error ? reason.message : "无法加载 PPTX。",
      reason,
    )
}

export function usePptxDocument(
  target: PptxStageTarget,
  options: UsePptxDocumentOptions = {},
): UsePptxDocumentReturn {
  const internalOptions = options as InternalDocumentOptions
  const stateValue = ref<PptxDocumentState>("idle")
  const errorValue = shallowRef<PptxPreviewError | null>(null)
  const documentValue = shallowRef<PptxPreviewDocument | null>(null)
  const capabilityValue = shallowRef<PptxCapabilityReport | null>(null)
  const activeIndexValue = ref(0)
  const zoomPercentValue = ref(100)
  const sessionValue = shallowRef<PptxPreviewSession | null>(null)
  let latestSource: PptxPreviewSource | null = null
  let generation = 0
  let navigationGeneration = 0
  let disposed = false
  let pendingOpen: PendingOpen | null = null

  function releaseSession(): void {
    const current = sessionValue.value
    sessionValue.value = null
    current?.dispose()
  }

  async function open(source: PptxPreviewSource): Promise<PptxPreviewDocument> {
    if (disposed) throw new PptxPreviewError("RENDER_FAILED", "PPTX 文档已经销毁。")
    if (pendingOpen && pendingOpen.source !== source) {
      pendingOpen.reject(new PptxPreviewError("STALE_RESULT", "PPTX 打开请求已经被替换。"))
      pendingOpen = null
    }
    latestSource = source
    const element = resolveTarget(target)
    if (!element) {
      stateValue.value = "waiting-for-stage"
      pendingOpen?.reject(new PptxPreviewError("STALE_RESULT", "PPTX 打开请求已经被替换。"))
      return new Promise<PptxPreviewDocument>((resolve, reject) => {
        pendingOpen = { source, resolve, reject }
      })
    }

    const currentGeneration = ++generation
    navigationGeneration += 1
    releaseSession()
    documentValue.value = null
    capabilityValue.value = null
    errorValue.value = null
    activeIndexValue.value = 0
    zoomPercentValue.value = 100
    stateValue.value = "loading"

    const sessionOptions = toValue(internalOptions.session) ?? {}
    const factory = internalOptions.factory ?? createPptxDocumentSession
    const current = factory(element, sessionOptions)
    sessionValue.value = current

    try {
      const initialSlide = Math.max(0, Math.trunc(toValue(internalOptions.initialSlide) ?? 0))
      const loaded = await current.open(source, {
        initialSlide,
        limits: sessionOptions.limits,
        lazyMedia: sessionOptions.lazyMedia,
        lazySlides: sessionOptions.lazySlides,
      })
      if (disposed || currentGeneration !== generation || sessionValue.value !== current) {
        throw new PptxPreviewError("STALE_RESULT", "PPTX 加载结果已经过期。")
      }
      documentValue.value = loaded
      activeIndexValue.value = Math.min(initialSlide, Math.max(loaded.slides.length - 1, 0))
      zoomPercentValue.value = current.zoomPercent
      capabilityValue.value = isDocumentSession(current) ? current.capabilityReport : null
      stateValue.value = "ready"
      if (pendingOpen?.source === source) {
        pendingOpen.resolve(loaded)
        pendingOpen = null
      }
      return loaded
    } catch (reason) {
      const failure = toPreviewError(reason)
      if (currentGeneration === generation && sessionValue.value === current) {
        if (failure.code !== "ABORTED" && failure.code !== "STALE_RESULT") {
          errorValue.value = failure
          stateValue.value = "error"
        }
      }
      if (pendingOpen?.source === source) {
        pendingOpen.reject(failure)
        pendingOpen = null
      }
      throw failure
    }
  }

  function close(): void {
    generation += 1
    navigationGeneration += 1
    latestSource = null
    pendingOpen?.reject(new PptxPreviewError("ABORTED", "PPTX 打开请求已经取消。"))
    pendingOpen = null
    releaseSession()
    documentValue.value = null
    capabilityValue.value = null
    errorValue.value = null
    activeIndexValue.value = 0
    zoomPercentValue.value = 100
    if (!disposed) stateValue.value = "idle"
  }

  async function goTo(index: number): Promise<void> {
    const current = sessionValue.value
    const slides = documentValue.value?.slides
    if (!current || !slides?.length || stateValue.value !== "ready") return
    const targetIndex = Math.min(Math.max(Math.trunc(index), 0), slides.length - 1)
    const currentGeneration = generation
    const currentNavigation = ++navigationGeneration
    try {
      await current.renderSlide(targetIndex)
      if (
        currentGeneration === generation
        && currentNavigation === navigationGeneration
        && sessionValue.value === current
      ) {
        activeIndexValue.value = targetIndex
      }
    } catch (reason) {
      if (currentGeneration === generation && currentNavigation === navigationGeneration) {
        const failure = reason instanceof PptxPreviewError
          ? reason
          : new PptxPreviewError("RENDER_FAILED", "页面渲染失败。", reason)
        errorValue.value = failure
        stateValue.value = "error"
        throw failure
      }
    }
  }

  async function setZoom(percent: number): Promise<void> {
    const current = sessionValue.value
    if (!current || stateValue.value !== "ready") return
    const currentGeneration = generation
    await current.setZoom(percent)
    if (currentGeneration === generation && sessionValue.value === current) {
      zoomPercentValue.value = current.zoomPercent
    }
  }

  function dispose(): void {
    if (disposed) return
    disposed = true
    generation += 1
    navigationGeneration += 1
    pendingOpen?.reject(new PptxPreviewError("ABORTED", "PPTX 打开请求已经取消。"))
    pendingOpen = null
    releaseSession()
    documentValue.value = null
    capabilityValue.value = null
    stateValue.value = "disposed"
  }

  const result: UsePptxDocumentReturn = {
    state: readonly(stateValue),
    error: shallowReadonly(errorValue),
    document: shallowReadonly(documentValue),
    capability: shallowReadonly(capabilityValue),
    activeIndex: readonly(activeIndexValue),
    zoomPercent: readonly(zoomPercentValue),
    open,
    close,
    goTo,
    nextSlide: () => goTo(activeIndexValue.value + 1),
    previousSlide: () => goTo(activeIndexValue.value - 1),
    setZoom,
    getSession: () => sessionValue.value as PptxDocumentSession | null,
    dispose,
  }

  registerDocumentBinding(result, {
    session: sessionValue,
    setActiveIndex(index) {
      activeIndexValue.value = index
    },
  })

  watch(
    [
      () => resolveTarget(target),
      () => toValue(options.source as MaybeRefOrGetter<PptxPreviewSource | null | undefined>),
    ],
    ([element, source], [previousElement]) => {
      if (disposed) return
      if (element !== previousElement && sessionValue.value) {
        generation += 1
        releaseSession()
      }
      if (options.source !== undefined && !source) {
        close()
        return
      }
      const nextSource = source ?? latestSource
      if (!nextSource) {
        return
      }
      latestSource = nextSource
      if (!element) {
        stateValue.value = "waiting-for-stage"
        return
      }
      void open(nextSource).catch(() => undefined)
    },
    { immediate: true },
  )

  onScopeDispose(dispose)
  return result
}
