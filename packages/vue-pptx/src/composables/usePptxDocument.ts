import {
  isRef,
  onScopeDispose,
  readonly,
  ref,
  shallowReadonly,
  shallowRef,
  watch,
  type Ref,
} from "vue"
import {
  PptxPreviewError,
  type PptxCapabilityReport,
  type PptxPreviewDocument,
  type PptxSearchResult,
} from "@arcships/pptx-core"
import {
  createSurfaceSearchSession,
} from "@arcships/office-runtime"
import {
  createPptxDocumentSession,
  type PptxDocumentSession,
  type PptxDocumentSessionFactory,
  type PptxPreviewSession,
  type PptxPreviewSessionFactory,
  type PptxPreviewSource,
} from "@arcships/pptx-core/browser"
import { formatOfficeMessage, officeMessagesEn } from "@arcships/office-interaction"
import type {
  PptxDocumentState,
  PptxOpenOptions,
  PptxSearchState,
  PptxSlideFailure,
  PptxStageTarget,
  PptxDocumentErrorMessages,
  UsePptxDocumentOptions,
  UsePptxDocumentReturn,
} from "../headless-types"
import { isDocumentSession, registerDocumentBinding } from "./internal"
import { watchFontLoading } from "./font-loading"

interface InternalDocumentOptions extends UsePptxDocumentOptions {
  factory?: PptxDocumentSessionFactory | PptxPreviewSessionFactory
}

interface PendingOpen {
  source: PptxPreviewSource
  resolve(document: PptxPreviewDocument): void
  reject(reason: unknown): void
  timer?: ReturnType<typeof setTimeout> | null
}

const DEFAULT_OPEN_TIMEOUT_MS = 30_000

function resolveTarget(target: PptxStageTarget): HTMLElement | null {
  return typeof target === "function" ? target() : target.value
}

function resolveValue<T>(value: T | Readonly<Ref<T>> | (() => T)): T {
  if (typeof value === "function") return (value as () => T)()
  return isRef(value) ? value.value : value
}

function toPreviewError(reason: unknown, loadFailed: string): PptxPreviewError {
  return reason instanceof PptxPreviewError
    ? reason
    : new PptxPreviewError(
      "PARSE_FAILED",
      reason instanceof Error ? reason.message : loadFailed,
      reason,
    )
}

function resolveOpenTimeoutMs(options: InternalDocumentOptions): number {
  const raw = options.openTimeoutMs === undefined
    ? undefined
    : resolveValue(options.openTimeoutMs)
  if (raw === undefined || raw === null) return DEFAULT_OPEN_TIMEOUT_MS
  const numeric = Number(raw)
  return Number.isFinite(numeric) && numeric > 0 ? numeric : 0
}

function isCancellation(error: PptxPreviewError): boolean {
  return error.code === "ABORTED" || error.code === "STALE_RESULT"
}

function clearPendingTimer(pending: PendingOpen): void {
  if (pending.timer != null) {
    clearTimeout(pending.timer)
    pending.timer = null
  }
}

export function usePptxDocument(
  target: PptxStageTarget,
  options: UsePptxDocumentOptions = {},
): UsePptxDocumentReturn {
  const internalOptions = options as InternalDocumentOptions
  const errorMessage = (
    key: keyof PptxDocumentErrorMessages,
    values: Record<string, string | number> = {},
  ): string => formatOfficeMessage(
    resolveValue(options.messages)?.[key]
      ?? officeMessagesEn.pptx.documentErrors[key],
    values,
  )
  const stateValue = ref<PptxDocumentState>("idle")
  const errorValue = shallowRef<PptxPreviewError | null>(null)
  const documentValue = shallowRef<PptxPreviewDocument | null>(null)
  const capabilityValue = shallowRef<PptxCapabilityReport | null>(null)
  const activeIndexValue = ref(0)
  const zoomPercentValue = ref(100)
  const sessionValue = shallowRef<PptxPreviewSession | null>(null)
  const slideErrorValue = shallowRef<PptxSlideFailure | null>(null)
  const searchStateValue = shallowRef<PptxSearchState>({
    status: "idle",
    query: "",
    matches: [],
    activeIndex: -1,
  })
  let latestSource: PptxPreviewSource | null = null
  let generation = 0
  let navigationGeneration = 0
  let disposed = false
  let pendingOpen: PendingOpen | null = null
  let stopFontLoadingWatcher: (() => void) | null = null

  const surfaceSearch = createSurfaceSearchSession<PptxSearchResult>({
    search(query, { signal }) {
      const current = sessionValue.value
      if (!current || signal.aborted) return []
      return current.searchText(query)
    },
    async activate(match, _index, { signal }) {
      const current = sessionValue.value
      if (!current) throw new PptxPreviewError("RENDER_FAILED", errorMessage("previewNotReady"))
      await goTo(match.slideIndex)
      if (signal.aborted || sessionValue.value !== current) return
      await current.highlightSearchResult(match)
    },
    clear() {
      sessionValue.value?.clearSearchHighlights()
    },
  })
  surfaceSearch.subscribe((next) => {
    searchStateValue.value = next
  })

  function stopFontWatcher(): void {
    stopFontLoadingWatcher?.()
    stopFontLoadingWatcher = null
  }

  async function refresh(): Promise<void> {
    const current = sessionValue.value
    if (!current || stateValue.value !== "ready" || typeof current.refresh !== "function") return
    await current.refresh()
  }

  function startFontWatcher(): void {
    stopFontWatcher()
    if (disposed) return
    stopFontLoadingWatcher = watchFontLoading(() => {
      void refresh().catch(() => undefined)
    })
  }

  function releaseSession(): void {
    stopFontWatcher()
    surfaceSearch.clearSearch()
    const current = sessionValue.value
    sessionValue.value = null
    current?.dispose()
  }

  async function open(source: PptxPreviewSource, openOptions?: PptxOpenOptions): Promise<PptxPreviewDocument> {
    if (disposed) throw new PptxPreviewError("RENDER_FAILED", errorMessage("disposed"))
    if (pendingOpen && pendingOpen.source !== source) {
      pendingOpen.reject(new PptxPreviewError("STALE_RESULT", errorMessage("staleOpen")))
      pendingOpen = null
    }
    latestSource = source
    const element = resolveTarget(target)
    const openTimeoutMs = resolveOpenTimeoutMs(internalOptions)
    if (!element) {
      stateValue.value = "waiting-for-stage"
      pendingOpen?.reject(new PptxPreviewError("STALE_RESULT", errorMessage("staleOpen")))
      return new Promise<PptxPreviewDocument>((resolve, reject) => {
        const timedOut = (): void => {
          if (pendingOpen !== record) return
          pendingOpen = null
          const failure = new PptxPreviewError("TIMEOUT", errorMessage("timeout"))
          if (!disposed && latestSource === source) {
            errorValue.value = failure
            stateValue.value = "error"
          }
          reject(failure)
        }
        const record: PendingOpen = {
          source,
          timer: openTimeoutMs > 0 ? setTimeout(timedOut, openTimeoutMs) : null,
          resolve(loaded) {
            clearPendingTimer(record)
            resolve(loaded)
          },
          reject(reason) {
            clearPendingTimer(record)
            reject(reason)
          },
        }
        pendingOpen = record
      })
    }

    const currentGeneration = ++generation
    navigationGeneration += 1
    // The stage finally arrived for an earlier waiting open. Stop its timeout
    // so only the real load can settle the request.
    if (pendingOpen?.source === source) clearPendingTimer(pendingOpen)
    releaseSession()
    documentValue.value = null
    capabilityValue.value = null
    errorValue.value = null
    slideErrorValue.value = null
    // Keep activeIndex until the new session reports its position so a reload
    // of the same deck does not flash back to slide 1.
    zoomPercentValue.value = 100
    stateValue.value = "loading"

    const sessionOptions = internalOptions.session === undefined
      ? {}
      : resolveValue(internalOptions.session) ?? {}
    const factory = internalOptions.factory ?? createPptxDocumentSession
    const current = factory(element, {
      ...sessionOptions,
      onSlideChange(index) {
        sessionOptions.onSlideChange?.(index)
        if (
          currentGeneration === generation
          && sessionValue.value === current
        ) {
          activeIndexValue.value = index
        }
      },
    })
    sessionValue.value = current
    let timeoutHandle: ReturnType<typeof setTimeout> | null = null

    try {
      const initialSlide = Math.max(0, Math.trunc(
        openOptions?.initialSlide
          ?? (internalOptions.initialSlide === undefined
            ? 0
            : resolveValue(internalOptions.initialSlide) ?? 0),
      ))
      const loading = current.open(source, {
        initialSlide,
        limits: sessionOptions.limits,
        lazyMedia: sessionOptions.lazyMedia,
        lazySlides: sessionOptions.lazySlides,
      })
      // The timeout may win the race; keep the late result from becoming an
      // unhandled rejection.
      void loading.catch(() => undefined)
      const loaded = openTimeoutMs > 0
        ? await Promise.race([
          loading,
          new Promise<never>((_, reject) => {
            timeoutHandle = setTimeout(() => {
              reject(new PptxPreviewError(
                "TIMEOUT",
                errorMessage("timeoutMs", { ms: openTimeoutMs }),
              ))
            }, openTimeoutMs)
          }),
        ])
        : await loading
      if (disposed || currentGeneration !== generation || sessionValue.value !== current) {
        throw new PptxPreviewError("STALE_RESULT", errorMessage("staleResult"))
      }
      documentValue.value = loaded
      activeIndexValue.value = Math.min(initialSlide, Math.max(loaded.slides.length - 1, 0))
      zoomPercentValue.value = current.zoomPercent
      capabilityValue.value = isDocumentSession(current) ? current.capabilityReport : null
      slideErrorValue.value = null
      stateValue.value = "ready"
      startFontWatcher()
      if (pendingOpen?.source === source) {
        pendingOpen.resolve(loaded)
        pendingOpen = null
      }
      return loaded
    } catch (reason) {
      const failure = toPreviewError(reason, errorMessage("loadFailed"))
      if (currentGeneration === generation && sessionValue.value === current) {
        errorValue.value = failure
        if (isCancellation(failure)) {
          stateValue.value = "cancelled"
          releaseSession()
        } else {
          stateValue.value = "error"
          if (failure.code === "TIMEOUT") {
            current.cancel?.()
            releaseSession()
          }
        }
      }
      if (pendingOpen?.source === source) {
        pendingOpen.reject(failure)
        pendingOpen = null
      }
      throw failure
    } finally {
      if (timeoutHandle !== null) clearTimeout(timeoutHandle)
    }
  }

  function close(): void {
    generation += 1
    navigationGeneration += 1
    latestSource = null
    pendingOpen?.reject(new PptxPreviewError("ABORTED", errorMessage("aborted")))
    pendingOpen = null
    releaseSession()
    documentValue.value = null
    capabilityValue.value = null
    errorValue.value = null
    slideErrorValue.value = null
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
        slideErrorValue.value = null
      }
    } catch (reason) {
      if (currentGeneration === generation && currentNavigation === navigationGeneration) {
        const failure = reason instanceof PptxPreviewError
          ? reason
          : new PptxPreviewError(
            "RENDER_FAILED",
            reason instanceof Error && reason.message
              ? errorMessage("slideRenderFailedWithMessage", { message: reason.message })
              : errorMessage("slideRenderFailed"),
            reason,
          )
        // A single broken slide is local to that slide: keep the document ready
        // so the rest of the deck stays reachable, and let the host retry or
        // skip it.
        slideErrorValue.value = { index: targetIndex, error: failure }
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
    pendingOpen?.reject(new PptxPreviewError("ABORTED", errorMessage("aborted")))
    pendingOpen = null
    releaseSession()
    surfaceSearch.dispose()
    documentValue.value = null
    capabilityValue.value = null
    slideErrorValue.value = null
    stateValue.value = "disposed"
  }

  function clearSlideError(): void {
    slideErrorValue.value = null
  }

  const result: UsePptxDocumentReturn = {
    state: readonly(stateValue),
    error: shallowReadonly(errorValue),
    document: shallowReadonly(documentValue),
    capability: shallowReadonly(capabilityValue),
    activeIndex: readonly(activeIndexValue),
    zoomPercent: readonly(zoomPercentValue),
    searchState: shallowReadonly(searchStateValue),
    slideError: shallowReadonly(slideErrorValue),
    open,
    close,
    goTo,
    nextSlide: () => goTo(activeIndexValue.value + 1),
    previousSlide: () => goTo(activeIndexValue.value - 1),
    setZoom,
    refresh,
    search: (query) => surfaceSearch.search(query),
    activateSearchMatch: (index) => surfaceSearch.activateSearchMatch(index),
    searchNext: () => surfaceSearch.searchNext(),
    searchPrevious: () => surfaceSearch.searchPrevious(),
    clearSearch: () => surfaceSearch.clearSearch(),
    clearSlideError,
    getSearchState: () => surfaceSearch.getSearchState(),
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
      () => options.source === undefined ? undefined : resolveValue(options.source),
    ],
    ([element, source], [previousElement]) => {
      if (disposed) return
      const preservedIndex = activeIndexValue.value
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
      // Re-opening the same deck (a replaced stage element, a re-mount or a
      // mode switch) keeps the reader on the slide they were viewing.
      const keepPosition = nextSource === latestSource
      latestSource = nextSource
      if (!element) {
        stateValue.value = "waiting-for-stage"
        return
      }
      void open(nextSource, keepPosition ? { initialSlide: preservedIndex } : undefined)
        .catch(() => undefined)
    },
    { immediate: true },
  )

  onScopeDispose(dispose)
  return result
}
