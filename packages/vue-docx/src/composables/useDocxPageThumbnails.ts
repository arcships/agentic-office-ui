import { getCurrentScope, onScopeDispose, ref } from "vue"
import {
  blitDocxThumbnailSurface,
  DocxThumbnailSurfaceCache,
  rasterizeDocxThumbnailSurface,
  SerialIdleTaskQueue,
  type DocxEditorController,
  type DocxPageThumbnailResolution,
  type DocxPageThumbnailResolutionOptions,
} from "@arcships/docx-core"
import {
  ensureDocxViewerPageSurfaceRegistry,
  subscribeDocxViewerPageSurfaces,
  type DocxViewerPageSurfaceChange,
} from "./page-surface-registry"

export type DocxPageThumbnailStatus = "idle" | "rendering" | "ready" | "error" | "unavailable"

export interface DocxPageThumbnailItem {
  pageIndex: number
  status: DocxPageThumbnailStatus
  error?: Error
  canvasRef?: HTMLCanvasElement
}

export interface DocxViewerThumbnails {
  /** This compatibility API has no document/controller input and cannot rasterize pages. */
  availability: "unavailable"
  /** Human-readable replacement guidance for callers and rendered canvases. */
  unavailableReason: string
  /** Attach a canvas for a specific page index. */
  attachCanvas: (pageIndex: number, canvas: HTMLCanvasElement) => void
  /** Detach a previously attached canvas. */
  detachCanvas: (pageIndex: number) => void
  /** Repaint the explicit unavailable state on every attached canvas. */
  rerender: () => Promise<void>
  /** Thumbnail resolution for all attached canvases. */
  resolution: DocxPageThumbnailResolution
}

export interface UseDocxPageThumbnailsOptions extends DocxPageThumbnailResolutionOptions {
  /** When true, all thumbnail rendering is skipped. */
  disabled?: boolean
  /** Minimum interval between consecutive raster tasks (ms). */
  minRasterIntervalMs?: number
  /** Visible page indexes used to prioritize attached renders. */
  visiblePageIndexes?: readonly number[]
  /** Additional page indexes to eagerly prefetch. */
  prefetchPageIndexes?: readonly number[]
  /** Maximum number of raster surfaces retained by this composable instance. */
  maxCacheEntries?: number
  /** Approximate pixel byte budget retained by this composable instance. */
  maxCacheBytes?: number
}

export interface UseDocxPageThumbnailsResult {
  /** Current status for each tracked page. */
  pageThumbnailStates: Map<number, { status: DocxPageThumbnailStatus; error?: Error }>
  /** Force-refresh all attached canvases (ignore the surface cache). */
  rerenderAttachedThumbnails: () => Promise<void>
  /** Lazily render a single mounted page onto the given canvas. */
  renderPageThumbnailToCanvas: (pageIndex: number, canvas?: HTMLCanvasElement) => Promise<void>
  /** Prefetch a mounted page into this composable instance's surface cache. */
  prefetchPageThumbnailSurface: (pageIndex: number) => Promise<void>
  /** Attach a canvas and register it for auto-render. */
  attachCanvasForPage: (pageIndex: number, canvas: HTMLCanvasElement) => void
  /** Detach a canvas from auto-render control. */
  detachCanvasForPage: (pageIndex: number) => void
}

interface CachedThumbnailSurface {
  contentKey: string
  resolution: DocxPageThumbnailResolution
  surface: HTMLCanvasElement
}

type PageSurfaceRenderOutcome =
  | { status: "ready"; entry: CachedThumbnailSurface }
  | { status: "unavailable" }
  | { status: "stale" }

const VIEWER_THUMBNAIL_UNAVAILABLE_MESSAGE =
  "此缩略图接口没有文档来源，无法渲染。请使用 useDocxPageThumbnails(editor) 或 DocxThumbnailPanel。"

function thumbnailResolution(
  sourceWidthPx: number,
  sourceHeightPx: number,
  options: UseDocxPageThumbnailsOptions
): DocxPageThumbnailResolution {
  const bounds = options.resolution
  const maxWidthFromBounds = typeof bounds === "number" ? bounds : bounds?.maxWidth
  const maxHeightFromBounds = typeof bounds === "number" ? bounds : bounds?.maxHeight
  const maxWidthPx = Math.max(1, options.maxWidthPx ?? maxWidthFromBounds ?? 180)
  const maxHeightPx = Math.max(
    1,
    options.maxHeightPx ?? maxHeightFromBounds ?? Number.POSITIVE_INFINITY
  )
  const scale = Math.min(1, maxWidthPx / sourceWidthPx, maxHeightPx / sourceHeightPx)
  const widthPx = Math.max(1, Math.round(sourceWidthPx * scale))
  const heightPx = Math.max(1, Math.round(sourceHeightPx * scale))
  const browserPixelRatio = typeof window === "undefined" ? 1 : window.devicePixelRatio
  const pixelRatio = Math.min(
    3,
    Math.max(1, Number.isFinite(options.pixelRatio) ? Number(options.pixelRatio) : browserPixelRatio || 1)
  )
  return {
    widthPx,
    heightPx,
    pixelWidthPx: Math.max(1, Math.round(widthPx * pixelRatio)),
    pixelHeightPx: Math.max(1, Math.round(heightPx * pixelRatio)),
    scale,
  }
}

function thumbnailResolutionsMatch(
  left: DocxPageThumbnailResolution,
  right: DocxPageThumbnailResolution
): boolean {
  return (
    left.widthPx === right.widthPx &&
    left.heightPx === right.heightPx &&
    left.pixelWidthPx === right.pixelWidthPx &&
    left.pixelHeightPx === right.pixelHeightPx &&
    left.scale === right.scale
  )
}

function paintUnavailableThumbnail(
  canvas: HTMLCanvasElement,
  resolution: DocxPageThumbnailResolution,
  message: string
): void {
  canvas.width = resolution.pixelWidthPx
  canvas.height = resolution.pixelHeightPx
  canvas.style.width = `${resolution.widthPx}px`
  canvas.style.height = `${resolution.heightPx}px`
  const context = canvas.getContext("2d")
  if (!context) return
  const scaleX = resolution.pixelWidthPx / resolution.widthPx
  const scaleY = resolution.pixelHeightPx / resolution.heightPx
  context.setTransform(scaleX, 0, 0, scaleY, 0, 0)
  context.fillStyle = "#f8fafc"
  context.fillRect(0, 0, resolution.widthPx, resolution.heightPx)
  context.strokeStyle = "#cbd5e1"
  context.strokeRect(0.5, 0.5, resolution.widthPx - 1, resolution.heightPx - 1)
  context.fillStyle = "#64748b"
  context.font = "12px sans-serif"
  context.textAlign = "center"
  context.textBaseline = "middle"
  const label = message.length > 24 ? `${message.slice(0, 24)}…` : message
  context.fillText(label, resolution.widthPx / 2, resolution.heightPx / 2)
}

function normalizeThumbnailError(error: unknown): Error {
  return error instanceof Error ? error : new Error("DOCX 缩略图渲染失败。")
}

function scheduleMicrotask(callback: () => void): void {
  if (typeof queueMicrotask === "function") {
    queueMicrotask(callback)
    return
  }
  void Promise.resolve().then(callback)
}

export function useDocxPageThumbnails(
  editor: DocxEditorController,
  options: UseDocxPageThumbnailsOptions = { sourceWidthPx: 816, sourceHeightPx: 1056 }
): UseDocxPageThumbnailsResult {
  const pageThumbnailStates = ref<Map<number, { status: DocxPageThumbnailStatus; error?: Error }>>(
    new Map()
  )
  const attachedCanvasByPage = ref<Map<number, HTMLCanvasElement>>(new Map())
  const registry = ensureDocxViewerPageSurfaceRegistry(editor)
  const surfaceCache = new DocxThumbnailSurfaceCache<CachedThumbnailSurface>({
    maxEntries: options.maxCacheEntries ?? 24,
    maxBytes: options.maxCacheBytes ?? 32 * 1024 * 1024,
    estimateBytes: (entry) =>
      entry.resolution.pixelWidthPx * entry.resolution.pixelHeightPx * 4,
  })
  const renderGeneration = new Map<number, number>()
  const surfaceContentKeys = new Map<number, string>()
  // One serial idle-time rasterizer per composable instance: replacement
  // requests for the same page coalesce, visible pages jump the queue, and
  // `minRasterIntervalMs` throttles repeated work on one page.
  const rasterQueue = new SerialIdleTaskQueue<number>({
    minTaskIntervalMs: Math.max(0, options.minRasterIntervalMs ?? 0),
  })
  const pendingPageChanges = new Map<number, DocxViewerPageSurfaceChange>()
  let pendingAllPages = false
  let notificationFlushScheduled = false
  let disposed = false

  const updatePageThumbnailState = (
    pageIndex: number,
    status: DocxPageThumbnailStatus,
    error?: Error
  ): void => {
    const next = new Map(pageThumbnailStates.value)
    const previous = next.get(pageIndex)
    if (previous?.status === status && previous?.error?.message === error?.message) return
    next.set(pageIndex, { status, error })
    pageThumbnailStates.value = next
  }

  const resolvePageResolution = (pageIndex: number): DocxPageThumbnailResolution => {
    const pageSize = registry.pageSizes.get(pageIndex) ?? registry.layoutPageSizes.get(pageIndex)
    return thumbnailResolution(
      pageSize?.widthPx ?? options.sourceWidthPx,
      pageSize?.heightPx ?? options.sourceHeightPx,
      options
    )
  }

  const renderPageSurface = async (
    pageIndex: number,
    force = false
  ): Promise<PageSurfaceRenderOutcome> => {
    const resolution = resolvePageResolution(pageIndex)
    if (options.disabled) {
      updatePageThumbnailState(pageIndex, "unavailable")
      return { status: "unavailable" }
    }

    const pageElement = registry.pageElements.get(pageIndex)
    const contentKey = registry.pageContentKeys.get(pageIndex)
    if (!pageElement || !contentKey) {
      surfaceCache.delete(String(pageIndex))
      surfaceContentKeys.delete(pageIndex)
      updatePageThumbnailState(pageIndex, "unavailable")
      return { status: "unavailable" }
    }

    const cached = surfaceCache.get(String(pageIndex))
    if (
      !force &&
      cached?.contentKey === contentKey &&
      thumbnailResolutionsMatch(cached.resolution, resolution)
    ) {
      updatePageThumbnailState(pageIndex, "ready")
      return { status: "ready", entry: cached }
    }

    const generation = (renderGeneration.get(pageIndex) ?? 0) + 1
    renderGeneration.set(pageIndex, generation)
    updatePageThumbnailState(pageIndex, "rendering")
    const sourceSize =
      registry.pageSizes.get(pageIndex) ?? registry.layoutPageSizes.get(pageIndex)
    try {
      const surface = await rasterizeDocxThumbnailSurface({
        pageElement,
        sourceWidthPx: sourceSize?.widthPx ?? options.sourceWidthPx,
        sourceHeightPx: sourceSize?.heightPx ?? options.sourceHeightPx,
        ...resolution,
      })
      if (disposed || renderGeneration.get(pageIndex) !== generation) {
        // A newer render superseded this one. Reporting "stale" keeps the
        // caller from painting an unavailable placeholder over a good image.
        return { status: "stale" }
      }
      const entry = { contentKey, resolution, surface }
      surfaceCache.set(String(pageIndex), entry)
      surfaceContentKeys.set(pageIndex, contentKey)
      updatePageThumbnailState(pageIndex, "ready")
      return { status: "ready", entry }
    } catch (error) {
      if (disposed || renderGeneration.get(pageIndex) !== generation) {
        return { status: "stale" }
      }
      const normalized = normalizeThumbnailError(error)
      updatePageThumbnailState(pageIndex, "error", normalized)
      throw normalized
    }
  }

  const paintPageThumbnailToCanvas = async (
    pageIndex: number,
    canvas?: HTMLCanvasElement,
    force = false
  ): Promise<void> => {
    const targetCanvas = canvas ?? attachedCanvasByPage.value.get(pageIndex)
    if (!targetCanvas) return
    const outcome = await renderPageSurface(pageIndex, force)
    if (outcome.status === "stale") return
    if (outcome.status === "unavailable") {
      paintUnavailableThumbnail(targetCanvas, resolvePageResolution(pageIndex), "缩略图暂不可用")
      return
    }
    blitDocxThumbnailSurface(outcome.entry.surface, targetCanvas, outcome.entry.resolution)
  }

  const pagePriority = (pageIndex: number): number =>
    options.visiblePageIndexes?.includes(pageIndex) ? 0 : 1

  const enqueuePageTask = (
    pageIndex: number,
    task: () => Promise<void>
  ): Promise<void> =>
    rasterQueue
      .enqueue(pageIndex, task, { priority: pagePriority(pageIndex) })
      .catch(() => undefined)

  const renderPageThumbnailToCanvas = (
    pageIndex: number,
    canvas?: HTMLCanvasElement
  ): Promise<void> =>
    enqueuePageTask(pageIndex, () => paintPageThumbnailToCanvas(pageIndex, canvas))

  const prefetchPageThumbnailSurface = (pageIndex: number): Promise<void> =>
    enqueuePageTask(pageIndex, async () => {
      await renderPageSurface(pageIndex)
    })

  const attachCanvasForPage = (pageIndex: number, canvas: HTMLCanvasElement): void => {
    attachedCanvasByPage.value = new Map(attachedCanvasByPage.value).set(pageIndex, canvas)
    void renderPageThumbnailToCanvas(pageIndex, canvas).catch(() => undefined)
  }

  const detachCanvasForPage = (pageIndex: number): void => {
    const next = new Map(attachedCanvasByPage.value)
    next.delete(pageIndex)
    attachedCanvasByPage.value = next
  }

  const rerenderAttachedThumbnails = async (): Promise<void> => {
    const indices = [...attachedCanvasByPage.value.keys()]
    for (const pageIndex of indices) {
      surfaceCache.delete(String(pageIndex))
      surfaceContentKeys.delete(pageIndex)
    }
    await Promise.all(
      indices.map((pageIndex) =>
        enqueuePageTask(pageIndex, () =>
          paintPageThumbnailToCanvas(pageIndex, undefined, true)
        )
      )
    )
  }

  const trackedPageIndexes = (): Set<number> =>
    new Set([
      ...attachedCanvasByPage.value.keys(),
      ...(options.prefetchPageIndexes ?? []),
    ])

  const renderTrackedPage = (pageIndex: number): void => {
    if (attachedCanvasByPage.value.has(pageIndex)) {
      void renderPageThumbnailToCanvas(pageIndex)
    } else if (options.prefetchPageIndexes?.includes(pageIndex)) {
      void prefetchPageThumbnailSurface(pageIndex)
    }
  }

  const invalidatePageCache = (
    pageIndex: number,
    contentKey?: string
  ): boolean => {
    // Legacy notifications may omit the key; only skip when the published key
    // and the requested resolution both match the surface we already rendered.
    if (contentKey !== undefined && surfaceContentKeys.get(pageIndex) === contentKey) {
      const cached = surfaceCache.get(String(pageIndex))
      if (
        cached &&
        thumbnailResolutionsMatch(cached.resolution, resolvePageResolution(pageIndex))
      ) {
        return false
      }
    }
    surfaceCache.delete(String(pageIndex))
    return true
  }

  const renderTrackedPages = (): void => {
    notificationFlushScheduled = false
    if (disposed) return

    const changedPages = [...pendingPageChanges.values()]
    pendingPageChanges.clear()
    const allPages = pendingAllPages
    pendingAllPages = false

    if (allPages) {
      // Legacy "something changed" notification: diff the registry content
      // keys so unchanged pages keep their cached canvas.
      const invalidated = new Set<number>()
      for (const [pageIndex, contentKey] of registry.pageContentKeys) {
        if (invalidatePageCache(pageIndex, contentKey)) invalidated.add(pageIndex)
      }
      for (const pageIndex of [...surfaceContentKeys.keys()]) {
        if (registry.pageContentKeys.has(pageIndex)) continue
        surfaceContentKeys.delete(pageIndex)
        surfaceCache.delete(String(pageIndex))
        invalidated.add(pageIndex)
      }
      for (const pageIndex of trackedPageIndexes()) {
        if (invalidated.has(pageIndex)) renderTrackedPage(pageIndex)
      }
      return
    }

    for (const change of changedPages) {
      const { pageIndex, kind } = change
      if (kind === "unpublish") {
        // Unmounting keeps the last good thumbnail on screen; only its cache
        // entry is released.
        surfaceContentKeys.delete(pageIndex)
        surfaceCache.delete(String(pageIndex))
        continue
      }
      if (!invalidatePageCache(pageIndex, change.contentKey)) continue
      renderTrackedPage(pageIndex)
    }
  }

  const handleSurfaceChange = (change?: DocxViewerPageSurfaceChange): void => {
    if (disposed) return
    if (!change) {
      pendingAllPages = true
    } else {
      pendingPageChanges.set(change.pageIndex, change)
    }
    if (notificationFlushScheduled) return
    notificationFlushScheduled = true
    scheduleMicrotask(renderTrackedPages)
  }

  const unsubscribe = subscribeDocxViewerPageSurfaces(editor, handleSurfaceChange)
  if (getCurrentScope()) {
    onScopeDispose(() => {
      disposed = true
      unsubscribe()
      rasterQueue.clear()
      surfaceCache.clear()
      surfaceContentKeys.clear()
      renderGeneration.clear()
      pendingPageChanges.clear()
      attachedCanvasByPage.value = new Map()
    })
  }

  return {
    get pageThumbnailStates() { return pageThumbnailStates.value },
    rerenderAttachedThumbnails,
    renderPageThumbnailToCanvas,
    prefetchPageThumbnailSurface,
    attachCanvasForPage,
    detachCanvasForPage,
  }
}

/**
 * Compatibility adapter retained for 0.x callers. It has no document or
 * controller input, so it now reports and paints an explicit unavailable
 * state instead of returning a successful-looking empty rerender function.
 */
export function useDocxViewerThumbnails(
  options: UseDocxPageThumbnailsOptions = { sourceWidthPx: 816, sourceHeightPx: 1056 }
): DocxViewerThumbnails {
  const attachedCanvasByPage = new Map<number, HTMLCanvasElement>()
  const resolution = thumbnailResolution(options.sourceWidthPx, options.sourceHeightPx, options)

  const attachCanvas = (pageIndex: number, canvas: HTMLCanvasElement): void => {
    attachedCanvasByPage.set(pageIndex, canvas)
    paintUnavailableThumbnail(canvas, resolution, VIEWER_THUMBNAIL_UNAVAILABLE_MESSAGE)
  }

  const detachCanvas = (pageIndex: number): void => {
    attachedCanvasByPage.delete(pageIndex)
  }

  const rerender = async (): Promise<void> => {
    for (const canvas of attachedCanvasByPage.values()) {
      paintUnavailableThumbnail(canvas, resolution, VIEWER_THUMBNAIL_UNAVAILABLE_MESSAGE)
    }
  }

  return {
    availability: "unavailable",
    unavailableReason: VIEWER_THUMBNAIL_UNAVAILABLE_MESSAGE,
    attachCanvas,
    detachCanvas,
    rerender,
    resolution,
  }
}
