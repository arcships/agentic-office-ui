import { getCurrentScope, onScopeDispose, ref } from "vue"
import {
  blitDocxThumbnailSurface,
  rasterizeDocxThumbnailSurface,
  type DocxEditorController,
  type DocxPageThumbnailResolution,
  type DocxPageThumbnailResolutionOptions,
} from "@extend-ai/docx-core"
import {
  ensureDocxViewerPageSurfaceRegistry,
  subscribeDocxViewerPageSurfaces,
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

export function useDocxPageThumbnails(
  editor: DocxEditorController,
  options: UseDocxPageThumbnailsOptions = { sourceWidthPx: 816, sourceHeightPx: 1056 }
): UseDocxPageThumbnailsResult {
  const pageThumbnailStates = ref<Map<number, { status: DocxPageThumbnailStatus; error?: Error }>>(
    new Map()
  )
  const attachedCanvasByPage = ref<Map<number, HTMLCanvasElement>>(new Map())
  const registry = ensureDocxViewerPageSurfaceRegistry(editor)
  const surfaceCache = new Map<number, CachedThumbnailSurface>()
  const renderGeneration = new Map<number, number>()
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
    const pageSize = registry.pageSizes.get(pageIndex)
    return thumbnailResolution(
      pageSize?.widthPx ?? options.sourceWidthPx,
      pageSize?.heightPx ?? options.sourceHeightPx,
      options
    )
  }

  const renderPageSurface = async (
    pageIndex: number,
    force = false
  ): Promise<CachedThumbnailSurface | undefined> => {
    const resolution = resolvePageResolution(pageIndex)
    if (options.disabled) {
      updatePageThumbnailState(pageIndex, "unavailable")
      return undefined
    }

    const pageElement = registry.pageElements.get(pageIndex)
    const contentKey = registry.pageContentKeys.get(pageIndex)
    if (!pageElement || !contentKey) {
      surfaceCache.delete(pageIndex)
      updatePageThumbnailState(pageIndex, "unavailable")
      return undefined
    }

    const cached = surfaceCache.get(pageIndex)
    if (!force && cached?.contentKey === contentKey) return cached

    const generation = (renderGeneration.get(pageIndex) ?? 0) + 1
    renderGeneration.set(pageIndex, generation)
    updatePageThumbnailState(pageIndex, "rendering")
    try {
      const surface = await rasterizeDocxThumbnailSurface({
        pageElement,
        sourceWidthPx: registry.pageSizes.get(pageIndex)?.widthPx ?? options.sourceWidthPx,
        sourceHeightPx: registry.pageSizes.get(pageIndex)?.heightPx ?? options.sourceHeightPx,
        ...resolution,
      })
      if (disposed || renderGeneration.get(pageIndex) !== generation) return undefined
      const entry = { contentKey, resolution, surface }
      surfaceCache.set(pageIndex, entry)
      updatePageThumbnailState(pageIndex, "ready")
      return entry
    } catch (error) {
      if (disposed || renderGeneration.get(pageIndex) !== generation) return undefined
      const normalized = normalizeThumbnailError(error)
      updatePageThumbnailState(pageIndex, "error", normalized)
      throw normalized
    }
  }

  const renderPageThumbnailToCanvas = async (
    pageIndex: number,
    canvas?: HTMLCanvasElement
  ): Promise<void> => {
    const targetCanvas = canvas ?? attachedCanvasByPage.value.get(pageIndex)
    if (!targetCanvas) return
    const entry = await renderPageSurface(pageIndex)
    if (!entry) {
      paintUnavailableThumbnail(targetCanvas, resolvePageResolution(pageIndex), "缩略图暂不可用")
      return
    }
    blitDocxThumbnailSurface(entry.surface, targetCanvas, entry.resolution)
  }

  const prefetchPageThumbnailSurface = async (pageIndex: number): Promise<void> => {
    await renderPageSurface(pageIndex)
  }

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
    surfaceCache.clear()
    const indices = [...attachedCanvasByPage.value.keys()]
    await Promise.all(indices.map((pageIndex) => renderPageThumbnailToCanvas(pageIndex)))
  }

  const renderTrackedPages = (): void => {
    surfaceCache.clear()
    const trackedPageIndexes = new Set([
      ...attachedCanvasByPage.value.keys(),
      ...(options.prefetchPageIndexes ?? []),
    ])
    for (const pageIndex of trackedPageIndexes) {
      if (attachedCanvasByPage.value.has(pageIndex)) {
        void renderPageThumbnailToCanvas(pageIndex).catch(() => undefined)
      } else {
        void prefetchPageThumbnailSurface(pageIndex).catch(() => undefined)
      }
    }
  }

  const unsubscribe = subscribeDocxViewerPageSurfaces(editor, renderTrackedPages)
  if (getCurrentScope()) {
    onScopeDispose(() => {
      disposed = true
      unsubscribe()
      surfaceCache.clear()
      renderGeneration.clear()
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
