// Vue composable: useDocxPageThumbnails
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 30612-31319
//
// Manages page thumbnail rasterization via the page surface registry.
// Supports three rendering paths: direct-draw snapshot → live-page DOM-to-SVG →
// detached-render surface rasterization.
//
// Full raster pipeline deferred to component integration phase.

import { ref, watch, onScopeDispose } from "vue"
import {
  type DocxEditorController,
  type DocxPageThumbnailResolution,
  type DocxPageThumbnailResolutionOptions,
} from "@extend-ai/docx-core"
import {
  docxViewerPageSurfaceRegistryOwner,
  ensureDocxViewerPageSurfaceRegistry,
  subscribeDocxViewerPageSurfaces,
  type DocxViewerPageSurfaceRegistry,
} from "./page-surface-registry"

export type DocxPageThumbnailStatus = "idle" | "rendering" | "ready" | "error" | "unavailable"

export interface DocxPageThumbnailItem {
  pageIndex: number
  status: DocxPageThumbnailStatus
  error?: Error
  canvasRef?: HTMLCanvasElement
}

export interface DocxViewerThumbnails {
  /** Attach a canvas for a specific page index. */
  attachCanvas: (pageIndex: number, canvas: HTMLCanvasElement) => void
  /** Detach a previously attached canvas. */
  detachCanvas: (pageIndex: number) => void
  /** Trigger re-render of all attached canvases. */
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
  /** Force-refresh all attached canvases (ignore content skip cache). */
  rerenderAttachedThumbnails: () => Promise<void>
  /** Lazily render a single page onto the given canvas. */
  renderPageThumbnailToCanvas: (pageIndex: number, canvas?: HTMLCanvasElement) => Promise<void>
  /** Prefetch a thumbnail surface for a page. */
  prefetchPageThumbnailSurface: (pageIndex: number) => Promise<void>
  /** Attach a canvas and register it for auto-render. */
  attachCanvasForPage: (pageIndex: number, canvas: HTMLCanvasElement) => void
  /** Detach a canvas from auto-render control. */
  detachCanvasForPage: (pageIndex: number) => void
}

export function useDocxPageThumbnails(
  editor: DocxEditorController,
  options: UseDocxPageThumbnailsOptions = { sourceWidthPx: 816, sourceHeightPx: 1056 }
): UseDocxPageThumbnailsResult {
  const pageSurfaceRegistryOwner = docxViewerPageSurfaceRegistryOwner(editor)

  const pageSurfaceEpoch = ref(0)
  const pageThumbnailStates = ref<Map<number, { status: DocxPageThumbnailStatus; error?: Error }>>(
    new Map()
  )
  const attachedCanvasByPage = ref<Map<number, HTMLCanvasElement>>(new Map())

  const pageSurfaceRegistry = ensureDocxViewerPageSurfaceRegistry({
    syncPaginationInfo: editor.syncPaginationInfo,
  })

  // Subscribe to surface changes
  watch(
    () => pageSurfaceRegistryOwner,
    () => {
      const unsubscribe = subscribeDocxViewerPageSurfaces(
        { syncPaginationInfo: editor.syncPaginationInfo },
        () => {
          pageSurfaceEpoch.value += 1
        }
      )
      onScopeDispose(unsubscribe)
    },
    { immediate: true }
  )

  const updatePageThumbnailState = (
    pageIndex: number,
    status: DocxPageThumbnailStatus,
    error?: Error
  ): void => {
    const next = new Map(pageThumbnailStates.value)
    const prev = next.get(pageIndex)
    if (prev?.status === status && prev?.error?.message === error?.message) return
    next.set(pageIndex, { status, error })
    pageThumbnailStates.value = next
  }

  const renderPageThumbnailToCanvas = async (
    pageIndex: number,
    canvas?: HTMLCanvasElement
  ): Promise<void> => {
    if (options.disabled) return
    const targetCanvas = canvas ?? attachedCanvasByPage.value.get(pageIndex)
    if (!targetCanvas) return
    updatePageThumbnailState(pageIndex, "rendering")
    // Full raster pipeline deferred
    updatePageThumbnailState(pageIndex, "unavailable")
  }

  const prefetchPageThumbnailSurface = async (_pageIndex: number): Promise<void> => {
    // Prefetch deferred to component integration
  }

  const attachCanvasForPage = (pageIndex: number, canvas: HTMLCanvasElement): void => {
    attachedCanvasByPage.value = new Map(attachedCanvasByPage.value).set(pageIndex, canvas)
  }

  const detachCanvasForPage = (pageIndex: number): void => {
    const next = new Map(attachedCanvasByPage.value)
    next.delete(pageIndex)
    attachedCanvasByPage.value = next
  }

  const rerenderAttachedThumbnails = async (): Promise<void> => {
    const indices = [...attachedCanvasByPage.value.keys()]
    await Promise.all(indices.map((i) => renderPageThumbnailToCanvas(i)))
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
 * Simplified observer-style thumbnails API for the DocxViewer.
 * Returns a DocxViewerThumbnails object that the viewer can attach canvases to.
 */
export function useDocxViewerThumbnails(
  options: UseDocxPageThumbnailsOptions = { sourceWidthPx: 816, sourceHeightPx: 1056 }
): DocxViewerThumbnails {
  const attachedCanvasByPage = ref<Map<number, HTMLCanvasElement>>(new Map())

  // Simple inline resolution calculation
  const sourceWidthPx = options.sourceWidthPx
  const sourceHeightPx = options.sourceHeightPx
  const maxWidthPx = options.maxWidthPx ?? 180
  const maxHeightPx = options.maxHeightPx ?? Number.POSITIVE_INFINITY
  const pixelRatio = Number.isFinite(options.pixelRatio) ? Math.max(1, options.pixelRatio as number) : 1
  const scale = Math.min(1, maxWidthPx / sourceWidthPx, maxHeightPx / sourceHeightPx)
  const widthPx = Math.max(1, Math.round(sourceWidthPx * scale))
  const heightPx = Math.max(1, Math.round(sourceHeightPx * scale))
  const resolution = {
    widthPx,
    heightPx,
    pixelWidthPx: Math.max(1, Math.round(widthPx * pixelRatio)),
    pixelHeightPx: Math.max(1, Math.round(heightPx * pixelRatio)),
    scale,
  }

  const attachCanvas = (pageIndex: number, canvas: HTMLCanvasElement): void => {
    attachedCanvasByPage.value = new Map(attachedCanvasByPage.value).set(pageIndex, canvas)
  }

  const detachCanvas = (pageIndex: number): void => {
    const next = new Map(attachedCanvasByPage.value)
    next.delete(pageIndex)
    attachedCanvasByPage.value = next
  }

  const rerender = async (): Promise<void> => {
    // Full raster pipeline deferred
  }

  return { attachCanvas, detachCanvas, rerender, resolution }
}
