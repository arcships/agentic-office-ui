import { ref, type Ref } from "vue"
import type { PdfRenderDocument, PdfRenderRuntime } from "../pdf/pdf-render-runtime"

export interface PdfThumbnailLoaderOptions {
  /** Concurrent thumbnail renders. Defaults to two, matching the viewer. */
  concurrency?: number
  /** Called for every object URL the loader releases. */
  onRevoke?: (url: string) => void
}

export interface PdfThumbnailUse {
  /** Object URLs keyed by 1-based page number. */
  urls: Ref<Record<number, string>>
  /** Rasterize missing thumbnails for a document; reuse already loaded pages. */
  load(
    runtime: PdfRenderRuntime,
    document: PdfRenderDocument,
    options?: { signal?: AbortSignal },
  ): Promise<void>
  /** Abort pending renders and release every thumbnail URL. */
  clear(): void
  /** Release the loader for good. */
  dispose(): void
}

/**
 * Thumbnail pipeline shared by `PdfViewer` and `PdfSurface`.
 *
 * Thumbnails are rasterized in a bounded worker pool, each page gets one object
 * URL that stays until the document changes, and stale generations or aborted
 * requests can never publish pixels for a document that is already gone.
 */
export function usePdfThumbnails(options: PdfThumbnailLoaderOptions = {}): PdfThumbnailUse {
  const urls = ref<Record<number, string>>({})
  const concurrency = Number.isFinite(options.concurrency)
    ? Math.max(1, Math.floor(options.concurrency as number))
    : 2
  let generation = 0
  let controller: AbortController | undefined
  let disposed = false

  const revoke = (url: string): void => {
    if (!url) return
    try {
      URL.revokeObjectURL(url)
      options.onRevoke?.(url)
    } catch {
      // Revocation is best-effort; the URL is already unreachable from the DOM.
    }
  }

  function clear(): void {
    generation += 1
    controller?.abort()
    controller = undefined
    for (const url of Object.values(urls.value)) revoke(url)
    urls.value = {}
  }

  async function load(
    runtime: PdfRenderRuntime,
    document: PdfRenderDocument,
    requestOptions: { signal?: AbortSignal } = {},
  ): Promise<void> {
    if (disposed || document.pageCount <= 0) return
    const current = ++generation
    controller?.abort()
    const abort = new AbortController()
    controller = abort
    const forwardAbort = (): void => abort.abort()
    requestOptions.signal?.addEventListener("abort", forwardAbort, { once: true })
    if (requestOptions.signal?.aborted) abort.abort()

    let nextIndex = 0
    const worker = async (): Promise<void> => {
      while (!disposed && current === generation && !abort.signal.aborted) {
        const pageIndex = nextIndex++
        if (pageIndex >= document.pageCount) return
        if (urls.value[pageIndex + 1]) continue
        let url = ""
        try {
          const blob = await runtime.renderThumbnail(document, pageIndex, { signal: abort.signal })
          if (disposed || current !== generation || abort.signal.aborted) return
          url = URL.createObjectURL(blob)
          urls.value = { ...urls.value, [pageIndex + 1]: url }
          url = ""
        } catch {
          // A failed thumbnail must not hide an otherwise usable page.
        } finally {
          if (url) revoke(url)
        }
      }
    }

    try {
      await Promise.all(Array.from({ length: concurrency }, () => worker()))
    } finally {
      requestOptions.signal?.removeEventListener("abort", forwardAbort)
      if (controller === abort) controller = undefined
    }
  }

  function dispose(): void {
    if (disposed) return
    disposed = true
    clear()
  }

  return { urls, load, clear, dispose }
}
