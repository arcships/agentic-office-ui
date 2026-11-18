import type { InjectionKey } from "vue"

/**
 * Shared laziness gate for slide thumbnails.
 *
 * One `IntersectionObserver` serves every thumbnail below the provider, so a
 * long deck no longer creates one observer per mounted thumbnail. Observers are
 * created on first use and released through `dispose()`.
 */
export interface PptxThumbnailObserver {
  observe(element: Element, onVisible: () => void): void
  unobserve(element: Element): void
  dispose(): void
}

export interface PptxThumbnailObserverOptions {
  /** Distance before the scrollport where a thumbnail starts rendering. */
  rootMargin?: string
}

export const PPTX_THUMBNAIL_OBSERVER: InjectionKey<PptxThumbnailObserver> =
  Symbol("pptx-thumbnail-observer")

export function createPptxThumbnailObserver(
  options: PptxThumbnailObserverOptions = {},
): PptxThumbnailObserver {
  const callbacks = new Map<Element, () => void>()
  const rootMargin = options.rootMargin ?? "240px"
  let observer: IntersectionObserver | null = null
  let disposed = false

  function intersectionObserver(): IntersectionObserver | null {
    if (observer) return observer
    if (disposed || typeof IntersectionObserver === "undefined") return null
    observer = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue
        const callback = callbacks.get(entry.target)
        if (!callback) continue
        callbacks.delete(entry.target)
        observer?.unobserve(entry.target)
        callback()
      }
    }, { rootMargin })
    return observer
  }

  return {
    observe(element, onVisible) {
      const target = intersectionObserver()
      if (!target) {
        onVisible()
        return
      }
      callbacks.set(element, onVisible)
      target.observe(element)
    },
    unobserve(element) {
      callbacks.delete(element)
      observer?.unobserve(element)
    },
    dispose() {
      disposed = true
      callbacks.clear()
      observer?.disconnect()
      observer = null
    },
  }
}
