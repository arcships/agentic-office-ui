/**
 * Width measurement with a ResizeObserver.
 *
 * `DocumentViewerSidebar` uses this to decide between its inline and overlay
 * layouts from the width of the viewer shell rather than the viewport: a
 * desktop window can still host a narrow docked preview. Keep the last
 * non-zero measurement when the element is hidden or detached (keep-alive
 * pools, `display: none` ancestors); a zero-width pass would flip the sidebar
 * layout for no reason and replay its open animation when it comes back.
 */

import {
  onBeforeUnmount,
  onMounted,
  ref,
  watch,
  type Ref,
} from "vue"

export interface UseElementWidthOptions {
  /** Treat widths below this value as hidden and keep the previous size. */
  fallback?: number
}

/**
 * Observe an element and return its content width in CSS pixels.
 *
 * The ref starts at `fallback` (0 by default) until the element is mounted and
 * measured. `ResizeObserver` is used when available; older environments fall
 * back to the window `resize` event.
 */
export function useElementWidth(
  target: Ref<HTMLElement | null>,
  options: UseElementWidthOptions = {},
): Ref<number> {
  const width = ref(options.fallback ?? 0)
  let observer: ResizeObserver | null = null
  let listeningToWindow = false
  let mounted = false

  function measure(): void {
    const element = target.value
    if (!element) return
    const next = element.getBoundingClientRect?.().width ?? 0
    if (!Number.isFinite(next) || next <= 0) return
    width.value = next
  }

  function stop(): void {
    observer?.disconnect()
    observer = null
    if (listeningToWindow && typeof window !== "undefined") {
      window.removeEventListener("resize", measure)
      listeningToWindow = false
    }
  }

  function start(): void {
    stop()
    const element = target.value
    if (!element) return
    measure()
    if (typeof ResizeObserver !== "undefined") {
      observer = new ResizeObserver(measure)
      observer.observe(element)
    } else if (typeof window !== "undefined") {
      window.addEventListener("resize", measure)
      listeningToWindow = true
    }
  }

  onMounted(() => {
    mounted = true
    start()
  })

  watch(
    target,
    () => {
      if (mounted) start()
    },
    { flush: "post" },
  )

  onBeforeUnmount(() => {
    mounted = false
    stop()
  })

  return width
}
