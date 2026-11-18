/**
 * Display scale tracking shared by the PDF surfaces.
 *
 * `devicePixelRatio` read once at render time freezes the bitmap scale: moving
 * the window to a display with a different scale factor never re-renders, and
 * high-density displays stay down-sampled. A `resolution` media query is the
 * only signal the platform emits when that ratio changes.
 */

/** Highest pixel ratio either surface renders at. */
export const MAX_RENDER_DPR = 3

export function clampRenderDpr(value: number): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) return 1
  return Math.min(MAX_RENDER_DPR, Math.max(1, value))
}

/** Pixel ratio the current display should render at. */
export function currentRenderDpr(): number {
  return clampRenderDpr(typeof devicePixelRatio === "number" ? devicePixelRatio : 1)
}

/**
 * Call `onChange` whenever the window's display scale changes. Returns a
 * disposer. Environments without `matchMedia` (SSR, tests) report no changes.
 */
export function watchDevicePixelRatio(onChange: (dpr: number) => void): () => void {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") return () => {}
  let disposed = false
  let unlisten: (() => void) | undefined

  const subscribe = (): void => {
    unlisten?.()
    unlisten = undefined
    if (disposed) return
    let query: MediaQueryList
    try {
      query = window.matchMedia(`(resolution: ${currentRenderDpr()}dppx)`)
    } catch {
      return
    }
    const handleChange = (): void => {
      if (disposed) return
      const next = currentRenderDpr()
      // The query string has to follow the ratio, so re-subscribe every time.
      subscribe()
      onChange(next)
    }
    if (typeof query.addEventListener === "function") {
      query.addEventListener("change", handleChange)
      unlisten = () => query.removeEventListener("change", handleChange)
    } else if (typeof query.addListener === "function") {
      query.addListener(handleChange)
      unlisten = () => query.removeListener?.(handleChange)
    }
  }

  subscribe()
  return () => {
    disposed = true
    unlisten?.()
    unlisten = undefined
  }
}
