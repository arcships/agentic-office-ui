/**
 * Deterministic route-capture switches for the verification demo.
 *
 * The L2 route parity harness (`pnpm ui:parity:routes`) opens the demo route
 * with a query string so both sides capture the same state:
 *
 *   `?sample=<file>`   fixture under `/samples/` (default: the page's own)
 *   `?zoom=<number>`   initial PDF zoom, clamped to the viewer's 0.5–2 range
 *   `?locale=<id>`     viewer locale for the capture (`en` matches the upstream
 *                      reference; the demo page otherwise uses `zh-CN`)
 *   `?statusbar=on`    opt into the optional status bar; the upstream reference
 *                      viewer has no status bar, so the demo default is off
 *
 * The switches only override the shipped defaults; without a query the demo
 * uses the upstream-aligned defaults (sidebar closed, no status bar, zh-CN
 * copy). See `docs/testing/ui-parity-routes.md`.
 */
export type DemoRouteCapture = {
  sample?: string
  zoom?: number
  locale?: "en" | "zh-CN"
  statusBar: boolean
}

const MIN_ZOOM = 0.5
const MAX_ZOOM = 2

export function readRouteCapture(): DemoRouteCapture {
  if (typeof window === "undefined") return { statusBar: false }
  const params = new URLSearchParams(window.location.search)
  const sample = params.get("sample")?.trim()
  const zoomParam = Number(params.get("zoom"))
  const zoom =
    Number.isFinite(zoomParam) && params.get("zoom") !== null
      ? Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoomParam))
      : undefined
  const localeParam = params.get("locale")
  const locale =
    localeParam === "en" || localeParam === "zh-CN" ? localeParam : undefined

  return {
    sample: sample ? sample : undefined,
    zoom,
    locale,
    statusBar: params.get("statusbar") === "on",
  }
}
