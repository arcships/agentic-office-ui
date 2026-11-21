/**
 * Deterministic route-capture switches for the verification demo.
 *
 * The L2 route parity harness (`pnpm ui:parity:routes`) opens the demo route
 * with a query string so both sides capture the same state:
 *
 *   `?sample=<file>`  fixture under `/samples/` (default: the page's own)
 *   `?zoom=<number>`  initial PDF zoom, clamped to the viewer's 0.5–2 range
 *   `?statusbar=off`  hide optional chrome the upstream reference does not have
 *
 * The switches never change the shipped default state: without a query the
 * demo behaves exactly as before. See `docs/testing/ui-parity-routes.md`.
 */
export type DemoRouteCapture = {
  sample?: string
  zoom?: number
  statusBar: boolean
}

const MIN_ZOOM = 0.5
const MAX_ZOOM = 2

export function readRouteCapture(): DemoRouteCapture {
  if (typeof window === "undefined") return { statusBar: true }
  const params = new URLSearchParams(window.location.search)
  const sample = params.get("sample")?.trim()
  const zoomParam = Number(params.get("zoom"))
  const zoom =
    Number.isFinite(zoomParam) && params.get("zoom") !== null
      ? Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoomParam))
      : undefined

  return {
    sample: sample ? sample : undefined,
    zoom,
    statusBar: params.get("statusbar") !== "off",
  }
}
