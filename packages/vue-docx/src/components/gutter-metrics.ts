// Screen-space metrics for the DOCX annotation gutter.
//
// The gutter sits beside the paper, inside the subtree that applies
// `transform: scale(zoomFactor)` to the page. Word's annotation pane does not
// scale with the document — at 50% zoom its text stays 12px and the pane stays
// 240px wide — so the pane counter-scales itself by `1 / zoomFactor` and every
// value here is expressed in *screen* px: what the user sees at any zoom.
//
// Page-space values (card `top`, anchor `y`) come from
// `@arcships/docx-core`'s layout pipeline, are measured against the page's
// untransformed box, and are converted to screen px at the edges of this module.

export const DOCX_ANNOTATION_GUTTER_WIDTH_PX = 240
export const DOCX_ANNOTATION_GUTTER_GAP_PX = 16
export const DOCX_ANNOTATION_GUTTER_MIN_ZOOM = 0.25
export const DOCX_ANNOTATION_GUTTER_MAX_ZOOM = 2
/**
 * Container width below which the gutter becomes an overlay drawer.
 *
 * Above it the pane sits beside the paper in its own column (Word's annotation
 * pane). Below it there is no room for the paper plus a 240px column, so the
 * pane slides over the paper instead of pushing it out of the viewport — the
 * old behaviour dropped the pane into the page flow and inflated page height.
 */
export const DOCX_ANNOTATION_GUTTER_DRAWER_MAX_WIDTH_PX = 1024
/** Width of the drawer panel, in screen px. */
export const DOCX_ANNOTATION_GUTTER_DRAWER_WIDTH_PX = 280
/** Width of the closed drawer's tab, in screen px. */
export const DOCX_ANNOTATION_GUTTER_DRAWER_TAB_WIDTH_PX = 28

/**
 * Clamps a viewer zoom factor to the range the gutter can counter-scale.
 * Invalid or non-positive input falls back to 1 (100%).
 */
export function normalizeAnnotationGutterZoom(
  zoomFactor: number | undefined
): number {
  if (!Number.isFinite(zoomFactor) || (zoomFactor as number) <= 0) {
    return 1
  }

  return Math.min(
    DOCX_ANNOTATION_GUTTER_MAX_ZOOM,
    Math.max(DOCX_ANNOTATION_GUTTER_MIN_ZOOM, zoomFactor as number)
  )
}

/**
 * Horizontal footprint the gutter occupies beside the paper, in screen px.
 * Constant across zoom levels because the pane counter-scales; the viewer
 * reserves exactly this much room so the paper stays centered.
 */
export function annotationGutterReservedWidthPx(): number {
  return DOCX_ANNOTATION_GUTTER_GAP_PX + DOCX_ANNOTATION_GUTTER_WIDTH_PX
}

function roundToHundredth(value: number): number {
  return Math.round(value * 100) / 100
}

/**
 * Style for the counter-scaled gutter pane.
 *
 * `left: calc(100% + gap / zoom)` keeps the visible gap constant in screen px,
 * and `height` is the page's *screen* height so the pane covers exactly the
 * paper at every zoom level (it is a child of the scaled page).
 */
export function annotationGutterPaneStyle(
  zoomFactor: number | undefined,
  pageHeightPx: number
): Record<string, string> {
  const zoom = normalizeAnnotationGutterZoom(zoomFactor)
  const pageHeight = Number.isFinite(pageHeightPx) ? Math.max(0, pageHeightPx) : 0

  return {
    left: `calc(100% + ${roundToHundredth(
      DOCX_ANNOTATION_GUTTER_GAP_PX / zoom
    )}px)`,
    top: "0px",
    width: `${DOCX_ANNOTATION_GUTTER_WIDTH_PX}px`,
    height: `${Math.round(pageHeight * zoom)}px`,
    transform: `scale(${roundToHundredth(1 / zoom)})`,
    transformOrigin: "left top",
  }
}

/**
 * Style for the counter-scaled gutter pane in drawer mode.
 *
 * The drawer is anchored to the page's right edge and slides over the paper
 * instead of occupying a column beside it, so the viewer reserves no horizontal
 * footprint and the paper keeps the full viewport width.
 */
export function annotationGutterDrawerPaneStyle(
  zoomFactor: number | undefined,
  pageHeightPx: number,
  drawerOpen = true
): Record<string, string> {
  const zoom = normalizeAnnotationGutterZoom(zoomFactor)
  const pageHeight = Number.isFinite(pageHeightPx) ? Math.max(0, pageHeightPx) : 0

  return {
    left: "auto",
    right: "0px",
    top: "0px",
    width: `${drawerOpen ? DOCX_ANNOTATION_GUTTER_DRAWER_WIDTH_PX : DOCX_ANNOTATION_GUTTER_DRAWER_TAB_WIDTH_PX}px`,
    height: `${Math.round(pageHeight * zoom)}px`,
    transform: `scale(${roundToHundredth(1 / zoom)})`,
    transformOrigin: "right top",
  }
}

/** Whether the gutter should render as an overlay drawer at this width. */
export function annotationGutterUsesDrawer(
  containerWidthPx: number | undefined
): boolean {
  if (!Number.isFinite(containerWidthPx)) {
    return false
  }
  return (containerWidthPx as number) < DOCX_ANNOTATION_GUTTER_DRAWER_MAX_WIDTH_PX
}

/**
 * Converts a card's page-space `top` into the pane's screen-space `top`.
 * Page space is unscaled; the counter-scaled pane is 1:1 with screen px, so
 * the conversion is a plain multiplication by the zoom factor.
 */
export function annotationGutterCardTopPx(
  pageTopPx: number,
  zoomFactor: number | undefined
): number {
  const zoom = normalizeAnnotationGutterZoom(zoomFactor)
  const top = Number.isFinite(pageTopPx) ? pageTopPx : 0
  return roundToHundredth(top * zoom)
}

/**
 * Marker chip placed on the anchor inside the page (page space). The chip
 * counter-scales around its own left-center so it keeps a constant screen size
 * while staying centered on the anchor's line at any zoom.
 */
export function annotationGutterMarkerStyle(
  anchorXPx: number,
  anchorYPx: number,
  zoomFactor: number | undefined,
  markerSizePx = 14
): Record<string, string> {
  const zoom = normalizeAnnotationGutterZoom(zoomFactor)
  const x = Number.isFinite(anchorXPx) ? anchorXPx : 0
  const y = Number.isFinite(anchorYPx) ? anchorYPx : 0

  return {
    left: `${roundToHundredth(x)}px`,
    top: `${roundToHundredth(y - markerSizePx / 2)}px`,
    transform: `scale(${roundToHundredth(1 / zoom)})`,
    transformOrigin: "left center",
  }
}

/**
 * Leader line from the anchor to the gutter pane (page space). The line runs
 * from the anchor to the pane's left edge, so its width includes the pane gap
 * expressed in page space (`gap / zoom`).
 */
export function annotationGutterLeaderStyle(
  startXPx: number,
  anchorYPx: number,
  pageWidthPx: number,
  zoomFactor: number | undefined
): Record<string, string> {
  const zoom = normalizeAnnotationGutterZoom(zoomFactor)
  const startX = Number.isFinite(startXPx) ? startXPx : 0
  const pageWidth = Number.isFinite(pageWidthPx) ? pageWidthPx : 0
  const spanPx =
    Math.max(0, pageWidth - startX) + DOCX_ANNOTATION_GUTTER_GAP_PX / zoom

  return {
    left: `${roundToHundredth(startX)}px`,
    top: `${roundToHundredth(anchorYPx)}px`,
    width: `${roundToHundredth(spanPx)}px`,
    height: "1px",
  }
}
