// Page-space layout for the DOCX annotation gutter.
//
// The gutter's cards are positioned against their anchors, not against a fixed
// offset at the top of the page. Positioning happens in two stages:
//
//   1. Resolution (DOM): find each annotation's inline anchor element, falling
//      back to its paragraph / table cell, and measure it against the page
//      surface's *untransformed* box. The result is page-space geometry that
//      does not change when the viewer zooms.
//   2. Layout (pure): `layoutTrackedChangesForPage` orders cards by anchor
//      position, stacks them with collision avoidance and derives the card
//      tops. No DOM access, so it behaves identically in tests and in a browser.
//
// Measured card heights feed back into stage 2 so long comments reserve the
// room they actually need. Heights are measured in screen px (the pane is
// counter-scaled) and converted back to page space before layout.

import {
  computed,
  onMounted,
  onUnmounted,
  shallowRef,
  watch,
  type Ref,
} from "vue"
import {
  layoutTrackedChangesForPage,
  resolveGutterAnnotationAnchorsInPage,
  type DocxGutterAnnotation,
  type PositionedGutterAnnotation,
  type TrackedChangeAnchorPoint,
} from "@arcships/docx-core"
import { normalizeAnnotationGutterZoom } from "../components/gutter-metrics"

export interface DocxGutterLayoutOptions {
  /** Page surface element the anchors are measured against. */
  pageElement: Ref<HTMLElement | undefined>
  /** Annotations to place, in document order. */
  annotations: () => readonly DocxGutterAnnotation[]
  /** Page size in page space (unscaled CSS px). */
  pageWidthPx: () => number
  pageHeightPx: () => number
  /** Viewer zoom factor applied to the page subtree. */
  zoomFactor: () => number
  /**
   * Any value that changes when paragraph geometry may have moved (model,
   * page segments, …). Changing it schedules a fresh anchor measurement.
   */
  layoutSignal?: () => unknown
}

export interface DocxGutterLayout {
  /** Cards with page-space `top`/`heightPx`, ordered by anchor position. */
  positioned: Ref<PositionedGutterAnnotation[]>
  zoomFactor: Ref<number>
  pageWidthPx: Ref<number>
  pageHeightPx: Ref<number>
  /** Receives a rendered card's height in screen px. */
  reportCardHeight: (annotationId: string, heightPx: number) => void
  /** Re-measures anchors (content, layout or page-height change). */
  refresh: () => void
}

function anchorPointsEqual(
  left: Map<string, TrackedChangeAnchorPoint>,
  right: Map<string, TrackedChangeAnchorPoint>
): boolean {
  if (left.size !== right.size) return false
  for (const [id, point] of left) {
    const other = right.get(id)
    if (!other || other.x !== point.x || other.y !== point.y) return false
  }

  return true
}

export function useDocxGutterLayout(
  options: DocxGutterLayoutOptions
): DocxGutterLayout {
  const anchorPoints = shallowRef<Map<string, TrackedChangeAnchorPoint>>(
    new Map()
  )
  const cardHeightsScreenPx = shallowRef<Map<string, number>>(new Map())

  const zoomFactor = computed(() =>
    normalizeAnnotationGutterZoom(options.zoomFactor())
  )
  const pageWidthPx = computed(() => Math.max(0, options.pageWidthPx() || 0))
  const pageHeightPx = computed(() => Math.max(0, options.pageHeightPx() || 0))

  const positioned = computed<PositionedGutterAnnotation[]>(() => {
    const zoom = zoomFactor.value
    const pageSpaceHeights = new Map<string, number>()
    cardHeightsScreenPx.value.forEach((heightPx, id) => {
      pageSpaceHeights.set(id, heightPx / zoom)
    })

    return layoutTrackedChangesForPage(
      options.annotations(),
      anchorPoints.value,
      pageSpaceHeights,
      pageWidthPx.value,
      pageHeightPx.value
    )
  })

  let measurementScheduled = false
  let disposed = false
  let scheduledHandle: ReturnType<typeof setTimeout> | undefined

  function measureAnchors(): void {
    const pageElement = options.pageElement.value
    // Non-browser renderers (tests, SSR) expose no layout box or selector
    // engine; the layout pipeline then falls back to document-order anchors.
    if (
      !pageElement ||
      typeof pageElement.getBoundingClientRect !== "function" ||
      typeof pageElement.querySelector !== "function"
    ) {
      return
    }

    const next = resolveGutterAnnotationAnchorsInPage(
      pageElement,
      options.annotations()
    )
    if (anchorPointsEqual(anchorPoints.value, next)) {
      return
    }

    anchorPoints.value = next
  }

  function refresh(): void {
    if (disposed) return
    if (measurementScheduled) return
    measurementScheduled = true

    const run = (): void => {
      measurementScheduled = false
      if (!disposed) measureAnchors()
    }

    if (typeof globalThis.requestAnimationFrame === "function") {
      globalThis.requestAnimationFrame(run)
      return
    }

    scheduledHandle = setTimeout(run, 0)
  }

  function reportCardHeight(annotationId: string, heightPx: number): void {
    if (!Number.isFinite(heightPx) || heightPx <= 0) return
    const previous = cardHeightsScreenPx.value.get(annotationId)
    if (previous !== undefined && Math.abs(previous - heightPx) < 0.5) return

    const next = new Map(cardHeightsScreenPx.value)
    next.set(annotationId, heightPx)
    cardHeightsScreenPx.value = next
  }

  // Deliberately no window/document listener: a read-only DOCX surface must not
  // own global listeners. Reflows are signalled through `refresh` by the page
  // surface's own ResizeObserver, `layoutSignal` and zoom changes.

  onMounted(refresh)
  watch(
    () => [options.annotations(), options.layoutSignal?.()] as const,
    refresh,
    { flush: "post" }
  )
  // Zooming does not move anchors (page space is unscaled) but it changes how
  // much page height is visible, which bounds card stacking.
  watch(zoomFactor, refresh, { flush: "post" })

  onUnmounted(() => {
    disposed = true
    if (scheduledHandle !== undefined) {
      clearTimeout(scheduledHandle)
      scheduledHandle = undefined
    }
  })

  return {
    positioned,
    zoomFactor,
    pageWidthPx,
    pageHeightPx,
    reportCardHeight,
    refresh,
  }
}
