// Clipped body-rect measurement for page-height feedback.
//
// Migration source: @extend-ai/react-docx commit 6f70b92,
// react-viewer/src/editor.tsx body-measurement.ts lines 13-77
// (`createClippedBodyRectReader`) feeding `resolveMeasuredBodyRenderedBottomPx`
// (editor.tsx:43676-43686, ported in page-measurement.ts).
//
// Page content height is resolved from the rendered descendants of the page
// body. A descendant that is cut off by an `overflow: hidden` / `overflow: clip`
// ancestor must only contribute the part the reader can see: a fixed-height
// table cell that clips overflowing paragraphs would otherwise report a taller
// body, which inflates the measured page height and pushes pagination away from
// Word. The reader intersects every clipping ancestor's border box, caches the
// result per element for the lifetime of the measurement pass, and converts the
// screen-space rects back to page-local layout pixels so callers can compare
// them with `offsetHeight` at any zoom level.

import { PAGE_CONTENT_MEASUREMENT_IGNORE_ATTRIBUTE } from "./constants";
import { resolveMeasuredBodyRenderedBottomPx } from "./page-measurement";

// `overflow` values whose box clips its descendants. `auto`/`scroll`/`overlay`
// are included because their scroll viewport is still the visible area.
const CLIPPING_OVERFLOW_VALUES = new Set([
  "auto",
  "clip",
  "hidden",
  "overlay",
  "scroll"
]);

/** A rendered box after every clipping ancestor has been applied. */
export interface MeasuredBodyRect {
  /** Distance from the measured root's top edge, in layout pixels. */
  topPx: number;
  /** Distance from the measured root's left edge, in layout pixels. */
  leftPx: number;
  /** Distance from the measured root's top edge to the visible right edge. */
  rightPx: number;
  /** Visible bottom, in layout pixels from the measured root's top edge. */
  bottomPx: number;
  widthPx: number;
  heightPx: number;
}

export interface ClippedBodyRectReaderOptions {
  /**
   * Zoom applied to the measured subtree. Screen-space rects are divided by it
   * so the reader reports unscaled layout pixels. Defaults to 1.
   */
  zoomFactor?: number;
}

export type ClippedBodyRectReader = (
  element: HTMLElement | null | undefined
) => MeasuredBodyRect | undefined;

interface BoxLike {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

/** Clip box for one element plus whether it is inside the measured root. */
interface ClipResult {
  clip: BoxLike | undefined;
  anchored: boolean;
}

interface ElementOverflowStyle {
  overflow?: string | null;
  overflowX?: string | null;
  overflowY?: string | null;
}

interface RectLike {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
  width?: number;
  height?: number;
}

function normalizeZoomFactor(value: unknown): number {
  const zoom = Number(value);
  return Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
}

function finiteNumber(value: unknown): number | undefined {
  if (value === null || value === undefined || value === "") {
    return undefined;
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : undefined;
}

function readBorderBox(element: HTMLElement | null | undefined): BoxLike | undefined {
  if (!element || typeof element.getBoundingClientRect !== "function") {
    return undefined;
  }
  let rect: RectLike | null | undefined;
  try {
    rect = element.getBoundingClientRect() as RectLike | null;
  } catch {
    return undefined;
  }
  if (!rect) {
    return undefined;
  }
  const left = finiteNumber(rect.left);
  const top = finiteNumber(rect.top);
  const width = finiteNumber(rect.width);
  const height = finiteNumber(rect.height);
  const right =
    finiteNumber(rect.right) ??
    (left !== undefined && width !== undefined ? left + width : undefined);
  const bottom =
    finiteNumber(rect.bottom) ??
    (top !== undefined && height !== undefined ? top + height : undefined);
  if (left === undefined || top === undefined || right === undefined || bottom === undefined) {
    return undefined;
  }
  return { top, right, bottom, left };
}

function intersectBox(
  box: BoxLike | undefined,
  clip: BoxLike | undefined
): BoxLike | undefined {
  // No box yet means nothing constrains it, so the clip becomes the box.
  if (!box) {
    return clip;
  }
  if (!clip) {
    return box;
  }
  // Fully clipped boxes collapse onto the clip edge they fell out of. Clamping
  // to their own top/left instead would report a bottom outside the clip and
  // inflate the measurement with content that is not visible at all.
  const left = Math.max(Math.min(box.left, clip.right), clip.left);
  const right = Math.min(Math.max(box.right, clip.left), clip.right);
  const top = Math.max(Math.min(box.top, clip.bottom), clip.top);
  const bottom = Math.min(Math.max(box.bottom, clip.top), clip.bottom);
  return { top, right, bottom, left };
}

function readOverflowStyle(
  element: HTMLElement
): ElementOverflowStyle | undefined {
  const getComputedStyleImpl = (
    globalThis as {
      getComputedStyle?: (target: Element) => ElementOverflowStyle | null;
    }
  ).getComputedStyle;
  if (typeof getComputedStyleImpl === "function") {
    try {
      const style = getComputedStyleImpl.call(globalThis, element);
      if (style) {
        return style;
      }
    } catch {
      // No layout engine (SSR, tests): fall back to the inline style.
    }
  }
  return (
    (element as unknown as { style?: ElementOverflowStyle | null }).style ??
    undefined
  );
}

function isClippingOverflow(value: string | null | undefined): boolean {
  return CLIPPING_OVERFLOW_VALUES.has(
    String(value ?? "").trim().toLowerCase()
  );
}

/**
 * True when the element's own box clips its descendants. Computed overflow is
 * preferred; tests and non-rendering environments fall back to `element.style`.
 */
export function elementClipsOverflow(
  element: HTMLElement | null | undefined
): boolean {
  if (!element) {
    return false;
  }
  const style = readOverflowStyle(element);
  if (!style) {
    return false;
  }
  return (
    isClippingOverflow(style.overflow) ||
    isClippingOverflow(style.overflowX) ||
    isClippingOverflow(style.overflowY)
  );
}

function parentElementOf(element: HTMLElement): HTMLElement | undefined {
  const parent = (
    element as unknown as { parentElement?: HTMLElement | null }
  ).parentElement;
  return parent ?? undefined;
}

function toLocalRect(
  box: BoxLike,
  origin: BoxLike,
  zoomFactor: number
): MeasuredBodyRect {
  const leftPx = (box.left - origin.left) / zoomFactor;
  const topPx = (box.top - origin.top) / zoomFactor;
  const rightPx = (box.right - origin.left) / zoomFactor;
  const bottomPx = (box.bottom - origin.top) / zoomFactor;
  return {
    topPx,
    leftPx,
    rightPx,
    bottomPx,
    widthPx: Math.max(0, rightPx - leftPx),
    heightPx: Math.max(0, bottomPx - topPx)
  };
}

/**
 * Build a reader that resolves the visible (clipped) rect of an element inside
 * `root`, in `root`-local layout pixels.
 *
 * The ancestor walk stops at `root`: the page body's own overflow clips its
 * content, but the viewer's scroll container (which also scrolls and clips)
 * must never leak into the measurement, or a page's height would depend on the
 * reader's scroll position.
 */
export function createClippedBodyRectReader(
  root: HTMLElement,
  options: ClippedBodyRectReaderOptions = {}
): ClippedBodyRectReader {
  const zoomFactor = normalizeZoomFactor(options?.zoomFactor);
  const originBox = readBorderBox(root);
  const rectCache = new WeakMap<HTMLElement, MeasuredBodyRect | undefined>();
  const clipCache = new WeakMap<HTMLElement, ClipResult>();
  const overflowCache = new WeakMap<HTMLElement, boolean>();

  function clips(element: HTMLElement): boolean {
    const cached = overflowCache.get(element);
    if (cached !== undefined) {
      return cached;
    }
    const clipsOverflow = elementClipsOverflow(element);
    overflowCache.set(element, clipsOverflow);
    return clipsOverflow;
  }

  // Clip box constraining `element`'s own border box, contributed by its
  // ancestors up to (and including) `root`. `anchored` is false for elements
  // that are not inside the measured root, which are never measurable.
  function clipBoxFor(element: HTMLElement): ClipResult {
    const cached = clipCache.get(element);
    if (cached) {
      return cached;
    }
    const parent = parentElementOf(element);
    if (!parent) {
      const detached: ClipResult = { clip: undefined, anchored: false };
      clipCache.set(element, detached);
      return detached;
    }
    const inherited = parent === root
      ? { clip: undefined as BoxLike | undefined, anchored: true }
      : clipBoxFor(parent);
    let clip = inherited.clip;
    if (clips(parent)) {
      const parentBox = readBorderBox(parent);
      if (parentBox) {
        clip = intersectBox(clip, parentBox);
      }
    }
    const result: ClipResult = { clip, anchored: inherited.anchored };
    clipCache.set(element, result);
    return result;
  }

  return (element) => {
    if (!element || element === root) {
      return undefined;
    }
    if (rectCache.has(element)) {
      return rectCache.get(element);
    }
    const box = readBorderBox(element);
    const clip = clipBoxFor(element);
    let measured: MeasuredBodyRect | undefined;
    if (box && originBox && clip.anchored) {
      const visible = intersectBox(box, clip.clip);
      if (visible) {
        measured = toLocalRect(visible, originBox, zoomFactor);
      }
    }
    rectCache.set(element, measured);
    return measured;
  };
}

/**
 * Collect the measurable descendants of `root`, pruning subtrees marked with
 * `data-docx-pagination-ignore` (overlays and other chrome that must not
 * contribute to page content height).
 */
export function collectPageContentMeasurementDescendants(
  root: HTMLElement | null | undefined
): HTMLElement[] {
  const descendants: HTMLElement[] = [];
  if (!root) {
    return descendants;
  }
  const visit = (parent: HTMLElement): void => {
    const children = (
      parent as unknown as { children?: ArrayLike<Element> | null }
    ).children;
    if (!children) {
      return;
    }
    for (let index = 0; index < children.length; index += 1) {
      const child = children[index] as unknown as HTMLElement | null;
      if (!child || typeof child.getBoundingClientRect !== "function") {
        continue;
      }
      if (
        typeof child.hasAttribute === "function" &&
        child.hasAttribute(PAGE_CONTENT_MEASUREMENT_IGNORE_ATTRIBUTE)
      ) {
        continue;
      }
      descendants.push(child);
      visit(child);
    }
  };
  visit(root);
  return descendants;
}

/**
 * Visible bottom of everything rendered inside `root`, in `root`-local layout
 * pixels. Descendants clipped by the root, a table cell, a column or any other
 * clipping ancestor only contribute the part that is actually visible.
 */
export function resolveClippedBodyRenderedBottomPx(
  root: HTMLElement | null | undefined,
  options: ClippedBodyRectReaderOptions = {}
): number | undefined {
  if (!root) {
    return undefined;
  }
  const reader = createClippedBodyRectReader(root, options);
  return resolveMeasuredBodyRenderedBottomPx(
    collectPageContentMeasurementDescendants(root).map((element) => {
      const rect = reader(element);
      return rect
        ? {
            bottomPx: rect.bottomPx,
            widthPx: rect.widthPx,
            heightPx: rect.heightPx
          }
        : { bottomPx: Number.NaN, widthPx: 0, heightPx: 0 };
    })
  );
}

/**
 * Resolve the measured height of a page box.
 *
 * The clipped content bottom wins whenever the DOM could be measured; the box's
 * `offsetHeight` (or another caller-supplied height) is only the fallback for a
 * box without measurable descendants. `minimumHeightPx` keeps a page at least
 * one paper tall even when its content is shorter.
 */
export function resolveClippedMeasuredHeightPx(params: {
  fallbackHeightPx?: number;
  minimumHeightPx?: number;
  clippedContentBottomPx?: number;
}): number | undefined {
  const minimumHeightPx = Math.max(
    0,
    finiteNumber(params.minimumHeightPx) ?? 0
  );
  const clippedContentBottomPx = finiteNumber(params.clippedContentBottomPx);
  if (clippedContentBottomPx !== undefined && clippedContentBottomPx > 0) {
    return Math.round(Math.max(minimumHeightPx, clippedContentBottomPx));
  }
  const fallbackHeightPx = finiteNumber(params.fallbackHeightPx);
  if (fallbackHeightPx !== undefined && fallbackHeightPx > 0) {
    return Math.round(Math.max(minimumHeightPx, fallbackHeightPx));
  }
  return minimumHeightPx > 0 ? Math.round(minimumHeightPx) : undefined;
}
