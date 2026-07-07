// Generic zoom/scroll utilities.
// Upstream editor.tsx: lines 833-891.

export function clampNumber(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

export function hasVerticalScrollOverflow(value: string): boolean {
  return (
    value === "auto" ||
    value === "scroll" ||
    value === "overlay" ||
    value === "hidden"
  );
}

export function nearestScrollableAncestor(
  element: HTMLElement | null
): HTMLElement | null {
  let current: HTMLElement | null = element;
  while (current) {
    const style = window.getComputedStyle(current);
    if (
      hasVerticalScrollOverflow(style.overflowY) ||
      hasVerticalScrollOverflow(style.overflow)
    ) {
      return current;
    }
    current = current.parentElement;
  }
  const scrollingElement =
    typeof document !== "undefined" ? document.scrollingElement : null;
  if (scrollingElement instanceof HTMLElement) {
    return scrollingElement;
  }
  return typeof document !== "undefined" ? document.documentElement : null;
}

export function resolveEffectiveZoomScale(element: HTMLElement): number {
  let current: HTMLElement | null = element;
  let scale = 1;
  while (current) {
    const zoomRaw = window.getComputedStyle(current).zoom;
    if (zoomRaw && zoomRaw !== "normal") {
      const zoom = Number.parseFloat(zoomRaw);
      if (Number.isFinite(zoom) && zoom > 0) {
        scale *= zoom;
      }
    }
    current = current.parentElement;
  }

  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

export function normalizePageVirtualizationZoomScale(
  value: unknown
): number | undefined {
  const scale = Number(value);
  return Number.isFinite(scale) && scale > 0 ? scale : undefined;
}
