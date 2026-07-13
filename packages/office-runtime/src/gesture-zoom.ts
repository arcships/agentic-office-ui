export const SURFACE_ZOOM_MIN = 0.5;
export const SURFACE_ZOOM_MAX = 2;
export const SURFACE_ZOOM_PRECISION = 100;

export function clampSurfaceZoom(value: number): number {
  if (!Number.isFinite(value)) return 1;
  const clamped = Math.min(SURFACE_ZOOM_MAX, Math.max(SURFACE_ZOOM_MIN, value));
  return Math.round(clamped * SURFACE_ZOOM_PRECISION) / SURFACE_ZOOM_PRECISION;
}

export function normalizeWheelDelta(deltaY: number, deltaMode = 0): number {
  if (!Number.isFinite(deltaY)) return 0;
  if (deltaMode === 1) return deltaY * 16;
  if (deltaMode === 2) return deltaY * 800;
  return deltaY;
}

export function nextSurfaceZoom(currentZoom: number, deltaY: number, deltaMode = 0): number {
  const current = clampSurfaceZoom(currentZoom);
  const delta = normalizeWheelDelta(deltaY, deltaMode);
  const next = clampSurfaceZoom(current * Math.exp(-delta * 0.005));
  if (delta === 0 || next !== current) return next;
  return clampSurfaceZoom(current + (delta < 0 ? 0.01 : -0.01));
}
