/**
 * DOCX font-metric lifecycle.
 *
 * Every text measurement in the DOCX layout graph is cached under
 * `textMeasurementGeneration()`. This module owns the browser hooks that retire
 * those caches when the font set or the rendering profile changes:
 *
 * - `document.fonts` `loadingdone` / `fonts.ready` (web fonts, embedded faces);
 * - `devicePixelRatio` and `visualViewport.scale` (zoom / monitor changes);
 * - an explicit `invalidateFontMetrics()` call from the host.
 *
 * The generation counter itself lives in `text-measurement-generation.ts` so
 * the pure core graph never imports `document`. `registerFontMetricCache()`
 * lets every advance/metric cache clear itself on the same revision, and
 * `subscribeFontMetrics()` lets pagination or highlight layers re-measure.
 */

import {
  invalidateTextMeasurements,
  onTextMeasurementInvalidated,
  textMeasurementGeneration,
} from "./text-measurement-generation";

export {
  invalidateTextMeasurements,
  onTextMeasurementInvalidated,
  textMeasurementGeneration,
} from "./text-measurement-generation";

const invalidators = new Set<() => void>();
const listeners = new Set<() => void>();

let invalidating = false;
let observedFonts: FontFaceSet | undefined;
let observedProfileWindow: Window | undefined;
let stopObservingProfile: (() => void) | undefined;
let fontsReadyHookInstalled = false;

function notifyListeners(): void {
  for (const listener of [...listeners]) {
    listener();
  }
}

// A plain `invalidateTextMeasurements()` call (the legacy, DOM-free API) must
// clear the same advance caches as `invalidateFontMetrics()`.
onTextMeasurementInvalidated(() => {
  for (const invalidate of [...invalidators]) {
    invalidate();
  }
  if (!invalidating) {
    notifyListeners();
  }
});

/**
 * Retire every cached font metric, advance width and vertical metric. Call
 * after a web font loads, an embedded face is added or removed, or the
 * rendering profile changes. Returns the new measurement generation.
 */
export function invalidateFontMetrics(): number {
  if (invalidating) {
    return invalidateTextMeasurements();
  }

  invalidating = true;
  try {
    const generation = invalidateTextMeasurements();
    notifyListeners();
    return generation;
  } finally {
    invalidating = false;
  }
}

/** Current cache revision. Font/advance caches must include this value. */
export function getFontMetricsRevision(): number {
  observeFonts();
  if (hasDocumentFontSet()) {
    observeRenderingProfile();
  }
  return textMeasurementGeneration();
}

function hasDocumentFontSet(): boolean {
  return (
    typeof document !== "undefined" &&
    Boolean((document as Document & { fonts?: FontFaceSet }).fonts)
  );
}

/** Headless hosts report a stable revision so SSR output is deterministic. */
export function getServerFontMetricsRevision(): number {
  return 0;
}

export function observeFonts(): void {
  const fonts = typeof document !== "undefined" ? document.fonts : undefined;
  if (fonts === observedFonts || !fonts?.addEventListener) {
    return;
  }
  observedFonts?.removeEventListener("loadingdone", invalidateFontMetrics);
  observedFonts = fonts;
  fonts.addEventListener("loadingdone", invalidateFontMetrics);
}

/**
 * Browsers can change the effective glyph advances without loading a font:
 * moving the window to a monitor with another DPR, or zooming the visual
 * viewport, loads optically-sized faces. Word's layout is independent of zoom,
 * but the measurement cache is keyed to the profile that produced it.
 */
function observeRenderingProfile(): void {
  const targetWindow = typeof window !== "undefined" ? window : undefined;
  if (
    !targetWindow ||
    typeof targetWindow.addEventListener !== "function" ||
    targetWindow === observedProfileWindow
  ) {
    return;
  }
  stopObservingProfile?.();
  observedProfileWindow = targetWindow;
  const profile = (): string =>
    `${targetWindow.devicePixelRatio || 1}\u0000${targetWindow.visualViewport?.scale || 1}`;
  let currentProfile = profile();
  let resolutionQuery: MediaQueryList | undefined;
  const removeResolutionListener = (): void => {
    if (resolutionQuery?.removeEventListener) {
      resolutionQuery.removeEventListener("change", handleProfileChange);
    } else {
      resolutionQuery?.removeListener?.(handleProfileChange);
    }
  };
  const watchResolution = (): void => {
    removeResolutionListener();
    resolutionQuery =
      typeof targetWindow.matchMedia === "function"
        ? targetWindow.matchMedia(
            `(resolution: ${targetWindow.devicePixelRatio || 1}dppx)`
          )
        : undefined;
    if (resolutionQuery?.addEventListener) {
      resolutionQuery.addEventListener("change", handleProfileChange);
    } else {
      resolutionQuery?.addListener?.(handleProfileChange);
    }
  };
  function handleProfileChange(): void {
    const nextProfile = profile();
    if (nextProfile === currentProfile) return;
    currentProfile = nextProfile;
    watchResolution();
    invalidateFontMetrics();
  }
  targetWindow.addEventListener("resize", handleProfileChange);
  const viewport = targetWindow.visualViewport;
  viewport?.addEventListener?.("resize", handleProfileChange);
  watchResolution();
  stopObservingProfile = () => {
    targetWindow.removeEventListener?.("resize", handleProfileChange);
    viewport?.removeEventListener?.("resize", handleProfileChange);
    removeResolutionListener();
  };
}

/**
 * Register a cache that must be dropped whenever font metrics are invalidated.
 * The cache is cleared on `document.fonts` load, a rendering-profile change and
 * an explicit `invalidateFontMetrics()` call.
 */
export function registerFontMetricCache(invalidate: () => void): void {
  invalidators.add(invalidate);
  observeFonts();
  if (hasDocumentFontSet()) {
    observeRenderingProfile();
  }
}

/** Subscribe to metric invalidation (e.g. to schedule a re-measure). */
export function subscribeFontMetrics(listener: () => void): () => void {
  observeFonts();
  if (hasDocumentFontSet()) {
    observeRenderingProfile();
  }
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/**
 * Install the `document.fonts` hooks once. Safe to call from every measurement
 * entry point: repeated calls are no-ops. Outside the browser it returns false.
 */
export function ensureDocumentFontInvalidation(): boolean {
  if (typeof document === "undefined") {
    return false;
  }

  const fonts = document.fonts as
    | (FontFaceSet & {
        addEventListener?: (type: string, listener: () => void) => void;
        ready?: Promise<FontFaceSet>;
      })
    | undefined;
  if (!fonts) {
    return false;
  }

  observeFonts();
  // The rendering-profile observer installs window listeners, so it is only
  // wired when a real font set exists. Tests and headless hosts that expose a
  // bare window must not gain a resize listener from a read-only mount.
  observeRenderingProfile();

  if (!fontsReadyHookInstalled) {
    fontsReadyHookInstalled = true;
    const readyPromise = fonts.ready;
    if (readyPromise && typeof readyPromise.then === "function") {
      readyPromise.then(invalidateFontMetrics, () => undefined);
    }
  }

  return true;
}

/**
 * The public `fontMetrics` facade. Hosts can read the revision, force a
 * re-measure and register their own dependent caches without importing the
 * internal cache modules.
 */
export const fontMetrics = {
  revision: getFontMetricsRevision,
  invalidate: invalidateFontMetrics,
  subscribe: subscribeFontMetrics,
  register: registerFontMetricCache,
} as const;
