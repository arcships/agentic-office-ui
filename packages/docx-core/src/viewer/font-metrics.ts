/**
 * Text-measurement generation shared by the DOCX layout caches.
 *
 * Width/height estimations are cached per paragraph/run. A cache entry taken
 * before a web font finishes loading is wrong forever unless the caches know
 * that the font set changed; the generation counter is part of every
 * measurement cache key, and `document.fonts` events bump it automatically.
 */

let textMeasurementGenerationCounter = 0;
let documentFontInvalidationInstalled = false;
const textMeasurementInvalidationListeners = new Set<() => void>();

/** Current measurement generation. Cache keys must include this value. */
export function textMeasurementGeneration(): number {
  return textMeasurementGenerationCounter;
}

/**
 * Invalidate every cached text measurement. Call after a web font loads or the
 * font set otherwise changes. Returns the new generation.
 */
export function invalidateTextMeasurements(): number {
  textMeasurementGenerationCounter += 1;
  for (const listener of [...textMeasurementInvalidationListeners]) {
    listener();
  }
  return textMeasurementGenerationCounter;
}

/**
 * Subscribe to measurement invalidation (e.g. to drop a caller-owned cache or
 * re-run pagination). Returns an unsubscribe function.
 */
export function onTextMeasurementInvalidated(
  listener: () => void
): () => void {
  textMeasurementInvalidationListeners.add(listener);
  return () => {
    textMeasurementInvalidationListeners.delete(listener);
  };
}

/**
 * Install the `document.fonts` hooks once. Safe to call from every measurement
 * entry point: repeated calls are no-ops. Outside the browser it returns false.
 */
export function ensureDocumentFontInvalidation(): boolean {
  if (documentFontInvalidationInstalled) {
    return true;
  }

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

  documentFontInvalidationInstalled = true;
  const invalidate = (): void => {
    invalidateTextMeasurements();
  };

  if (typeof fonts.addEventListener === "function") {
    fonts.addEventListener("loadingdone", invalidate);
  }

  // `fonts.ready` resolves once the initially requested faces have loaded. It
  // is a fresh promise per spec revision, but resolving it again after
  // `loadingdone` is harmless (the invalidation is idempotent for caches).
  const readyPromise = fonts.ready;
  if (readyPromise && typeof readyPromise.then === "function") {
    readyPromise.then(invalidate, () => undefined);
  }

  return true;
}
