/**
 * Browser hooks for text-measurement invalidation.
 *
 * The generation counter itself lives in `text-measurement-generation.ts` so
 * the pure core graph never imports `document`. Hosts call
 * `ensureDocumentFontInvalidation()` once before layout.
 */

import { invalidateTextMeasurements } from "./text-measurement-generation";

export {
  invalidateTextMeasurements,
  onTextMeasurementInvalidated,
  textMeasurementGeneration,
} from "./text-measurement-generation";

let documentFontInvalidationInstalled = false;


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
