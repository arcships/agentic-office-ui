/**
 * Text-measurement generation shared by the DOCX layout caches.
 *
 * This module is intentionally platform-free: the `document.fonts` hooks live
 * in `font-metrics.ts`, which browser hosts call once. Keeping the counter in
 * a closure (rather than module-level mutable bindings) keeps it inside the
 * pure core graph.
 */

const measurementState = (() => {
  let generation = 0;
  const listeners = new Set<() => void>();
  return {
    add(listener: () => void): () => void {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    bump(): number {
      generation += 1;
      for (const listener of [...listeners]) {
        listener();
      }
      return generation;
    },
    current(): number {
      return generation;
    },
  };
})();

/** Current measurement generation. Cache keys must include this value. */
export function textMeasurementGeneration(): number {
  return measurementState.current();
}

/**
 * Invalidate every cached text measurement. Call after a web font loads or the
 * font set otherwise changes. Returns the new generation.
 */
export function invalidateTextMeasurements(): number {
  return measurementState.bump();
}

/**
 * Subscribe to measurement invalidation (e.g. to drop a caller-owned cache or
 * re-run pagination). Returns an unsubscribe function.
 */
export function onTextMeasurementInvalidated(
  listener: () => void
): () => void {
  return measurementState.add(listener);
}
