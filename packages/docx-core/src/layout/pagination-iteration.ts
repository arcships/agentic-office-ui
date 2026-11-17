/**
 * Measure→repaginate loop (issue #99).
 *
 * The estimator can only approximate rendered paragraph/table heights. This
 * module runs the convergence loop: estimate → (caller renders and) measures →
 * feed measured heights back into the segmenter, bounded by `maxIterations`
 * and guarded against oscillation by comparing page identities and feedback
 * between passes.
 */

export interface PaginationMeasurement<TPage, TFeedback> {
  /** Rendered content height of every page (used for convergence checks). */
  pageContentHeightsPx: number[];
  /** Opaque feedback (paragraph/table/page heights) for the next build pass. */
  feedback: TFeedback;
}

export interface IteratePaginationOptions<TPage, TFeedback = number[]> {
  /**
   * Build a pagination pass. `feedback` is the value from the previous
   * measurement; `undefined` performs the initial estimate.
   */
  buildPages: (feedback?: TFeedback) => TPage[];
  /** Measure the rendered pages and produce the next feedback. */
  measurePages: (pages: TPage[]) => PaginationMeasurement<TPage, TFeedback>;
  /** Stable identity of a page's segments; used to reject stale measurements. */
  pageIdentityKeys?: (pages: TPage[]) => string[];
  /** Optional conservative stabilizer for the measured page content heights. */
  stabilizePageHeights?: (current: number[], next: number[]) => number[];
  maxIterations?: number;
  tolerancePx?: number;
}

export interface IteratedPaginationResult<TPage, TFeedback> {
  pages: TPage[];
  iterations: number;
  converged: boolean;
  oscillationDetected: boolean;
  feedback?: TFeedback;
  measuredPageContentHeightsPx?: number[];
}

function pageHeightsDiffer(
  previous: number[],
  next: number[],
  tolerancePx: number
): boolean {
  if (previous.length !== next.length) {
    return true;
  }
  return previous.some(
    (heightPx, pageIndex) =>
      Math.abs(heightPx - (next[pageIndex] ?? 0)) > tolerancePx
  );
}

function feedbackIsStable(
  previous: unknown,
  next: unknown,
  tolerancePx: number
): boolean {
  if (previous === next) {
    return true;
  }
  if (
    typeof previous === "number" &&
    typeof next === "number"
  ) {
    return Math.abs(previous - next) <= tolerancePx;
  }
  if (Array.isArray(previous) && Array.isArray(next)) {
    return (
      previous.length === next.length &&
      previous.every((value, index) =>
        feedbackIsStable(value, next[index], tolerancePx)
      )
    );
  }
  if (
    previous &&
    next &&
    typeof previous === "object" &&
    typeof next === "object"
  ) {
    const previousKeys = Object.keys(previous).sort();
    const nextKeys = Object.keys(next).sort();
    if (previousKeys.join("\u0000") !== nextKeys.join("\u0000")) {
      return false;
    }
    return previousKeys.every((key) =>
      feedbackIsStable(
        (previous as Record<string, unknown>)[key],
        (next as Record<string, unknown>)[key],
        tolerancePx
      )
    );
  }
  return false;
}

export function iteratePaginationWithMeasurements<TPage, TFeedback = number[]>(
  options: IteratePaginationOptions<TPage, TFeedback>
): IteratedPaginationResult<TPage, TFeedback> {
  const maxIterations = Math.max(1, Math.round(options.maxIterations ?? 3));
  const tolerancePx = Math.max(0, options.tolerancePx ?? 1);
  const seenSignatures = new Set<string>();

  let pages = options.buildPages();
  let feedback: TFeedback | undefined;
  let measuredPageHeights: number[] | undefined;
  let iterations = 0;

  while (iterations < maxIterations) {
    const measurement = options.measurePages(pages);
    const measured = measurement.pageContentHeightsPx;
    if (
      measured.length !== pages.length ||
      measured.some((heightPx) => !Number.isFinite(heightPx) || heightPx <= 0)
    ) {
      return {
        pages,
        iterations,
        converged: false,
        oscillationDetected: false,
        feedback,
        measuredPageContentHeightsPx: measuredPageHeights,
      };
    }

    const stabilizedHeights =
      measuredPageHeights && options.stabilizePageHeights
        ? options.stabilizePageHeights(measuredPageHeights, measured)
        : sanitizeMeasuredPageHeights(measured);
    const identityKeys =
      options.pageIdentityKeys?.(pages) ?? [];
    const feedbackStable =
      feedback !== undefined &&
      feedbackIsStable(feedback, measurement.feedback, tolerancePx);
    const pageHeightsStable =
      measuredPageHeights !== undefined &&
      !pageHeightsDiffer(measuredPageHeights, stabilizedHeights, tolerancePx);
    if (feedbackStable && pageHeightsStable) {
      return {
        pages,
        iterations,
        converged: true,
        oscillationDetected: false,
        feedback,
        measuredPageContentHeightsPx: measuredPageHeights,
      };
    }

    const signature = `${pages.length}|${identityKeys.join("::")}|${stabilizedHeights
      .map((heightPx) => Math.round(heightPx / Math.max(1, tolerancePx)))
      .join(",")}`;
    if (seenSignatures.has(signature)) {
      return {
        pages,
        iterations,
        converged: false,
        oscillationDetected: true,
        feedback,
        measuredPageContentHeightsPx: measuredPageHeights,
      };
    }
    seenSignatures.add(signature);

    const nextPages = options.buildPages(measurement.feedback);
    iterations += 1;

    const nextKeys = options.pageIdentityKeys?.(nextPages) ?? [];
    const identitiesStable =
      nextKeys.length === identityKeys.length &&
      nextKeys.every((key, pageIndex) => key === identityKeys[pageIndex]);
    if (
      identitiesStable &&
      nextPages.length === pages.length &&
      pageHeightsStable
    ) {
      return {
        pages: nextPages,
        iterations,
        converged: true,
        oscillationDetected: false,
        feedback: measurement.feedback,
        measuredPageContentHeightsPx: stabilizedHeights,
      };
    }

    pages = nextPages;
    feedback = measurement.feedback;
    measuredPageHeights = stabilizedHeights;
  }

  return {
    pages,
    iterations,
    converged: false,
    oscillationDetected: false,
    feedback,
    measuredPageContentHeightsPx: measuredPageHeights,
  };
}

function sanitizeMeasuredPageHeights(heights: number[]): number[] {
  return heights.map((heightPx) =>
    Number.isFinite(heightPx) && heightPx > 0 ? heightPx : 1
  );
}
