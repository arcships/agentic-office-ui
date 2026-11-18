/**
 * Measure→repaginate loop (issue #99).
 *
 * The estimator can only approximate rendered paragraph/table heights. This
 * module runs the convergence loop: estimate → (caller renders and) measures →
 * feed measured heights back into the segmenter, bounded by `maxIterations`
 * and guarded against oscillation by comparing page identities and feedback
 * between passes.
 *
 * `documentPageCount` is accepted as an advisory `targetPageCount`: the loop
 * reports whether it was reached, but it never changes pagination to force it.
 */

export interface PaginationMeasurement<TPage, TFeedback> {
  /** Rendered content height of every page (used for convergence checks). */
  pageContentHeightsPx: number[];
  /** Opaque feedback (paragraph/table/page heights) for the next build pass. */
  feedback: TFeedback;
}

export type PaginationMeasurementAction =
  | "repaginate"
  | "converged"
  | "oscillating"
  | "exhausted";

export interface PaginationMeasurementObservation<TFeedback> {
  /** Stable identity of every page in the current render. */
  pageIdentityKeys: readonly string[];
  /** Rendered content height of every page, aligned with the identity keys. */
  pageContentHeightsPx: readonly number[];
  /** Feedback the next build pass should apply if another pass is needed. */
  feedback: TFeedback;
}

export interface PaginationMeasurementDecision<TFeedback> {
  action: PaginationMeasurementAction;
  /** Repagination passes applied so far (the initial estimate is pass 0). */
  iteration: number;
  /** Page count that produced this decision. */
  pageCount: number;
  /** Advisory target from `documentPageCount`. */
  targetPageCount?: number;
  /** True when the rendered page count equals the advisory target. */
  targetReached: boolean;
  /** Feedback to apply when `action === "repaginate"`; the observed one otherwise. */
  feedback?: TFeedback;
  pageContentHeightsPx: number[];
}

export interface PaginationMeasurementLoopOptions {
  /** Maximum repagination passes after the initial estimate. Defaults to 3. */
  maxIterations?: number;
  /** Convergence tolerance for page heights/feedback. Defaults to 1 px. */
  tolerancePx?: number;
  /** Advisory page count (`documentPageCount`); never enforced. */
  targetPageCount?: number;
  /** Conservative stabilizer for the measured page content heights. */
  stabilizePageHeights?: (current: number[], next: number[]) => number[];
  /**
   * Require stable measured page heights for convergence. The viewer disables
   * this because only virtualized (visible) pages can be measured; pagination
   * stability is then decided by page identities and feedback alone.
   */
  requireStablePageHeights?: boolean;
}

function normalizeTolerance(tolerancePx: number | undefined): number {
  return Number.isFinite(tolerancePx) ? Math.max(0, tolerancePx as number) : 1;
}

function normalizeIterations(maxIterations: number | undefined): number {
  return Number.isFinite(maxIterations)
    ? Math.max(1, Math.round(maxIterations as number))
    : 3;
}

function normalizeTargetPageCount(
  targetPageCount: number | undefined
): number | undefined {
  if (!Number.isFinite(targetPageCount) || (targetPageCount as number) < 1) {
    return undefined;
  }
  return Math.max(1, Math.round(targetPageCount as number));
}

function sanitizePageContentHeights(
  heights: readonly number[] | undefined
): number[] {
  return (heights ?? []).map((heightPx) =>
    Number.isFinite(heightPx) && heightPx > 0 ? heightPx : 1
  );
}

function stringArraysEqual(
  previous: readonly string[] | undefined,
  next: readonly string[]
): boolean {
  return (
    previous !== undefined &&
    previous.length === next.length &&
    previous.every((value, index) => value === next[index])
  );
}

function numberArraysEqual(
  previous: readonly number[] | undefined,
  next: readonly number[],
  tolerancePx: number
): boolean {
  return (
    previous !== undefined &&
    previous.length === next.length &&
    previous.every(
      (value, index) => Math.abs(value - (next[index] ?? 0)) <= tolerancePx
    )
  );
}

/** Structural tolerance-aware comparison of arbitrary feedback values. */
export function isPaginationFeedbackStable(
  previous: unknown,
  next: unknown,
  tolerancePx = 1
): boolean {
  if (previous === next) {
    return true;
  }
  if (typeof previous === "number" && typeof next === "number") {
    return Math.abs(previous - next) <= tolerancePx;
  }
  if (Array.isArray(previous) && Array.isArray(next)) {
    return (
      previous.length === next.length &&
      previous.every((value, index) =>
        isPaginationFeedbackStable(value, next[index], tolerancePx)
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
      isPaginationFeedbackStable(
        (previous as Record<string, unknown>)[key],
        (next as Record<string, unknown>)[key],
        tolerancePx
      )
    );
  }
  return false;
}

/**
 * Event-driven convergence gate for one viewer. The viewer cannot build and
 * measure synchronously (measurements arrive from ResizeObserver/DOM), so it
 * feeds each completed render into `observe` and applies the returned feedback
 * only for `repaginate`.
 */
export class PaginationMeasurementLoop<TFeedback> {
  private readonly maxIterations: number;
  private readonly tolerancePx: number;
  private readonly targetPageCount?: number;
  private readonly stabilizePageHeights?: (
    current: number[],
    next: number[]
  ) => number[];
  private readonly requireStablePageHeights: boolean;

  private previousIdentityKeys?: string[];
  private previousPageContentHeights?: number[];
  private previousFeedback?: TFeedback;
  private readonly seenSignatures = new Set<string>();
  private _iteration = 0;
  private _state: "idle" | PaginationMeasurementAction = "idle";
  private _pageCount = 0;
  private _targetReached = false;
  private _lastDecision?: PaginationMeasurementDecision<TFeedback>;

  constructor(options: PaginationMeasurementLoopOptions = {}) {
    this.maxIterations = normalizeIterations(options.maxIterations);
    this.tolerancePx = normalizeTolerance(options.tolerancePx);
    this.targetPageCount = normalizeTargetPageCount(options.targetPageCount);
    this.stabilizePageHeights = options.stabilizePageHeights;
    this.requireStablePageHeights =
      options.requireStablePageHeights !== false;
  }

  get iteration(): number {
    return this._iteration;
  }

  get state(): "idle" | PaginationMeasurementAction {
    return this._state;
  }

  get pageCount(): number {
    return this._pageCount;
  }

  get targetReached(): boolean {
    return this._targetReached;
  }

  get lastDecision(): PaginationMeasurementDecision<TFeedback> | undefined {
    return this._lastDecision;
  }

  reset(): void {
    this.previousIdentityKeys = undefined;
    this.previousPageContentHeights = undefined;
    this.previousFeedback = undefined;
    this.seenSignatures.clear();
    this._iteration = 0;
    this._state = "idle";
    this._pageCount = 0;
    this._targetReached = false;
    this._lastDecision = undefined;
  }

  observe(
    observation: PaginationMeasurementObservation<TFeedback>
  ): PaginationMeasurementDecision<TFeedback> {
    const identityKeys = [...observation.pageIdentityKeys];
    const pageCount = identityKeys.length;
    const measuredHeights = sanitizePageContentHeights(
      observation.pageContentHeightsPx
    );
    this._pageCount = pageCount;
    this._targetReached =
      this.targetPageCount !== undefined && pageCount === this.targetPageCount;

    if (pageCount === 0) {
      return this.finish({
        action: "converged",
        iteration: this._iteration,
        pageCount,
        targetPageCount: this.targetPageCount,
        targetReached: this._targetReached,
        feedback: observation.feedback,
        pageContentHeightsPx: measuredHeights,
      });
    }

    const signature = [
      pageCount,
      identityKeys.join("\u0000"),
      measuredHeights
        .map((heightPx) => Math.round(heightPx / Math.max(1, this.tolerancePx)))
        .join(","),
    ].join("|");

    const stabilizedHeights =
      this.previousPageContentHeights && this.stabilizePageHeights
        ? sanitizePageContentHeights(
            this.stabilizePageHeights(
              this.previousPageContentHeights,
              measuredHeights
            )
          )
        : measuredHeights;

    const identityKeysStable = stringArraysEqual(
      this.previousIdentityKeys,
      identityKeys
    );
    const pageHeightsStable =
      !this.requireStablePageHeights ||
      numberArraysEqual(
        this.previousPageContentHeights,
        stabilizedHeights,
        this.tolerancePx
      );
    const feedbackStable =
      this.previousFeedback !== undefined &&
      isPaginationFeedbackStable(
        this.previousFeedback,
        observation.feedback,
        this.tolerancePx
      );

    if (identityKeysStable && pageHeightsStable && feedbackStable) {
      return this.finish({
        action: "converged",
        iteration: this._iteration,
        pageCount,
        targetPageCount: this.targetPageCount,
        targetReached: this._targetReached,
        feedback: observation.feedback,
        pageContentHeightsPx: stabilizedHeights,
      });
    }

    if (this.seenSignatures.has(signature)) {
      return this.finish({
        action: "oscillating",
        iteration: this._iteration,
        pageCount,
        targetPageCount: this.targetPageCount,
        targetReached: this._targetReached,
        feedback: this.previousFeedback ?? observation.feedback,
        pageContentHeightsPx: stabilizedHeights,
      });
    }
    this.seenSignatures.add(signature);

    if (this._iteration >= this.maxIterations) {
      return this.finish({
        action: "exhausted",
        iteration: this._iteration,
        pageCount,
        targetPageCount: this.targetPageCount,
        targetReached: this._targetReached,
        feedback: observation.feedback,
        pageContentHeightsPx: stabilizedHeights,
      });
    }

    this.previousIdentityKeys = identityKeys;
    this.previousPageContentHeights = stabilizedHeights;
    this.previousFeedback = observation.feedback;
    this._iteration += 1;

    return this.finish({
      action: "repaginate",
      iteration: this._iteration,
      pageCount,
      targetPageCount: this.targetPageCount,
      targetReached: this._targetReached,
      feedback: observation.feedback,
      pageContentHeightsPx: stabilizedHeights,
    });
  }

  private finish(
    decision: PaginationMeasurementDecision<TFeedback>
  ): PaginationMeasurementDecision<TFeedback> {
    this._state = decision.action;
    this._lastDecision = decision;
    return decision;
  }
}

export interface IteratePaginationOptions<TPages, TFeedback = number[]> {
  /**
   * Build a pagination pass. `feedback` is the value from the previous
   * measurement; `undefined` performs the initial estimate.
   */
  buildPages: (feedback?: TFeedback) => TPages;
  /** Measure the rendered pages; `TPages` may be an array or a snapshot object. */
  measurePages: (
    pages: TPages
  ) => PaginationMeasurement<TPages, TFeedback>;
  /** Stable identity of every rendered page; drives convergence/oscillation. */
  pageIdentityKeys: (pages: TPages) => string[];
  /** Conservative stabilizer for the measured page content heights. */
  stabilizePageHeights?: (current: number[], next: number[]) => number[];
  maxIterations?: number;
  tolerancePx?: number;
  /** Advisory `documentPageCount`; reported, never enforced. */
  targetPageCount?: number;
  /** Defaults to true; set false when only a virtualized page window measures. */
  requireStablePageHeights?: boolean;
}

export interface IteratedPaginationResult<TPages, TFeedback> {
  pages: TPages;
  /** Repagination passes applied (the initial estimate is pass 0). */
  iterations: number;
  converged: boolean;
  oscillationDetected: boolean;
  feedback?: TFeedback;
  measuredPageContentHeightsPx?: number[];
  pageCount: number;
  targetPageCount?: number;
  targetReached: boolean;
}

export function iteratePaginationWithMeasurements<
  TPages,
  TFeedback = number[]
>(
  options: IteratePaginationOptions<TPages, TFeedback>
): IteratedPaginationResult<TPages, TFeedback> {
  const loop = new PaginationMeasurementLoop<TFeedback>({
    maxIterations: options.maxIterations,
    tolerancePx: options.tolerancePx,
    targetPageCount: options.targetPageCount,
    stabilizePageHeights: options.stabilizePageHeights,
    requireStablePageHeights: options.requireStablePageHeights,
  });

  let pages = options.buildPages();
  let feedback: TFeedback | undefined;
  let measuredPageContentHeightsPx: number[] | undefined;

  for (;;) {
    const measurement = options.measurePages(pages);
    const decision = loop.observe({
      pageIdentityKeys: options.pageIdentityKeys(pages),
      pageContentHeightsPx: measurement.pageContentHeightsPx,
      feedback: measurement.feedback,
    });
    measuredPageContentHeightsPx = decision.pageContentHeightsPx;

    if (decision.action === "repaginate") {
      feedback = decision.feedback;
      pages = options.buildPages(feedback);
      continue;
    }

    return {
      pages,
      iterations: decision.iteration,
      converged: decision.action === "converged",
      oscillationDetected: decision.action === "oscillating",
      feedback,
      measuredPageContentHeightsPx,
      pageCount: decision.pageCount,
      targetPageCount: decision.targetPageCount,
      targetReached: decision.targetReached,
    };
  }
}
