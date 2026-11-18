// Per-line vertical metrics for a paragraph, and the text-flow height they
// imply.
//
// DOCX line boxes are not `lineCount * lineHeight`: a line grows around an
// inline image or a run in a larger font, and the page break may trim the
// suppressible bottom space of the last line it keeps. Pagination therefore asks
// this module for the paragraph's line metrics before it splits the paragraph
// across pages.
//
// Upstream: `packages/layout-engine/src/line-metrics.ts` (aggregateLineMetrics /
// checkLineVerticalFit / constrainParagraphLineSplit) plus the per-fragment
// ascent/descent carried by `pretext-layout.ts:102-135`.

import type {
  NumberingDefinitionSet,
  ParagraphNode
} from "../../engine/types";
import type { PretextExclusionRect } from "../../viewer/pretext-layout";
import {
  emptyLineMetrics,
  sumLineHeightsPx,
  type LineVerticalMetrics
} from "../../viewer/line-metrics";
import {
  estimateParagraphLineHeightPx,
  resolveParagraphDocGridLinePitchPx
} from "./line-height";
import { paragraphAvailableTextWidthPx } from "./paragraph-geometry";
import {
  buildParagraphPretextLayoutItems,
  buildParagraphPretextLayoutSource,
  layoutParagraphPretextSource
} from "./pretext-build";

// Pagination asks for the same paragraph's line metrics on every pass; cache
// them per paragraph object (an edit produces a new paragraph object).
const paragraphLineMetricsByParagraph = new WeakMap<
  ParagraphNode,
  Map<string, LineVerticalMetrics[] | null>
>();

export interface ParagraphLineMetricsOptions {
  /** Width of the paragraph's content box (page content width). */
  availableWidthPx?: number;
  numberingDefinitions?: NumberingDefinitionSet;
  /** Paragraph line height; resolved from the paragraph when omitted. */
  lineHeightPx?: number;
  docGridLinePitchPx?: number;
  disableDocGridSnap?: boolean;
  exclusions?: PretextExclusionRect[];
  /**
   * Line count the caller already computed. Metrics are only returned when the
   * layout agrees with it, so pagination never mixes two line counts.
   */
  expectedLineCount?: number;
}

/**
 * True when the paragraph contains content whose vertical metrics a font-only
 * line box cannot describe: an in-flow image, or a run in a different font
 * size than the paragraph's dominant size.
 */
export function paragraphHasCustomLineMetrics(
  paragraph: ParagraphNode,
  source = buildParagraphPretextLayoutSource(paragraph, {
    allowExplicitLineBreakText: true,
    expandTabsForLayout: true
  })
): boolean {
  if (!source) {
    return false;
  }

  return buildParagraphPretextLayoutItems(paragraph, source).some(
    (item) => item.hasCustomVerticalMetrics === true
  );
}

/**
 * Line-box metrics of `paragraph` at `availableWidthPx`, or `undefined` when
 * the paragraph has no custom vertical metrics (plain text keeps the existing
 * `lineCount * lineHeight` model) or when the layout cannot be trusted.
 */
export function resolveParagraphLineMetrics(
  paragraph: ParagraphNode,
  options: ParagraphLineMetricsOptions = {}
): LineVerticalMetrics[] | undefined {
  const availableWidthPx = options.availableWidthPx;
  if (
    !Number.isFinite(availableWidthPx) ||
    (availableWidthPx as number) <= 0
  ) {
    return undefined;
  }

  const source = buildParagraphPretextLayoutSource(paragraph, {
    allowExplicitLineBreakText: true,
    expandTabsForLayout: true
  });
  if (!source || !paragraphHasCustomLineMetrics(paragraph, source)) {
    return undefined;
  }

  const contentWidthPx = paragraphAvailableTextWidthPx(
    paragraph,
    availableWidthPx as number,
    options.numberingDefinitions
  );
  if (!Number.isFinite(contentWidthPx) || contentWidthPx <= 0) {
    return undefined;
  }

  const lineHeightPx =
    Number.isFinite(options.lineHeightPx) && (options.lineHeightPx as number) > 0
      ? Math.round(options.lineHeightPx as number)
      : estimateParagraphLineHeightPx(
          paragraph,
          options.docGridLinePitchPx,
          options.disableDocGridSnap ?? false
        );
  const docGridLinePitchPx = resolveParagraphDocGridLinePitchPx(
    paragraph,
    options.docGridLinePitchPx,
    options.disableDocGridSnap ?? false
  );
  const cacheKey = [
    Math.round(contentWidthPx),
    lineHeightPx,
    docGridLinePitchPx ?? 0,
    Number.isFinite(options.expectedLineCount)
      ? Math.round(options.expectedLineCount as number)
      : -1,
    (options.exclusions ?? [])
      .map(
        (exclusion) =>
          `${exclusion.left},${exclusion.right},${exclusion.top},${exclusion.bottom}`
      )
      .join(";")
  ].join("\u0000");
  const cachedByKey = paragraphLineMetricsByParagraph.get(paragraph);
  const cached = cachedByKey?.get(cacheKey);
  if (cached !== undefined) {
    return cached ?? undefined;
  }
  const store = (
    metrics: LineVerticalMetrics[] | undefined
  ): LineVerticalMetrics[] | undefined => {
    const cache =
      paragraphLineMetricsByParagraph.get(paragraph) ??
      new Map<string, LineVerticalMetrics[] | null>();
    cache.set(cacheKey, metrics ?? null);
    paragraphLineMetricsByParagraph.set(paragraph, cache);
    return metrics;
  };

  const layout = layoutParagraphPretextSource(
    paragraph,
    source,
    contentWidthPx,
    lineHeightPx,
    options.exclusions ?? [],
    { docGridLinePitchPx }
  );
  if (!layout || layout.lines.length === 0) {
    return store(undefined);
  }

  if (
    Number.isFinite(options.expectedLineCount) &&
    (options.expectedLineCount as number) > 0 &&
    layout.lines.length !== Math.round(options.expectedLineCount as number)
  ) {
    // The wrap estimate and the real layout disagree; keeping the estimate is
    // safer than mixing two line counts inside one page.
    return store(undefined);
  }

  return store(
    layout.lines.map((line) => line.metrics ?? emptyLineMetrics(lineHeightPx))
  );
}

/**
 * Height of the paragraph's text flow. Uses aggregated line metrics when the
 * paragraph carries custom vertical metrics and falls back to
 * `lineCount * lineHeightPx` otherwise.
 */
export function paragraphTextFlowHeightPx(
  paragraph: ParagraphNode,
  lineCount: number,
  lineHeightPx: number,
  options: ParagraphLineMetricsOptions = {}
): number {
  const safeLineCount = Math.max(1, Math.round(lineCount));
  const safeLineHeightPx = Math.max(
    1,
    Math.round(Number.isFinite(lineHeightPx) ? lineHeightPx : 1)
  );
  const lineMetrics = resolveParagraphLineMetrics(paragraph, {
    ...options,
    lineHeightPx: safeLineHeightPx,
    expectedLineCount: options.expectedLineCount ?? safeLineCount
  });
  if (!lineMetrics) {
    return safeLineHeightPx * safeLineCount;
  }

  return Math.max(
    1,
    Math.round(sumLineHeightsPx(lineMetrics, 0, lineMetrics.length))
  );
}
