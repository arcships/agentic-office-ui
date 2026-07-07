import type { PaginationSectionMetrics } from "./pagination";
import type {
  DocModel,
  NumberingDefinitionSet,
  ParagraphNode,
  TableNode
} from "../engine/types";
import {
  paragraphAfterSpacingPx,
  paragraphBeforeSpacingPx,
  resolvePaginationSectionMetricsIndexForNodeIndex,
  resolveParagraphBeforeSpacingPx
} from "./pagination";

export function normalizedPositivePixelValue(value: number | undefined, fallback: number): number {
  if (!Number.isFinite(value) || (value as number) <= 0) {
    return fallback;
  }

  return Math.max(1, Math.round(value as number));
}

export function sumEstimatedTableRowHeightsPx(
  rowHeightsPx: number[],
  startRowIndex: number,
  endRowIndex: number,
  minParagraphLineHeightPx: number
): number {
  let total = 0;
  const clampedStart = Math.max(0, startRowIndex);
  const clampedEnd = Math.max(clampedStart, Math.min(endRowIndex, rowHeightsPx.length));
  for (let rowIndex = clampedStart; rowIndex < clampedEnd; rowIndex += 1) {
    total += Math.max(1, rowHeightsPx[rowIndex] ?? minParagraphLineHeightPx);
  }
  return total;
}

export function fitTableRowsWithinHeightPx(
  rowHeightsPx: number[],
  startRowIndex: number,
  availableHeightPx: number,
  forceAtLeastOneRow: boolean,
  minParagraphLineHeightPx: number,
  pageOverflowTolerancePx: number
): number {
  if (startRowIndex >= rowHeightsPx.length) {
    return startRowIndex;
  }

  const safeAvailableHeightPx =
    Number.isFinite(availableHeightPx) && availableHeightPx > 0 ? availableHeightPx : 0;
  let consumedHeightPx = 0;
  let rowCursor = startRowIndex;

  while (rowCursor < rowHeightsPx.length) {
    const rowHeightPx = Math.max(1, rowHeightsPx[rowCursor] ?? minParagraphLineHeightPx);
    if (consumedHeightPx + rowHeightPx > safeAvailableHeightPx + pageOverflowTolerancePx) {
      break;
    }

    consumedHeightPx += rowHeightPx;
    rowCursor += 1;
  }

  if (rowCursor === startRowIndex && forceAtLeastOneRow) {
    return Math.min(rowHeightsPx.length, startRowIndex + 1);
  }

  return rowCursor;
}

export function normalizeFallbackMetrics(
  pageContentHeightPx: number,
  pageContentWidthPx: number
): PaginationSectionMetrics {
  return {
    startNodeIndex: 0,
    pageContentWidthPx: Math.max(120, Math.round(pageContentWidthPx)),
    pageContentHeightPx: Math.max(120, Math.round(pageContentHeightPx)),
    docGridLinePitchPx: undefined
  };
}

export function normalizedMetricsBySection(
  pageContentHeightPx: number,
  pageContentWidthPx: number,
  paginationMetricsBySection?: PaginationSectionMetrics[]
): {
  fallbackMetrics: PaginationSectionMetrics;
  metricsBySection: PaginationSectionMetrics[];
} {
  const fallbackMetrics = normalizeFallbackMetrics(pageContentHeightPx, pageContentWidthPx);
  return {
    fallbackMetrics,
    metricsBySection: paginationMetricsBySection?.length
      ? paginationMetricsBySection
      : [fallbackMetrics]
  };
}

export function normalizedMeasuredTableRowHeights(
  measuredRowHeightsPx: number[] | undefined,
  expectedRowCount: number,
  minParagraphLineHeightPx: number
): number[] | undefined {
  if (!measuredRowHeightsPx || measuredRowHeightsPx.length !== expectedRowCount) {
    return undefined;
  }

  return measuredRowHeightsPx.map((heightPx) =>
    Math.max(
      minParagraphLineHeightPx,
      Number.isFinite(heightPx)
        ? Math.round(heightPx as number)
        : minParagraphLineHeightPx
    )
  );
}

export interface PageSegmentationTableCallbacks {
  estimateDocNodeHeightPx: (
    node: DocModel["nodes"][number],
    availableWidthPx?: number,
    numberingDefinitions?: NumberingDefinitionSet,
    docGridLinePitchPx?: number
  ) => number;
  paragraphHasVisibleText: (paragraph: ParagraphNode) => boolean;
  estimateParagraphHeightPx: (
    paragraph: ParagraphNode,
    availableWidthPx?: number,
    numberingDefinitions?: NumberingDefinitionSet,
    docGridLinePitchPx?: number
  ) => number;
  estimateParagraphLineHeightPx: (
    paragraph: ParagraphNode,
    docGridLinePitchPx?: number
  ) => number;
  estimateTableRowHeightsPx: (
    table: TableNode,
    maxAvailableWidthPx?: number,
    numberingDefinitions?: NumberingDefinitionSet,
    docGridLinePitchPx?: number
  ) => number[];
}

export interface OverflowBreakCollectionOptions {
  suppressSpacingBeforeAfterPageBreak?: boolean;
  pageOverflowTolerancePx?: number;
}

export function estimateKeepNextTableRequiredHeightPx(
  table: TableNode,
  callbacks: PageSegmentationTableCallbacks,
  metrics: PaginationSectionMetrics,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  minParagraphLineHeightPx: number,
  pageOverflowTolerancePx: number
): number {
  const rowHeightsPx = callbacks.estimateTableRowHeightsPx(
    table,
    metrics.pageContentWidthPx,
    numberingDefinitions,
    metrics.docGridLinePitchPx
  );
  if (rowHeightsPx.length === 0) {
    return Math.max(
      1,
      normalizedPositivePixelValue(
        callbacks.estimateDocNodeHeightPx(
          table,
          metrics.pageContentWidthPx,
          numberingDefinitions,
          metrics.docGridLinePitchPx
        ),
        minParagraphLineHeightPx
      )
    );
  }

  const normalizedRowHeightsPx = rowHeightsPx.map((rowHeightPx) =>
    Math.max(
      minParagraphLineHeightPx,
      normalizedPositivePixelValue(rowHeightPx, minParagraphLineHeightPx)
    )
  );
  const totalTableHeightPx = normalizedRowHeightsPx.reduce(
    (sum, rowHeightPx) => sum + rowHeightPx,
    0
  );
  if (totalTableHeightPx <= metrics.pageContentHeightPx + pageOverflowTolerancePx) {
    return totalTableHeightPx;
  }

  return Math.max(
    minParagraphLineHeightPx,
    normalizedPositivePixelValue(normalizedRowHeightsPx[0], minParagraphLineHeightPx)
  );
}

export function collectDocxEstimatedOverflowBreakStartNodeIndexes(
  model: DocModel,
  hardBreakStartNodeIndexes: Set<number>,
  pageContentHeightPx: number,
  pageContentWidthPx: number,
  callbacks: PageSegmentationTableCallbacks,
  numberingDefinitions?: NumberingDefinitionSet,
  paginationMetricsBySection?: PaginationSectionMetrics[],
  options?: OverflowBreakCollectionOptions
): Set<number> {
  const breaks = new Set<number>();
  if (!Number.isFinite(pageContentHeightPx) || pageContentHeightPx <= 0) {
    return breaks;
  }

  const DEFAULT_OVERFLOW_TOLERANCE_PX = 2;
  const DEFAULT_MIN_LINE_HEIGHT_PX = 14;
  const pageOverflowTolerancePx =
    options?.pageOverflowTolerancePx ?? DEFAULT_OVERFLOW_TOLERANCE_PX;
  const { fallbackMetrics, metricsBySection } = normalizedMetricsBySection(
    pageContentHeightPx,
    pageContentWidthPx,
    paginationMetricsBySection
  );

  let pageConsumedHeightPx = 0;
  let previousParagraphAfterPx = 0;
  let currentMetricsIndex = 0;
  let committedKeepNextChainEndNodeIndex = -1;
  const suppressSpacingBeforeAfterPageBreak =
    options?.suppressSpacingBeforeAfterPageBreak ?? false;
  let currentPageContentHeightPx =
    metricsBySection[0]?.pageContentHeightPx ?? fallbackMetrics.pageContentHeightPx;

  for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
    currentMetricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
      metricsBySection,
      nodeIndex,
      currentMetricsIndex
    );
    const nodeMetrics = metricsBySection[currentMetricsIndex] ?? fallbackMetrics;

    if (hardBreakStartNodeIndexes.has(nodeIndex)) {
      pageConsumedHeightPx = 0;
      previousParagraphAfterPx = 0;
      currentPageContentHeightPx = nodeMetrics.pageContentHeightPx;
    }

    const node = model.nodes[nodeIndex];
    const rawNodeHeightPx = Math.max(
      1,
      normalizedPositivePixelValue(
        callbacks.estimateDocNodeHeightPx(
          node,
          nodeMetrics.pageContentWidthPx,
          numberingDefinitions,
          nodeMetrics.docGridLinePitchPx
        ),
        1
      )
    );
    const nodeBeforeSpacingPx =
      node.type === "paragraph"
        ? resolveParagraphBeforeSpacingPx(
            model,
            nodeIndex,
            node,
            pageConsumedHeightPx,
            suppressSpacingBeforeAfterPageBreak
          )
        : 0;
    const collapsedMarginPx =
      node.type === "paragraph" && pageConsumedHeightPx > 0
        ? Math.min(previousParagraphAfterPx, nodeBeforeSpacingPx)
        : 0;
    const collapsedNodeHeightPx = Math.max(1, rawNodeHeightPx - collapsedMarginPx);

    let requiredHeightPx = collapsedNodeHeightPx;
    let keepNextChainEndNodeIndex = -1;

    if (
      node.type === "paragraph" &&
      node.style?.keepNext === true &&
      nodeIndex > committedKeepNextChainEndNodeIndex &&
      callbacks.paragraphHasVisibleText(node)
    ) {
      let chainCursor = nodeIndex;
      let chainPreviousParagraphAfterPx = paragraphAfterSpacingPx(node);
      while (chainCursor < model.nodes.length - 1) {
        const currentChainNode = model.nodes[chainCursor];
        if (
          currentChainNode.type !== "paragraph" ||
          currentChainNode.style?.keepNext !== true ||
          !callbacks.paragraphHasVisibleText(currentChainNode)
        ) {
          break;
        }
        if (hardBreakStartNodeIndexes.has(chainCursor + 1)) {
          break;
        }
        const nextChainNode = model.nodes[chainCursor + 1];
        const chainMetricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
          metricsBySection,
          chainCursor + 1,
          currentMetricsIndex
        );
        const chainMetrics = metricsBySection[chainMetricsIndex] ?? fallbackMetrics;
        if (nextChainNode.type === "table") {
          requiredHeightPx += estimateKeepNextTableRequiredHeightPx(
            nextChainNode,
            callbacks,
            chainMetrics,
            numberingDefinitions,
            DEFAULT_MIN_LINE_HEIGHT_PX,
            pageOverflowTolerancePx
          );
          break;
        }
        if (nextChainNode.type !== "paragraph") {
          break;
        }

        chainCursor += 1;
        const nextRawHeightPx = Math.max(
          1,
          normalizedPositivePixelValue(
            callbacks.estimateParagraphHeightPx(
              nextChainNode,
              chainMetrics.pageContentWidthPx,
              numberingDefinitions,
              chainMetrics.docGridLinePitchPx
            ),
            1
          )
        );
        const collapsedChainMarginPx = Math.min(
          chainPreviousParagraphAfterPx,
          paragraphBeforeSpacingPx(nextChainNode)
        );
        requiredHeightPx += Math.max(1, nextRawHeightPx - collapsedChainMarginPx);
        chainPreviousParagraphAfterPx = paragraphAfterSpacingPx(nextChainNode);
      }
      keepNextChainEndNodeIndex = chainCursor;
    }

    const remainingHeightPx = currentPageContentHeightPx - pageConsumedHeightPx;
    if (
      pageConsumedHeightPx > 0 &&
      requiredHeightPx > remainingHeightPx + pageOverflowTolerancePx
    ) {
      breaks.add(nodeIndex);
      pageConsumedHeightPx = 0;
      previousParagraphAfterPx = 0;
      currentPageContentHeightPx = nodeMetrics.pageContentHeightPx;
    }

    if (pageConsumedHeightPx === 0 && keepNextChainEndNodeIndex > nodeIndex) {
      committedKeepNextChainEndNodeIndex = keepNextChainEndNodeIndex;
    }
    const effectiveNodeHeightPx =
      pageConsumedHeightPx > 0 ? collapsedNodeHeightPx : rawNodeHeightPx;
    pageConsumedHeightPx += effectiveNodeHeightPx;
    previousParagraphAfterPx = node.type === "paragraph" ? paragraphAfterSpacingPx(node) : 0;
  }

  for (const breakIndex of [...breaks]) {
    if (
      breakIndex <= 0 ||
      breakIndex >= model.nodes.length ||
      hardBreakStartNodeIndexes.has(breakIndex)
    ) {
      breaks.delete(breakIndex);
    }
  }

  return breaks;
}
