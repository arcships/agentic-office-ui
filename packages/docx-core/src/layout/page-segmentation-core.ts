import type {
  DocModel,
  NumberingDefinitionSet,
  ParagraphNode,
  TableNode
} from "../engine/types";
import {
  collectDocxHardPageBreakStartNodeIndexes,
  collectParagraphExplicitPageBreakOffsets,
  collectTableExplicitPageBreakInfo,
  paragraphAfterSpacingPx,
  paragraphBeforeSpacingPx,
  paragraphLastRenderedPageBreakPosition,
  paragraphLastRenderedPageBreakTextOffset,
  paragraphStartsWithLastRenderedPageBreak,
  paragraphHasPageBreakBefore,
  resolveDocumentSectionsFromMetadata,
  resolvePaginationSectionMetricsIndexForNodeIndex,
  resolveParagraphBeforeSpacingPx,
  resolveParagraphConsumedHeightPx,
  sectionBreakRequiredPageParity,
  type PaginationSectionMetrics,
  type SectionStartPageParity
} from "./pagination";
import {
  collectDocxEstimatedOverflowBreakStartNodeIndexes,
  estimateKeepNextTableRequiredHeightPx,
  fitTableRowsWithinHeightPx,
  normalizeFallbackMetrics,
  normalizedMeasuredTableRowHeights,
  normalizedMetricsBySection,
  normalizedPositivePixelValue,
  sumEstimatedTableRowHeightsPx,
  type PageSegmentationTableCallbacks,
  type OverflowBreakCollectionOptions
} from "./page-segmentation-table";
import {
  constrainParagraphLineSplit,
  sumLineHeightsPx,
  type LineVerticalMetrics
} from "../viewer/line-metrics";

export const DEFAULT_PAGE_OVERFLOW_TOLERANCE_PX = 2;
export const DEFAULT_MIN_PARAGRAPH_LINE_HEIGHT_PX = 14;
const PARAGRAPH_SEGMENT_TOP_BLEED_PX = 22;
const PARAGRAPH_SEGMENT_DESCENDER_BLEED_PX = 6;
const PARAGRAPH_SEGMENT_VISUAL_SAFETY_PX = 24;
const LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_RATIO = 0.18;
const LAST_RENDERED_PAGE_BREAK_HINT_MIN_REMAINING_SPACE_PX =
  DEFAULT_MIN_PARAGRAPH_LINE_HEIGHT_PX * 3;
const LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_PX = 120;

export interface TableRowRange {
  startRowIndex: number;
  endRowIndex: number;
  /** Leading `w:tblHeader` rows repeated at the top of this continuation page. */
  repeatHeaderRowCount?: number;
}

export interface TableRowSlice {
  rowIndex: number;
  startOffsetPx: number;
  sliceHeightPx: number;
  totalRowHeightPx: number;
}

export interface ParagraphLineRange {
  startLineIndex: number;
  endLineIndex: number;
  totalLineCount: number;
  lineHeightPx: number;
}

export interface DocumentPageNodeSegment {
  nodeIndex: number;
  tableRowRange?: TableRowRange;
  tableRowSlice?: TableRowSlice;
  paragraphLineRange?: ParagraphLineRange;
}

export interface LetterheadColumnSegmentGroup {
  startOffset: number;
  endOffset: number;
  leftSegments: DocumentPageNodeSegment[];
  rightSegments: DocumentPageNodeSegment[];
}

export interface ParagraphSplitControlOptions {
  allowKeepLinesOverflow?: boolean;
  allowKeepNextOverflow?: boolean;
}

export interface PageSegmentationCallbacks extends PageSegmentationTableCallbacks {
  paragraphIsStructuralSectionBreakSpacer: (paragraph: ParagraphNode) => boolean;
  paragraphLineCountWithinWidth: (
    paragraph: ParagraphNode,
    availableWidthPx?: number,
    numberingDefinitions?: NumberingDefinitionSet
  ) => number;
  paragraphWidowControlEnabled: (paragraph: ParagraphNode) => boolean;
  paragraphCanSplitAcrossPages: (
    paragraph: ParagraphNode,
    lineCount: number,
    options?: ParagraphSplitControlOptions
  ) => boolean;
  /**
   * Per-line vertical metrics of a paragraph, or `undefined` when the paragraph
   * has no custom vertical metrics and the uniform line height applies.
   * Supplied by the editor helper layer to keep this module cycle-free.
   */
  resolveParagraphLineMetrics?: (params: {
    paragraph: ParagraphNode;
    availableWidthPx?: number;
    numberingDefinitions?: NumberingDefinitionSet;
    lineHeightPx: number;
    docGridLinePitchPx?: number;
    expectedLineCount: number;
  }) => LineVerticalMetrics[] | undefined;
}

export interface DocumentPageSegmentationOptions extends OverflowBreakCollectionOptions {
  allowParagraphLineSplitting?: boolean;
  measuredTableRowHeightsByNodeIndex?: Record<number, number[]>;
  measuredPageContentHeightsPxByPageIndex?: number[];
  minParagraphLineHeightPx?: number;
  preferLastRenderedParagraphStartBreaks?: boolean;
  /**
   * Resolves the wrap line on which Word's cached
   * `w:lastRenderedPageBreak` sits. Used to break inside a paragraph instead
   * of moving the whole paragraph to the next page. When omitted, the break
   * line is estimated proportionally from the marker's text offset.
   */
  paragraphLastRenderedPageBreakLineIndex?: (
    paragraph: ParagraphNode,
    lineCount: number,
    availableWidthPx?: number,
    numberingDefinitions?: NumberingDefinitionSet
  ) => number | undefined;
}

function paragraphSegmentHasPartialLineRange(paragraphLineRange?: ParagraphLineRange): boolean {
  if (!paragraphLineRange) {
    return false;
  }

  return (
    paragraphLineRange.startLineIndex > 0 ||
    paragraphLineRange.endLineIndex < paragraphLineRange.totalLineCount
  );
}

function resolveParagraphSegmentNonFlowReservePx(
  paragraphLineRange?: ParagraphLineRange
): number {
  if (!paragraphSegmentHasPartialLineRange(paragraphLineRange)) {
    return 0;
  }

  const topPx =
    paragraphLineRange && paragraphLineRange.startLineIndex > 0
      ? Math.max(0, PARAGRAPH_SEGMENT_TOP_BLEED_PX)
      : 0;
  const bottomPx = Math.max(0, PARAGRAPH_SEGMENT_DESCENDER_BLEED_PX);
  const lineHeightSafetyPx = Math.max(
    0,
    Math.ceil((paragraphLineRange?.lineHeightPx ?? 0) * 0.9)
  );
  return (
    topPx +
    bottomPx +
    Math.max(0, PARAGRAPH_SEGMENT_VISUAL_SAFETY_PX, lineHeightSafetyPx)
  );
}

function shouldHonorParagraphStartLastRenderedPageBreak(params: {
  pageConsumedHeightPx: number;
  pageContentHeightPx: number;
}): boolean {
  const pageConsumedHeightPx = Math.max(
    0,
    Math.round(params.pageConsumedHeightPx)
  );
  const pageContentHeightPx = Math.max(
    0,
    Math.round(params.pageContentHeightPx)
  );
  if (pageConsumedHeightPx <= 0 || pageContentHeightPx <= 0) {
    return false;
  }

  const remainingHeightPx = Math.max(
    0,
    pageContentHeightPx - pageConsumedHeightPx
  );
  const maxAllowedRemainingHeightPx = Math.max(
    LAST_RENDERED_PAGE_BREAK_HINT_MIN_REMAINING_SPACE_PX,
    Math.min(
      LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_PX,
      Math.round(
        pageContentHeightPx *
          LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_RATIO
      )
    )
  );
  return remainingHeightPx <= maxAllowedRemainingHeightPx;
}

function paragraphTextContentLength(paragraph: ParagraphNode): number {
  return paragraph.children.reduce((length, child) => {
    if (child.type === "text") {
      return length + child.text.length;
    }
    if (child.type === "form-field") {
      return length + (child.value ?? "").length;
    }
    return length;
  }, 0);
}

function estimateBreakLineIndexForTextOffset(
  textOffset: number,
  textLength: number,
  lineCount: number
): number | undefined {
  if (textLength <= 0 || lineCount <= 0) {
    return undefined;
  }

  const clampedOffset = Math.max(0, Math.min(textOffset, textLength));
  return Math.max(
    0,
    Math.min(
      lineCount,
      Math.round((clampedOffset / textLength) * lineCount)
    )
  );
}

/**
 * Wrap lines at which the paragraph must start a new page. Explicit
 * `<w:br w:type="page"/>` breaks inside the paragraph are authoritative;
 * Word's cached `w:lastRenderedPageBreak` marker is only honoured when the
 * caller opted into render hints. A mid-paragraph marker is a line-level
 * break, never a node-level page start.
 */
function resolveParagraphForcedPageBreakLineIndexes(params: {
  paragraph: ParagraphNode;
  lineCount: number;
  availableWidthPx?: number;
  numberingDefinitions?: NumberingDefinitionSet;
  preferLastRenderedPageBreakHints: boolean;
  resolveLastRenderedPageBreakLineIndex?: DocumentPageSegmentationOptions["paragraphLastRenderedPageBreakLineIndex"];
}): number[] {
  const {
    paragraph,
    lineCount,
    availableWidthPx,
    numberingDefinitions,
    preferLastRenderedPageBreakHints,
    resolveLastRenderedPageBreakLineIndex
  } = params;
  const textLength = paragraphTextContentLength(paragraph);
  if (textLength <= 0) {
    return [];
  }

  // A paragraph estimated as a single line can still carry an inline break;
  // model it as two flow lines so the break produces a real page boundary
  // instead of being dropped.
  const lineCountForBreak = Math.max(lineCount, 2);
  const breakLineIndexes = new Set<number>();
  const addBreakLineIndex = (lineIndex: number | undefined): void => {
    if (lineIndex === undefined) {
      return;
    }
    const clamped = Math.max(1, Math.min(lineCountForBreak, lineIndex));
    if (clamped < lineCountForBreak) {
      breakLineIndexes.add(clamped);
    }
  };

  collectParagraphExplicitPageBreakOffsets(paragraph).forEach((textOffset) => {
    if (textOffset <= 0 || textOffset >= textLength) {
      return;
    }
    addBreakLineIndex(
      estimateBreakLineIndexForTextOffset(
        textOffset,
        textLength,
        lineCountForBreak
      )
    );
  });

  if (
    preferLastRenderedPageBreakHints &&
    paragraphLastRenderedPageBreakPosition(paragraph) === "internal"
  ) {
    const hintedLineIndex =
      resolveLastRenderedPageBreakLineIndex?.(
        paragraph,
        lineCountForBreak,
        availableWidthPx,
        numberingDefinitions
      ) ??
      (() => {
        const textOffset = paragraphLastRenderedPageBreakTextOffset(paragraph);
        return textOffset === undefined
          ? undefined
          : estimateBreakLineIndexForTextOffset(
              textOffset,
              textLength,
              lineCountForBreak
            );
      })();
    addBreakLineIndex(hintedLineIndex);
  }

  return [...breakLineIndexes].sort((left, right) => left - right);
}

function tableHeaderRowCount(table: TableNode): number {
  let headerRowCount = 0;
  for (const row of table.rows) {
    if (row.style?.isHeader !== true) {
      break;
    }
    headerRowCount += 1;
  }
  return headerRowCount;
}

function tableRowCanSplitAcrossPages(row: TableNode["rows"][number]): boolean {
  return row.style?.cantSplit !== true && row.style?.heightRule !== "exact";
}

function keepNextParagraphReservePx(
  paragraph: ParagraphNode,
  nextParagraph: ParagraphNode | undefined,
  callbacks: PageSegmentationCallbacks,
  docGridLinePitchPx?: number
): number {
  if (
    paragraph.style?.keepNext !== true ||
    !nextParagraph ||
    !Number.isFinite(paragraph.style?.headingLevel)
  ) {
    return 0;
  }

  const nextParagraphText = nextParagraph.children
    .filter((child): child is Extract<typeof child, { type: "text" }> => child.type === "text")
    .map((child) => child.text)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
  if (
    nextParagraph.style?.numbering === undefined &&
    nextParagraphText.length < 80
  ) {
    return 0;
  }

  return Math.max(
    nextParagraph.style?.numbering ? 10 : 6,
    Math.round(
      callbacks.estimateParagraphLineHeightPx(nextParagraph, docGridLinePitchPx) *
        (nextParagraph.style?.numbering ? 1 : 0.5)
    )
  );
}

export function paragraphLetterheadColumnGroupAtSegmentOffset(
  nodeSegments: DocumentPageNodeSegment[],
  startOffset: number,
  resolveFloatSideAtNodeIndex: (nodeIndex: number) => "left" | "right" | undefined
): LetterheadColumnSegmentGroup | undefined {
  if (startOffset < 0 || startOffset >= nodeSegments.length) {
    return undefined;
  }

  const resolveSideForSegment = (
    segment: DocumentPageNodeSegment | undefined
  ): "left" | "right" | undefined => {
    if (
      !segment ||
      segment.tableRowRange ||
      paragraphSegmentHasPartialLineRange(segment.paragraphLineRange)
    ) {
      return undefined;
    }

    return resolveFloatSideAtNodeIndex(segment.nodeIndex);
  };

  const startSide = resolveSideForSegment(nodeSegments[startOffset]);
  const previousSide =
    startOffset > 0 ? resolveSideForSegment(nodeSegments[startOffset - 1]) : undefined;
  if (!startSide || previousSide) {
    return undefined;
  }

  const leftSegments: DocumentPageNodeSegment[] = [];
  const rightSegments: DocumentPageNodeSegment[] = [];
  let endOffset = startOffset;

  while (endOffset < nodeSegments.length) {
    const segment = nodeSegments[endOffset];
    const side = resolveSideForSegment(segment);
    if (!side) {
      break;
    }

    if (side === "left") {
      leftSegments.push(segment);
    } else {
      rightSegments.push(segment);
    }
    endOffset += 1;
  }

  if (leftSegments.length === 0 || rightSegments.length === 0) {
    return undefined;
  }

  return {
    startOffset,
    endOffset,
    leftSegments,
    rightSegments
  };
}

export function buildDocumentPageNodeSegments(
  model: DocModel,
  pageContentHeightPx: number,
  pageContentWidthPx: number,
  callbacks: PageSegmentationCallbacks,
  numberingDefinitions?: NumberingDefinitionSet,
  paginationMetricsBySection?: PaginationSectionMetrics[],
  options?: DocumentPageSegmentationOptions
): DocumentPageNodeSegment[][] {
  if (model.nodes.length === 0) {
    return [];
  }

  const pageOverflowTolerancePx =
    options?.pageOverflowTolerancePx ?? DEFAULT_PAGE_OVERFLOW_TOLERANCE_PX;
  const minParagraphLineHeightPx =
    options?.minParagraphLineHeightPx ?? DEFAULT_MIN_PARAGRAPH_LINE_HEIGHT_PX;
  const { fallbackMetrics, metricsBySection } = normalizedMetricsBySection(
    pageContentHeightPx,
    pageContentWidthPx,
    paginationMetricsBySection
  );

  const pages: DocumentPageNodeSegment[][] = [];
  let currentPageSegments: DocumentPageNodeSegment[] = [];
  const hardBreakStartNodeIndexes = collectDocxHardPageBreakStartNodeIndexes(model);
  const sectionStartPageParityByNodeIndex = new Map<
    number,
    SectionStartPageParity
  >();
  for (const section of resolveDocumentSectionsFromMetadata(model.metadata)) {
    const startNodeIndex = Math.max(0, Math.round(section.startNodeIndex));
    const parity = sectionBreakRequiredPageParity(section.sectionPropertiesXml);
    if (startNodeIndex > 0 && parity) {
      sectionStartPageParityByNodeIndex.set(startNodeIndex, parity);
    }
  }
  const estimatedRowHeightsByTableNodeIndex = new Map<number, number[]>();
  const allowParagraphLineSplitting = options?.allowParagraphLineSplitting ?? true;
  const suppressSpacingBeforeAfterPageBreak =
    options?.suppressSpacingBeforeAfterPageBreak ?? false;
  const preferLastRenderedParagraphStartBreaks =
    options?.preferLastRenderedParagraphStartBreaks ?? false;
  const measuredPageContentHeightsPxByPageIndex =
    options?.measuredPageContentHeightsPxByPageIndex;
  const resolvePageContentHeightPx = (
    pageIndex: number,
    fallbackHeightPx: number
  ): number => {
    const overrideHeightPx = measuredPageContentHeightsPxByPageIndex?.[pageIndex];
    if (Number.isFinite(overrideHeightPx) && (overrideHeightPx as number) > 0) {
      return Math.max(120, Math.round(overrideHeightPx as number));
    }
    return Math.max(120, Math.round(fallbackHeightPx));
  };

  let currentPageIndex = 0;
  const startNextPage = (): void => {
    if (currentPageSegments.length > 0) {
      pages.push(currentPageSegments);
    }
    currentPageSegments = [];
    currentPageIndex = pages.length;
  };
  const startNextPageForSection = (parity?: SectionStartPageParity): void => {
    startNextPage();
    if (parity) {
      // Word numbers pages from 1. `evenPage` sections must start on an even
      // page number and `oddPage` sections on an odd one; otherwise insert a
      // blank page so headers/footers bind to the right parity.
      const nextPageNumber = pages.length + 1;
      const startsOnEvenPage = nextPageNumber % 2 === 0;
      if (startsOnEvenPage !== (parity === "even")) {
        pages.push([]);
      }
    }
    currentPageIndex = pages.length;
  };

  if (!Number.isFinite(pageContentHeightPx) || pageContentHeightPx <= 0) {
    return [model.nodes.map((_, nodeIndex) => ({ nodeIndex }))];
  }

  let pageConsumedHeightPx = 0;
  let currentMetricsIndex = 0;
  let committedKeepNextChainEndNodeIndex = -1;
  let currentPageContentHeightPx = resolvePageContentHeightPx(
    0,
    metricsBySection[0]?.pageContentHeightPx ?? fallbackMetrics.pageContentHeightPx
  );

  for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
    currentMetricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
      metricsBySection,
      nodeIndex,
      currentMetricsIndex
    );
    const nodeMetrics = metricsBySection[currentMetricsIndex] ?? fallbackMetrics;

    if (hardBreakStartNodeIndexes.has(nodeIndex) && currentPageSegments.length > 0) {
      startNextPageForSection(sectionStartPageParityByNodeIndex.get(nodeIndex));
      pageConsumedHeightPx = 0;
      currentPageContentHeightPx = resolvePageContentHeightPx(
        currentPageIndex,
        nodeMetrics.pageContentHeightPx
      );
    } else if (
      hardBreakStartNodeIndexes.has(nodeIndex) &&
      sectionStartPageParityByNodeIndex.has(nodeIndex)
    ) {
      // The previous node ended exactly at the page boundary; the section
      // still needs its parity page.
      startNextPageForSection(sectionStartPageParityByNodeIndex.get(nodeIndex));
      pageConsumedHeightPx = 0;
      currentPageContentHeightPx = resolvePageContentHeightPx(
        currentPageIndex,
        nodeMetrics.pageContentHeightPx
      );
    }

    const node = model.nodes[nodeIndex];
    if (node.type === "paragraph") {
      if (callbacks.paragraphIsStructuralSectionBreakSpacer(node)) {
        continue;
      }

      if (paragraphHasPageBreakBefore(node) && currentPageSegments.length > 0) {
        startNextPage();
        pageConsumedHeightPx = 0;
        currentPageContentHeightPx = resolvePageContentHeightPx(
          currentPageIndex,
          nodeMetrics.pageContentHeightPx
        );
      }

      if (
        preferLastRenderedParagraphStartBreaks &&
        paragraphStartsWithLastRenderedPageBreak(node) &&
        shouldHonorParagraphStartLastRenderedPageBreak({
          pageConsumedHeightPx,
          pageContentHeightPx: currentPageContentHeightPx
        }) &&
        currentPageSegments.length > 0
      ) {
        startNextPage();
        pageConsumedHeightPx = 0;
        currentPageContentHeightPx = resolvePageContentHeightPx(
          currentPageIndex,
          nodeMetrics.pageContentHeightPx
        );
      }

      const rawNodeHeightPx = Math.max(
        1,
        normalizedPositivePixelValue(
          callbacks.estimateParagraphHeightPx(
            node,
            nodeMetrics.pageContentWidthPx,
            numberingDefinitions,
            nodeMetrics.docGridLinePitchPx
          ),
          1
        )
      );
      const paragraphTooTallForSinglePage =
        rawNodeHeightPx > nodeMetrics.pageContentHeightPx + pageOverflowTolerancePx;
      const keepLinesOverflowSplit =
        node.style?.keepLines === true && paragraphTooTallForSinglePage;
      const keepNextOverflowSplit =
        node.style?.keepNext === true && paragraphTooTallForSinglePage;
      const forceOverflowSplit = keepLinesOverflowSplit || keepNextOverflowSplit;
      const nodeIsWithinCommittedKeepNextChain =
        nodeIndex <= committedKeepNextChainEndNodeIndex;
      if (
        forceOverflowSplit &&
        !nodeIsWithinCommittedKeepNextChain &&
        pageConsumedHeightPx > 0 &&
        currentPageSegments.length > 0
      ) {
        startNextPage();
        pageConsumedHeightPx = 0;
        currentPageContentHeightPx = resolvePageContentHeightPx(
          currentPageIndex,
          nodeMetrics.pageContentHeightPx
        );
      }

      const beforeSpacingPx = resolveParagraphBeforeSpacingPx(
        model,
        nodeIndex,
        node,
        pageConsumedHeightPx,
        suppressSpacingBeforeAfterPageBreak
      );
      const afterSpacingPx = paragraphAfterSpacingPx(node);
      const atPageTop = pageConsumedHeightPx <= 0;
      const consumedNodeHeightPx = resolveParagraphConsumedHeightPx(
        rawNodeHeightPx,
        paragraphBeforeSpacingPx(node),
        atPageTop
      );
      const paragraphLineHeightPx = Math.max(
        1,
        normalizedPositivePixelValue(
          callbacks.estimateParagraphLineHeightPx(node, nodeMetrics.docGridLinePitchPx),
          1
        )
      );
      let paragraphLineCount = Math.max(
        1,
        normalizedPositivePixelValue(
          callbacks.paragraphLineCountWithinWidth(
            node,
            nodeMetrics.pageContentWidthPx,
            numberingDefinitions
          ),
          1
        )
      );
      const widowControlEnabled = callbacks.paragraphWidowControlEnabled(node);
      const minLinesPerSegment = widowControlEnabled ? 2 : 1;
      const canSplitParagraphAcrossPages =
        callbacks.paragraphCanSplitAcrossPages(node, paragraphLineCount, {
          allowKeepLinesOverflow: keepLinesOverflowSplit,
          allowKeepNextOverflow: keepNextOverflowSplit
        }) &&
        (!widowControlEnabled || paragraphLineCount > 3);
      const forcedPageBreakLineIndexes = resolveParagraphForcedPageBreakLineIndexes({
        paragraph: node,
        lineCount: paragraphLineCount,
        availableWidthPx: nodeMetrics.pageContentWidthPx,
        numberingDefinitions,
        preferLastRenderedPageBreakHints: preferLastRenderedParagraphStartBreaks,
        resolveLastRenderedPageBreakLineIndex:
          options?.paragraphLastRenderedPageBreakLineIndex
      });
      const mustSplitForForcedBreak = forcedPageBreakLineIndexes.length > 0;
      if (mustSplitForForcedBreak) {
        // A single-line estimate still needs two flow lines so an inline break
        // can split the paragraph instead of being ignored.
        paragraphLineCount = Math.max(
          paragraphLineCount,
          Math.max(...forcedPageBreakLineIndexes) + 1
        );
      }

      if (
        mustSplitForForcedBreak ||
        (canSplitParagraphAcrossPages && allowParagraphLineSplitting)
      ) {
        // Inline images and runs in another font size give the paragraph
        // per-line heights that differ from the uniform estimate; the split
        // below then uses those heights and lets the last line on a page give
        // back its suppressible bottom space.
        const paragraphLineMetrics = callbacks.resolveParagraphLineMetrics?.({
          paragraph: node,
          availableWidthPx: nodeMetrics.pageContentWidthPx,
          numberingDefinitions,
          lineHeightPx: paragraphLineHeightPx,
          docGridLinePitchPx: nodeMetrics.docGridLinePitchPx,
          expectedLineCount: paragraphLineCount
        });
        const segmentLinesHeightPx = (
          startLineIndex: number,
          endLineIndex: number
        ): number =>
          paragraphLineMetrics
            ? sumLineHeightsPx(paragraphLineMetrics, startLineIndex, endLineIndex)
            : Math.max(0, endLineIndex - startLineIndex) *
              paragraphLineHeightPx;
        // A page break may give back the suppressible bottom space of the last
        // line it keeps, so this is the height the page really consumes.
        const segmentConsumedHeightPx = (
          startLineIndex: number,
          endLineIndex: number
        ): number => {
          const fullHeightPx = segmentLinesHeightPx(
            startLineIndex,
            endLineIndex
          );
          if (!paragraphLineMetrics || endLineIndex <= startLineIndex) {
            return fullHeightPx;
          }
          const lastLineMetrics = paragraphLineMetrics[endLineIndex - 1];
          const suppressibleBottomSpacePx = Math.max(
            0,
            lastLineMetrics?.suppressibleBottomSpacePx ?? 0
          );
          return Math.max(0, fullHeightPx - suppressibleBottomSpacePx);
        };

        let lineCursor = 0;
        let isFirstSegment = true;
        while (lineCursor < paragraphLineCount) {
          const linesRemaining = paragraphLineCount - lineCursor;
          const nextForcedBreakLineIndex = forcedPageBreakLineIndexes.find(
            (lineIndex) => lineIndex > lineCursor
          );
          const linesAvailableBeforeForcedBreak =
            nextForcedBreakLineIndex === undefined
              ? linesRemaining
              : Math.max(0, nextForcedBreakLineIndex - lineCursor);
          const topSpacingPx = isFirstSegment
            ? pageConsumedHeightPx > 0
              ? beforeSpacingPx
              : 0
            : 0;
          const mustKeepBottomSpacing =
            linesRemaining <= minLinesPerSegment &&
            nextForcedBreakLineIndex === undefined;
          const bottomSpacingPx = mustKeepBottomSpacing ? afterSpacingPx : 0;
          const remainingHeightPx = Math.max(0, currentPageContentHeightPx - pageConsumedHeightPx);
          const allRemainingSegmentReservePx = resolveParagraphSegmentNonFlowReservePx({
            startLineIndex: lineCursor,
            endLineIndex: paragraphLineCount,
            totalLineCount: paragraphLineCount,
            lineHeightPx: paragraphLineHeightPx
          });
          const allRemainingHeightPx =
            topSpacingPx +
            segmentLinesHeightPx(lineCursor, paragraphLineCount) +
            bottomSpacingPx;

          if (
            nextForcedBreakLineIndex === undefined &&
            allRemainingHeightPx + allRemainingSegmentReservePx <= remainingHeightPx
          ) {
            currentPageSegments.push({
              nodeIndex,
              paragraphLineRange: {
                startLineIndex: lineCursor,
                endLineIndex: paragraphLineCount,
                totalLineCount: paragraphLineCount,
                lineHeightPx: paragraphLineHeightPx
              }
            });
            pageConsumedHeightPx += allRemainingHeightPx;
            lineCursor = paragraphLineCount;
            break;
          }

          const maxLinesThisPage = Math.max(
            0,
            Math.min(
              linesRemaining - minLinesPerSegment,
              linesAvailableBeforeForcedBreak
            )
          );
          const continuingSegmentReservePx = resolveParagraphSegmentNonFlowReservePx({
            startLineIndex: lineCursor,
            endLineIndex: Math.min(paragraphLineCount, lineCursor + maxLinesThisPage),
            totalLineCount: paragraphLineCount,
            lineHeightPx: paragraphLineHeightPx
          });
          const availableForLinesPx = Math.max(
            0,
            remainingHeightPx - topSpacingPx - continuingSegmentReservePx
          );
          let linesThatFit = paragraphLineMetrics
            ? Math.max(
                0,
                constrainParagraphLineSplit(
                  paragraphLineMetrics,
                  lineCursor,
                  Math.min(paragraphLineCount, lineCursor + maxLinesThisPage),
                  availableForLinesPx
                ) - lineCursor
              )
            : Math.floor(availableForLinesPx / paragraphLineHeightPx);
          linesThatFit = Math.min(linesThatFit, maxLinesThisPage);

          if (linesThatFit < minLinesPerSegment) {
            const forcedBreakBeforeWidowMinimum =
              nextForcedBreakLineIndex !== undefined &&
              linesAvailableBeforeForcedBreak > 0 &&
              linesAvailableBeforeForcedBreak < minLinesPerSegment;
            if (forcedBreakBeforeWidowMinimum) {
              // An explicit page break outranks widow/orphan control.
              linesThatFit = Math.max(1, linesThatFit);
            } else if (currentPageSegments.length > 0) {
              startNextPage();
              pageConsumedHeightPx = 0;
              currentPageContentHeightPx = resolvePageContentHeightPx(
                currentPageIndex,
                nodeMetrics.pageContentHeightPx
              );
              continue;
            } else {
              const fallbackLines = Math.max(
                1,
                paragraphLineMetrics
                  ? constrainParagraphLineSplit(
                      paragraphLineMetrics,
                      lineCursor,
                      Math.min(paragraphLineCount, lineCursor + linesRemaining),
                      Math.max(1, availableForLinesPx)
                    ) - lineCursor
                  : Math.floor(
                      Math.max(1, availableForLinesPx) / paragraphLineHeightPx
                    )
              );
              linesThatFit = Math.max(
                1,
                Math.min(
                  maxLinesThisPage > 0
                    ? maxLinesThisPage
                    : linesAvailableBeforeForcedBreak > 0
                    ? linesAvailableBeforeForcedBreak
                    : linesRemaining,
                  fallbackLines
                )
              );
            }
          }

          let segmentEndLineIndex = Math.min(
            paragraphLineCount,
            lineCursor + linesThatFit,
            nextForcedBreakLineIndex ?? paragraphLineCount
          );
          if (nextForcedBreakLineIndex !== undefined) {
            linesThatFit = Math.min(
              linesThatFit,
              Math.max(1, linesAvailableBeforeForcedBreak)
            );
            segmentEndLineIndex = Math.min(
              segmentEndLineIndex,
              nextForcedBreakLineIndex
            );
          }
          while (linesThatFit > minLinesPerSegment) {
            const segmentReservePx = resolveParagraphSegmentNonFlowReservePx({
              startLineIndex: lineCursor,
              endLineIndex: segmentEndLineIndex,
              totalLineCount: paragraphLineCount,
              lineHeightPx: paragraphLineHeightPx
            });
            if (
              topSpacingPx +
                segmentConsumedHeightPx(lineCursor, segmentEndLineIndex) +
                segmentReservePx <=
              remainingHeightPx
            ) {
              break;
            }
            linesThatFit -= 1;
            segmentEndLineIndex = Math.min(
              paragraphLineCount,
              lineCursor + linesThatFit,
              nextForcedBreakLineIndex ?? paragraphLineCount
            );
          }
          currentPageSegments.push({
            nodeIndex,
            paragraphLineRange: {
              startLineIndex: lineCursor,
              endLineIndex: segmentEndLineIndex,
              totalLineCount: paragraphLineCount,
              lineHeightPx: paragraphLineHeightPx
            }
          });

          pageConsumedHeightPx +=
            topSpacingPx +
            segmentConsumedHeightPx(lineCursor, segmentEndLineIndex);
          lineCursor = segmentEndLineIndex;
          isFirstSegment = false;

          if (lineCursor < paragraphLineCount) {
            startNextPage();
            pageConsumedHeightPx = 0;
            currentPageContentHeightPx = resolvePageContentHeightPx(
              currentPageIndex,
              nodeMetrics.pageContentHeightPx
            );
          }
        }
        continue;
      }

      let requiredHeightPx = consumedNodeHeightPx;
      let keepNextChainEndNodeIndex = -1;
      if (
        node.style?.keepNext === true &&
        !nodeIsWithinCommittedKeepNextChain &&
        callbacks.paragraphHasVisibleText(node)
      ) {
        let chainCursor = nodeIndex;
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
              minParagraphLineHeightPx,
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
          // OOXML keeps both spacing values; no max()-style collapsing between
          // paragraphs that share a page.
          requiredHeightPx += Math.max(1, nextRawHeightPx);
          requiredHeightPx += keepNextParagraphReservePx(
            currentChainNode,
            nextChainNode,
            callbacks,
            chainMetrics.docGridLinePitchPx
          );
        }
        keepNextChainEndNodeIndex = chainCursor;
      }

      const remainingHeightPx = currentPageContentHeightPx - pageConsumedHeightPx;
      if (
        pageConsumedHeightPx > 0 &&
        requiredHeightPx > remainingHeightPx + pageOverflowTolerancePx
      ) {
        startNextPage();
        pageConsumedHeightPx = 0;
        currentPageContentHeightPx = resolvePageContentHeightPx(
          currentPageIndex,
          nodeMetrics.pageContentHeightPx
        );
      }

      if (pageConsumedHeightPx === 0 && keepNextChainEndNodeIndex > nodeIndex) {
        committedKeepNextChainEndNodeIndex = keepNextChainEndNodeIndex;
      }
      currentPageSegments.push({ nodeIndex });
      pageConsumedHeightPx += resolveParagraphConsumedHeightPx(
        rawNodeHeightPx,
        paragraphBeforeSpacingPx(node),
        pageConsumedHeightPx <= 0
      );
      continue;
    }

    const measuredRowHeightsPx = normalizedMeasuredTableRowHeights(
      options?.measuredTableRowHeightsByNodeIndex?.[nodeIndex],
      node.rows.length,
      minParagraphLineHeightPx
    );
    const estimatedRowHeightsPx =
      measuredRowHeightsPx ??
      estimatedRowHeightsByTableNodeIndex.get(nodeIndex) ??
      callbacks.estimateTableRowHeightsPx(
        node,
        nodeMetrics.pageContentWidthPx,
        numberingDefinitions,
        nodeMetrics.docGridLinePitchPx
      );
    if (!measuredRowHeightsPx && !estimatedRowHeightsByTableNodeIndex.has(nodeIndex)) {
      estimatedRowHeightsByTableNodeIndex.set(nodeIndex, estimatedRowHeightsPx);
    }

    if (estimatedRowHeightsPx.length === 0) {
      currentPageSegments.push({ nodeIndex });
      continue;
    }

    const tableExplicitPageBreakInfo = collectTableExplicitPageBreakInfo(node);
    const tableBreakStartRows = tableExplicitPageBreakInfo.startRowIndexes;
    if (tableBreakStartRows.includes(0) && currentPageSegments.length > 0) {
      startNextPage();
      pageConsumedHeightPx = 0;
      currentPageContentHeightPx = resolvePageContentHeightPx(
        currentPageIndex,
        nodeMetrics.pageContentHeightPx
      );
    }

    const headerRowCount = tableHeaderRowCount(node);
    const headerRowsHeightPx =
      headerRowCount > 0
        ? sumEstimatedTableRowHeightsPx(
            estimatedRowHeightsPx,
            0,
            headerRowCount,
            minParagraphLineHeightPx
          )
        : 0;
    const rowSliceStartOffsetPxByRowIndex = new Map<number, number>();

    let rowStartIndex = 0;
    let tableBreakStartRowCursor = 0;
    while (rowStartIndex < estimatedRowHeightsPx.length) {
      const startsInsideHeaderRows = rowStartIndex < headerRowCount;
      const repeatsHeaderRows =
        headerRowCount > 0 &&
        headerRowCount < estimatedRowHeightsPx.length &&
        rowStartIndex > 0 &&
        !startsInsideHeaderRows;
      const headerReservePx = repeatsHeaderRows
        ? Math.min(
            headerRowsHeightPx,
            Math.max(0, currentPageContentHeightPx - minParagraphLineHeightPx)
          )
        : 0;
      const availableHeightPx = Math.max(
        0,
        currentPageContentHeightPx - pageConsumedHeightPx - headerReservePx
      );
      const currentRow = node.rows[rowStartIndex];
      const currentRowHeightPx = Math.max(
        minParagraphLineHeightPx,
        estimatedRowHeightsPx[rowStartIndex] ?? minParagraphLineHeightPx
      );
      const currentRowNeedsSlicing =
        currentRow !== undefined &&
        tableRowCanSplitAcrossPages(currentRow) &&
        currentRowHeightPx > availableHeightPx + pageOverflowTolerancePx;
      const fittedRowEndIndex = fitTableRowsWithinHeightPx(
        estimatedRowHeightsPx,
        rowStartIndex,
        availableHeightPx,
        pageConsumedHeightPx <= 0 && !currentRowNeedsSlicing,
        minParagraphLineHeightPx,
        pageOverflowTolerancePx
      );
      let rowEndIndex = fittedRowEndIndex;
      while (
        tableBreakStartRowCursor < tableBreakStartRows.length &&
        tableBreakStartRows[tableBreakStartRowCursor] <= rowStartIndex
      ) {
        tableBreakStartRowCursor += 1;
      }
      const forcedBreakRowIndex = tableBreakStartRows[tableBreakStartRowCursor];
      if (forcedBreakRowIndex !== undefined) {
        rowEndIndex = Math.min(rowEndIndex, forcedBreakRowIndex);
      }

      const remainingRowsAfterSegment = estimatedRowHeightsPx.length - rowEndIndex;
      const segmentRowCount = rowEndIndex - rowStartIndex;
      if (
        forcedBreakRowIndex === undefined &&
        remainingRowsAfterSegment === 1 &&
        segmentRowCount > 1
      ) {
        rowEndIndex = fittedRowEndIndex - 1;
      }

      const segmentRowRange = (
        segmentStartRowIndex: number,
        segmentEndRowIndex: number
      ): TableRowRange => ({
        startRowIndex: segmentStartRowIndex,
        endRowIndex: segmentEndRowIndex,
        repeatHeaderRowCount: repeatsHeaderRows ? headerRowCount : undefined
      });

      if (rowEndIndex <= rowStartIndex) {
        if (currentPageSegments.length > 0) {
          startNextPage();
          pageConsumedHeightPx = 0;
          currentPageContentHeightPx = resolvePageContentHeightPx(
            currentPageIndex,
            nodeMetrics.pageContentHeightPx
          );
          continue;
        }

        const row = node.rows[rowStartIndex];
        const rowHeightPx = Math.max(
          minParagraphLineHeightPx,
          estimatedRowHeightsPx[rowStartIndex] ?? minParagraphLineHeightPx
        );
        const rowStartOffsetPx = Math.min(
          rowSliceStartOffsetPxByRowIndex.get(rowStartIndex) ?? 0,
          Math.max(0, rowHeightPx - 1)
        );
        const availableSliceHeightPx = Math.max(
          minParagraphLineHeightPx,
          currentPageContentHeightPx - pageConsumedHeightPx - headerReservePx
        );
        const sliceHeightPx = Math.min(
          availableSliceHeightPx,
          rowHeightPx - rowStartOffsetPx
        );

        if (
          row &&
          tableRowCanSplitAcrossPages(row) &&
          sliceHeightPx < rowHeightPx - rowStartOffsetPx
        ) {
          currentPageSegments.push({
            nodeIndex,
            tableRowRange: segmentRowRange(rowStartIndex, rowStartIndex + 1),
            tableRowSlice: {
              rowIndex: rowStartIndex,
              startOffsetPx: rowStartOffsetPx,
              sliceHeightPx,
              totalRowHeightPx: rowHeightPx
            }
          });
          pageConsumedHeightPx += headerReservePx + sliceHeightPx;
          const nextRowOffsetPx = rowStartOffsetPx + sliceHeightPx;
          if (nextRowOffsetPx >= rowHeightPx) {
            rowSliceStartOffsetPxByRowIndex.delete(rowStartIndex);
            rowStartIndex += 1;
          } else {
            rowSliceStartOffsetPxByRowIndex.set(
              rowStartIndex,
              nextRowOffsetPx
            );
          }
        } else {
          const forcedEndIndex = Math.min(
            estimatedRowHeightsPx.length,
            rowStartIndex + 1
          );
          const forcedHeightPx = sumEstimatedTableRowHeightsPx(
            estimatedRowHeightsPx,
            rowStartIndex,
            forcedEndIndex,
            minParagraphLineHeightPx
          );
          const remainingRowHeightPx = Math.max(
            1,
            rowHeightPx - rowStartOffsetPx
          );
          currentPageSegments.push({
            nodeIndex,
            tableRowRange: segmentRowRange(rowStartIndex, forcedEndIndex),
            tableRowSlice:
              rowStartOffsetPx > 0
                ? {
                    rowIndex: rowStartIndex,
                    startOffsetPx: rowStartOffsetPx,
                    sliceHeightPx: remainingRowHeightPx,
                    totalRowHeightPx: rowHeightPx
                  }
                : undefined
          });
          pageConsumedHeightPx +=
            headerReservePx +
            (rowStartOffsetPx > 0 ? remainingRowHeightPx : forcedHeightPx);
          rowSliceStartOffsetPxByRowIndex.delete(rowStartIndex);
          rowStartIndex = forcedEndIndex;
        }

        if (rowStartIndex < estimatedRowHeightsPx.length) {
          startNextPage();
          pageConsumedHeightPx = 0;
          currentPageContentHeightPx = resolvePageContentHeightPx(
            currentPageIndex,
            nodeMetrics.pageContentHeightPx
          );
        }
        continue;
      }

      const segmentHeightPx = sumEstimatedTableRowHeightsPx(
        estimatedRowHeightsPx,
        rowStartIndex,
        rowEndIndex,
        minParagraphLineHeightPx
      );
      const coversWholeTable =
        rowStartIndex === 0 && rowEndIndex >= estimatedRowHeightsPx.length;
      currentPageSegments.push({
        nodeIndex,
        tableRowRange: coversWholeTable
          ? undefined
          : segmentRowRange(rowStartIndex, rowEndIndex)
      });
      pageConsumedHeightPx += headerReservePx + segmentHeightPx;
      rowStartIndex = rowEndIndex;

      if (rowStartIndex < estimatedRowHeightsPx.length) {
        startNextPage();
        pageConsumedHeightPx = 0;
        currentPageContentHeightPx = resolvePageContentHeightPx(
          currentPageIndex,
          nodeMetrics.pageContentHeightPx
        );
      }
    }
  }

  if (currentPageSegments.length > 0 || pages.length === 0) {
    pages.push(currentPageSegments);
  }

  return pages;
}

export function resolveDocumentPageSegmentStartNodeIndex(
  pageSegments: DocumentPageNodeSegment[]
): number | undefined {
  const firstSegment = pageSegments.find((segment) => Number.isFinite(segment.nodeIndex));
  return Number.isFinite(firstSegment?.nodeIndex) ? firstSegment?.nodeIndex : undefined;
}

export function scorePaginationAgainstStoredPageBreaks(
  pages: DocumentPageNodeSegment[][],
  storedBreakStartNodeIndexes: number[]
): number {
  const comparableBreakCount = Math.min(
    storedBreakStartNodeIndexes.length,
    Math.max(0, pages.length - 1)
  );
  if (comparableBreakCount <= 0) {
    return 0;
  }

  let score = 0;
  for (let breakIndex = 0; breakIndex < comparableBreakCount; breakIndex += 1) {
    const expectedStartNodeIndex = storedBreakStartNodeIndexes[breakIndex];
    const actualStartNodeIndex = resolveDocumentPageSegmentStartNodeIndex(
      pages[breakIndex + 1] ?? []
    );
    if (!Number.isFinite(actualStartNodeIndex)) {
      score -= 1000;
      continue;
    }

    const delta = Math.abs((actualStartNodeIndex as number) - expectedStartNodeIndex);
    score -= delta * 100;
    if (delta === 0) {
      score += 25;
    }
  }

  return score;
}
