// pagination-plan-iterate — break collection, overflow estimation, and page
// break index iteration.
//
// Migrated from @extend-ai/react-docx editor.tsx lines:
//   8092-8135  (effectiveParagraphAfterSpacingPx, effectiveParagraphBeforeSpacingPx)
//   8709-8921  (section start page breaks, next-break lookup, trailing section tail)
//   11563-11680 (collectDocxEstimatedOverflowBreakStartNodeIndexes)
//   11922-11943 (collectDocxPageBreakStartNodeIndexes)
//   15004-15034 (keepNextPaginationReservePx)

import type { DocModel, NumberingDefinitionSet, ParagraphNode } from "../../engine/types";
import {
  paragraphAfterSpacingPx,
  paragraphBeforeSpacingPx,
  paragraphHasExplicitPageBreak,
  paragraphHasPageBreakBefore,
  paragraphStartsWithLastRenderedPageBreak,
  resolveDocumentSectionsFromMetadata,
  resolveParagraphBeforeSpacingPx,
  resolvePaginationSectionMetricsIndexForNodeIndex,
  sectionBreakAfterParagraphStartsNewPage,
  sectionBreakPropertiesStartNewPage,
  type PaginationSectionMetrics,
  collectDocxHardPageBreakStartNodeIndexes,
} from "../../layout/pagination";
import { MAX_TRAILING_SECTION_TAIL_OVERFLOW_PX, MAX_TRAILING_SECTION_TAIL_PARAGRAPHS, PAGE_OVERFLOW_TOLERANCE_PX } from "./constants";
import { estimateParagraphLineHeightPx } from "./line-height";
import { estimateDocNodeHeightPx } from "./line-height-table";
import { paragraphHasImage, paragraphText } from "./paragraph-inspect";
import {
  paragraphActsAsDecorativeBehindTextBackgroundOverlay,
  paragraphActsAsLeadingCoverLayoutOverlay,
  paragraphHasFormField,
  paragraphHasVisibleText,
} from "./paragraph-geometry";
import {
  paragraphActsAsSectionBreakCarryoverSpacer,
  paragraphActsAsTrailingRenderedPageBreakSpacer,
  paragraphCollapsesIntoPreviousParagraph,
  paragraphIsStructuralSectionBreakSpacer,
  paragraphsSuppressInterParagraphSpacing,
} from "./paragraph-tracked";

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

const docxSectionStartPageBreakNodeIndexesByModel = new WeakMap<
  DocModel,
  Set<number>
>();

function computeDocxSectionStartPageBreakNodeIndexes(
  model: DocModel
): Set<number> {
  const breaks = new Set<number>();
  const sections = resolveDocumentSectionsFromMetadata(model.metadata);
  for (
    let sectionIndex = 1;
    sectionIndex < sections.length;
    sectionIndex += 1
  ) {
    const section = sections[sectionIndex];
    const startNodeIndex = Math.max(0, Math.round(section.startNodeIndex));
    if (startNodeIndex <= 0 || startNodeIndex >= model.nodes.length) {
      continue;
    }

    const sectionPropertiesXml = section.sectionPropertiesXml;
    if (!sectionPropertiesXml) {
      continue;
    }

    if (sectionBreakPropertiesStartNewPage(sectionPropertiesXml)) {
      breaks.add(startNodeIndex);
    }
  }

  return breaks;
}

// ---------------------------------------------------------------------------
// Effective spacing — collapsed margin-aware helpers
// ---------------------------------------------------------------------------

export function effectiveParagraphAfterSpacingPx(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode
): number {
  const afterSpacingPx = paragraphAfterSpacingPx(paragraph);
  const nextNode = model.nodes[nodeIndex + 1];
  if (nextNode?.type !== "paragraph") {
    return afterSpacingPx;
  }

  return paragraphsSuppressInterParagraphSpacing(paragraph, nextNode)
    ? 0
    : afterSpacingPx;
}

export function effectiveParagraphBeforeSpacingPx(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode,
  pageConsumedHeightPx: number,
  suppressSpacingBeforeAfterPageBreak: boolean
): number {
  let beforeSpacingPx = resolveParagraphBeforeSpacingPx(
    model,
    nodeIndex,
    paragraph,
    pageConsumedHeightPx,
    suppressSpacingBeforeAfterPageBreak
  );

  if (pageConsumedHeightPx <= 0) {
    return beforeSpacingPx;
  }

  const previousNode = model.nodes[nodeIndex - 1];
  if (previousNode?.type !== "paragraph") {
    return beforeSpacingPx;
  }

  return paragraphsSuppressInterParagraphSpacing(previousNode, paragraph)
    ? 0
    : beforeSpacingPx;
}

// ---------------------------------------------------------------------------
// keepNext reserve
// ---------------------------------------------------------------------------

export function keepNextPaginationReservePx(
  paragraph: ParagraphNode,
  nextParagraph: ParagraphNode | undefined,
  docGridLinePitchPx?: number
): number {
  if (
    paragraph.style?.keepNext !== true ||
    !nextParagraph ||
    !Number.isFinite(paragraph.style?.headingLevel)
  ) {
    return 0;
  }

  const nextParagraphText = paragraphText(nextParagraph)
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
      estimateParagraphLineHeightPx(nextParagraph, docGridLinePitchPx) *
        (nextParagraph.style?.numbering ? 1 : 0.5)
    )
  );
}

// ---------------------------------------------------------------------------
// Section start page break collection
// ---------------------------------------------------------------------------

export function collectDocxSectionStartPageBreakNodeIndexes(
  model: DocModel
): Set<number> {
  const cached = docxSectionStartPageBreakNodeIndexesByModel.get(model);
  if (cached) {
    return cached;
  }
  const result = computeDocxSectionStartPageBreakNodeIndexes(model);
  docxSectionStartPageBreakNodeIndexesByModel.set(model, result);
  return result;
}

// ---------------------------------------------------------------------------
// Next hard break lookup
// ---------------------------------------------------------------------------

export function buildNextHardBreakStartNodeIndexLookup(
  nodeCount: number,
  hardBreakStartNodeIndexes: Set<number>
): number[] {
  const nextBreakStartNodeIndexes = new Array<number>(nodeCount).fill(-1);
  let nextBreakStartNodeIndex = -1;

  for (let nodeIndex = nodeCount - 1; nodeIndex >= 0; nodeIndex -= 1) {
    nextBreakStartNodeIndexes[nodeIndex] = nextBreakStartNodeIndex;
    if (hardBreakStartNodeIndexes.has(nodeIndex)) {
      nextBreakStartNodeIndex = nodeIndex;
    }
  }

  return nextBreakStartNodeIndexes;
}

// ---------------------------------------------------------------------------
// Trailing section tail
// ---------------------------------------------------------------------------

export function paragraphIsSimpleTrailingSectionTailCandidate(
  paragraph: ParagraphNode
): boolean {
  if (
    !paragraphHasVisibleText(paragraph) ||
    paragraphHasImage(paragraph) ||
    paragraphHasFormField(paragraph) ||
    paragraphHasExplicitPageBreak(paragraph) ||
    paragraphHasPageBreakBefore(paragraph) ||
    paragraphStartsWithLastRenderedPageBreak(paragraph) ||
    sectionBreakAfterParagraphStartsNewPage(paragraph)
  ) {
    return false;
  }

  return true;
}

export function shouldKeepTrailingSectionTailOnCurrentPage(
  model: DocModel,
  startNodeIndex: number,
  pageConsumedHeightPx: number,
  previousParagraphAfterPx: number,
  pageContentWidthPx: number,
  pageContentHeightPx: number,
  hardBreakStartNodeIndexes: Set<number>,
  sectionStartPageBreakNodeIndexes: Set<number>,
  nextHardBreakStartNodeIndexByNodeIndex: number[],
  estimateNodeHeightPx: (
    nodeIndex: number,
    node: DocModel["nodes"][number],
    pageContentWidthPx: number
  ) => number,
  resolveNodeBeforeSpacingPx: (
    nodeIndex: number,
    paragraph: ParagraphNode,
    consumedHeightPx: number
  ) => number,
  resolveNodeAfterSpacingPx: (
    nodeIndex: number,
    paragraph: ParagraphNode
  ) => number
): boolean {
  if (pageConsumedHeightPx <= 0 || startNodeIndex <= 0) {
    return false;
  }

  const nextHardBreakStartNodeIndex =
    nextHardBreakStartNodeIndexByNodeIndex[startNodeIndex] ?? -1;
  if (
    nextHardBreakStartNodeIndex <= startNodeIndex ||
    !hardBreakStartNodeIndexes.has(nextHardBreakStartNodeIndex) ||
    !sectionStartPageBreakNodeIndexes.has(nextHardBreakStartNodeIndex)
  ) {
    return false;
  }

  const remainingHeightPx = pageContentHeightPx - pageConsumedHeightPx;

  let tailConsumedHeightPx = 0;
  let tailPreviousParagraphAfterPx = previousParagraphAfterPx;
  let substantiveParagraphCount = 0;

  for (
    let nodeIndex = startNodeIndex;
    nodeIndex < nextHardBreakStartNodeIndex;
    nodeIndex += 1
  ) {
    const node = model.nodes[nodeIndex];
    if (
      node.type === "paragraph" &&
      paragraphIsStructuralSectionBreakSpacer(node)
    ) {
      tailPreviousParagraphAfterPx = 0;
      continue;
    }
    if (
      node.type === "paragraph" &&
      paragraphActsAsLeadingCoverLayoutOverlay(
        model,
        nodeIndex,
        node,
        pageContentWidthPx,
        pageContentHeightPx
      )
    ) {
      return false;
    }
    if (
      node.type === "paragraph" &&
      paragraphActsAsDecorativeBehindTextBackgroundOverlay(node)
    ) {
      return false;
    }
    if (
      node.type === "paragraph" &&
      paragraphCollapsesIntoPreviousParagraph(node, model.nodes[nodeIndex - 1])
    ) {
      continue;
    }
    if (node.type !== "paragraph") {
      return false;
    }
    if (!paragraphIsSimpleTrailingSectionTailCandidate(node)) {
      return false;
    }

    substantiveParagraphCount += 1;
    if (substantiveParagraphCount > MAX_TRAILING_SECTION_TAIL_PARAGRAPHS) {
      return false;
    }

    const directBeforeSpacingPx = paragraphBeforeSpacingPx(node);
    const directAfterSpacingPx = paragraphAfterSpacingPx(node);
    const nodeBeforeSpacingPx = resolveNodeBeforeSpacingPx(
      nodeIndex,
      node,
      pageConsumedHeightPx + tailConsumedHeightPx
    );
    const nodeAfterSpacingPx = resolveNodeAfterSpacingPx(nodeIndex, node);
    const rawNodeHeightPx = Math.max(
      1,
      estimateNodeHeightPx(nodeIndex, node, pageContentWidthPx) -
        directBeforeSpacingPx -
        directAfterSpacingPx +
        nodeBeforeSpacingPx +
        nodeAfterSpacingPx
    );
    const collapsedMarginPx =
      pageConsumedHeightPx + tailConsumedHeightPx > 0
        ? Math.min(tailPreviousParagraphAfterPx, nodeBeforeSpacingPx)
        : 0;
    const effectiveNodeHeightPx = Math.max(
      1,
      rawNodeHeightPx - collapsedMarginPx
    );
    tailConsumedHeightPx += effectiveNodeHeightPx;
    tailPreviousParagraphAfterPx = nodeAfterSpacingPx;
  }

  if (substantiveParagraphCount === 0) {
    return false;
  }

  const overflowPx = tailConsumedHeightPx - remainingHeightPx;
  return (
    overflowPx > PAGE_OVERFLOW_TOLERANCE_PX &&
    overflowPx <= MAX_TRAILING_SECTION_TAIL_OVERFLOW_PX
  );
}

// ---------------------------------------------------------------------------
// Estimated overflow break collection
// ---------------------------------------------------------------------------

function collectDocxEstimatedOverflowBreakStartNodeIndexes(
  model: DocModel,
  hardBreakStartNodeIndexes: Set<number>,
  pageContentHeightPx: number,
  pageContentWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  paginationMetricsBySection?: PaginationSectionMetrics[],
  options?: {
    suppressSpacingBeforeAfterPageBreak?: boolean;
  }
): Set<number> {
  const breaks = new Set<number>();
  if (!Number.isFinite(pageContentHeightPx) || pageContentHeightPx <= 0) {
    return breaks;
  }

  const fallbackMetrics: PaginationSectionMetrics = {
    startNodeIndex: 0,
    pageContentWidthPx: Math.max(120, Math.round(pageContentWidthPx)),
    pageContentHeightPx: Math.max(120, Math.round(pageContentHeightPx)),
    pageContentHeightMultiplier: 1,
    docGridLinePitchPx: undefined,
  };
  const metricsBySection = paginationMetricsBySection?.length
    ? paginationMetricsBySection
    : [fallbackMetrics];
  const sectionStartPageBreakNodeIndexes =
    collectDocxSectionStartPageBreakNodeIndexes(model);
  const nextHardBreakStartNodeIndexByNodeIndex =
    buildNextHardBreakStartNodeIndexLookup(
      model.nodes.length,
      hardBreakStartNodeIndexes
    );

  let pageConsumedHeightPx = 0;
  let previousParagraphAfterPx = 0;
  let currentMetricsIndex = 0;
  // Mirrors buildDocumentPageNodeSegments: once a keepNext chain head starts a
  // page, later chain members must not re-trigger a keep-induced break — that
  // would strand the head on a near-empty page.
  let committedKeepNextChainEndNodeIndex = -1;
  const suppressSpacingBeforeAfterPageBreak =
    options?.suppressSpacingBeforeAfterPageBreak ?? false;
  let currentPageContentHeightPx =
    metricsBySection[0]?.pageContentHeightPx ??
    fallbackMetrics.pageContentHeightPx;
  const projectConsumedHeightAcrossSectionMultipliers = (
    consumedHeightPx: number,
    fromMetrics: PaginationSectionMetrics | undefined,
    toMetrics: PaginationSectionMetrics | undefined
  ): number => {
    const safeConsumedHeightPx = Math.max(0, Math.round(consumedHeightPx));
    if (safeConsumedHeightPx <= 0) {
      return 0;
    }

    const fromMultiplier = Math.max(
      1,
      Math.round(fromMetrics?.pageContentHeightMultiplier ?? 1)
    );
    const toMultiplier = Math.max(
      1,
      Math.round(toMetrics?.pageContentHeightMultiplier ?? 1)
    );
    if (fromMultiplier === toMultiplier) {
      return safeConsumedHeightPx;
    }

    const approximateVisualDepthPx = safeConsumedHeightPx / fromMultiplier;
    return Math.max(0, Math.round(approximateVisualDepthPx * toMultiplier));
  };
  for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
    const previousMetricsIndex = currentMetricsIndex;
    currentMetricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
      metricsBySection,
      nodeIndex,
      currentMetricsIndex
    );
    const nodeMetrics =
      metricsBySection[currentMetricsIndex] ?? fallbackMetrics;
    if (nodeIndex > 0 && currentMetricsIndex !== previousMetricsIndex) {
      pageConsumedHeightPx = projectConsumedHeightAcrossSectionMultipliers(
        pageConsumedHeightPx,
        metricsBySection[previousMetricsIndex],
        nodeMetrics
      );
      currentPageContentHeightPx = nodeMetrics.pageContentHeightPx;
    }

    if (hardBreakStartNodeIndexes.has(nodeIndex)) {
      pageConsumedHeightPx = 0;
      previousParagraphAfterPx = 0;
      currentPageContentHeightPx = nodeMetrics.pageContentHeightPx;
    }

    const node = model.nodes[nodeIndex];
    if (
      node.type === "paragraph" &&
      (paragraphIsStructuralSectionBreakSpacer(node) ||
        paragraphActsAsSectionBreakCarryoverSpacer(model, nodeIndex, node))
    ) {
      previousParagraphAfterPx = 0;
      continue;
    }
    if (
      node.type === "paragraph" &&
      paragraphActsAsTrailingRenderedPageBreakSpacer(model, nodeIndex, node)
    ) {
      previousParagraphAfterPx = 0;
      continue;
    }
    if (
      node.type === "paragraph" &&
      paragraphCollapsesIntoPreviousParagraph(node, model.nodes[nodeIndex - 1])
    ) {
      continue;
    }
    const directNodeBeforeSpacingPx =
      node.type === "paragraph" ? paragraphBeforeSpacingPx(node) : 0;
    const directNodeAfterSpacingPx =
      node.type === "paragraph" ? paragraphAfterSpacingPx(node) : 0;
    const nodeBeforeSpacingPx =
      node.type === "paragraph"
        ? effectiveParagraphBeforeSpacingPx(
            model,
            nodeIndex,
            node,
            pageConsumedHeightPx,
            suppressSpacingBeforeAfterPageBreak
          )
        : 0;
    const nodeAfterSpacingPx =
      node.type === "paragraph"
        ? effectiveParagraphAfterSpacingPx(model, nodeIndex, node)
        : 0;
    const rawNodeHeightPx = Math.max(
      1,
      estimateDocNodeHeightPx(
        node,
        nodeMetrics.pageContentWidthPx,
        numberingDefinitions,
        nodeMetrics.docGridLinePitchPx
      ) -
        directNodeBeforeSpacingPx -
        directNodeAfterSpacingPx +
        nodeBeforeSpacingPx +
        nodeAfterSpacingPx
    );
    const collapsedMarginPx =
      node.type === "paragraph" && pageConsumedHeightPx > 0
        ? Math.min(previousParagraphAfterPx, nodeBeforeSpacingPx)
        : 0;
    const collapsedNodeHeightPx = Math.max(
      1,
      rawNodeHeightPx - collapsedMarginPx
    );

    let requiredHeightPx = collapsedNodeHeightPx;
    let keepNextChainEndNodeIndex = -1;

    if (
      node.type === "paragraph" &&
      node.style?.keepNext === true &&
      nodeIndex > committedKeepNextChainEndNodeIndex &&
      paragraphHasVisibleText(node)
    ) {
      let chainCursor = nodeIndex;
      let chainPreviousParagraphAfterPx = nodeAfterSpacingPx;
      while (chainCursor < model.nodes.length - 1) {
        const currentChainNode = model.nodes[chainCursor];
        if (
          currentChainNode.type !== "paragraph" ||
          currentChainNode.style?.keepNext !== true ||
          !paragraphHasVisibleText(currentChainNode)
        ) {
          break;
        }
        if (hardBreakStartNodeIndexes.has(chainCursor + 1)) {
          break;
        }
        const nextChainNode = model.nodes[chainCursor + 1];
        if (nextChainNode.type !== "paragraph") {
          break;
        }

        chainCursor += 1;
        const chainMetricsIndex =
          resolvePaginationSectionMetricsIndexForNodeIndex(
            metricsBySection,
            chainCursor,
            currentMetricsIndex
          );
        const chainMetrics =
          metricsBySection[chainMetricsIndex] ?? fallbackMetrics;
        const nextDirectBeforeSpacingPx =
          nextChainNode.type === "paragraph"
            ? paragraphBeforeSpacingPx(nextChainNode)
            : 0;
        const nextDirectAfterSpacingPx =
          nextChainNode.type === "paragraph"
            ? paragraphAfterSpacingPx(nextChainNode)
            : 0;
        const nextBeforeSpacingPx =
          nextChainNode.type === "paragraph"
            ? effectiveParagraphBeforeSpacingPx(
                model,
                chainCursor,
                nextChainNode,
                1,
                suppressSpacingBeforeAfterPageBreak
              )
            : 0;
        const nextAfterSpacingPx =
          nextChainNode.type === "paragraph"
            ? effectiveParagraphAfterSpacingPx(
                model,
                chainCursor,
                nextChainNode
              )
            : 0;
        const nextRawHeightPx = Math.max(
          1,
          estimateDocNodeHeightPx(
            nextChainNode,
            chainMetrics.pageContentWidthPx,
            numberingDefinitions,
            chainMetrics.docGridLinePitchPx
          ) -
            nextDirectBeforeSpacingPx -
            nextDirectAfterSpacingPx +
            nextBeforeSpacingPx +
            nextAfterSpacingPx
        );
        const collapsedChainMarginPx =
          nextChainNode.type === "paragraph"
            ? Math.min(chainPreviousParagraphAfterPx, nextBeforeSpacingPx)
            : 0;
        requiredHeightPx += Math.max(
          1,
          nextRawHeightPx - collapsedChainMarginPx
        );
        requiredHeightPx += keepNextPaginationReservePx(
          currentChainNode,
          nextChainNode,
          chainMetrics.docGridLinePitchPx
        );
        chainPreviousParagraphAfterPx = nextAfterSpacingPx;
      }
      keepNextChainEndNodeIndex = chainCursor;
    }

    const remainingHeightPx = currentPageContentHeightPx - pageConsumedHeightPx;
    const canKeepTrailingSectionTailOnCurrentPage =
      shouldKeepTrailingSectionTailOnCurrentPage(
        model,
        nodeIndex,
        pageConsumedHeightPx,
        previousParagraphAfterPx,
        nodeMetrics.pageContentWidthPx,
        currentPageContentHeightPx,
        hardBreakStartNodeIndexes,
        sectionStartPageBreakNodeIndexes,
        nextHardBreakStartNodeIndexByNodeIndex,
        (_candidateNodeIndex, candidateNode, candidatePageContentWidthPx) =>
          estimateDocNodeHeightPx(
            candidateNode,
            candidatePageContentWidthPx,
            numberingDefinitions,
            nodeMetrics.docGridLinePitchPx
          ),
        (candidateNodeIndex, candidateParagraph, consumedHeightPx) =>
          effectiveParagraphBeforeSpacingPx(
            model,
            candidateNodeIndex,
            candidateParagraph,
            consumedHeightPx,
            suppressSpacingBeforeAfterPageBreak
          ),
        (candidateNodeIndex, candidateParagraph) =>
          effectiveParagraphAfterSpacingPx(
            model,
            candidateNodeIndex,
            candidateParagraph
          )
      );
    if (
      pageConsumedHeightPx > 0 &&
      requiredHeightPx > remainingHeightPx + PAGE_OVERFLOW_TOLERANCE_PX &&
      !canKeepTrailingSectionTailOnCurrentPage
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
    previousParagraphAfterPx =
      node.type === "paragraph" ? nodeAfterSpacingPx : 0;
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

// ---------------------------------------------------------------------------
// Unified page break start node index collection
// ---------------------------------------------------------------------------

export function collectDocxPageBreakStartNodeIndexes(
  model: DocModel,
  pageContentHeightPx: number,
  pageContentWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  paginationMetricsBySection?: PaginationSectionMetrics[],
  options?: {
    suppressSpacingBeforeAfterPageBreak?: boolean;
  }
): Set<number> {
  const hardBreaks = collectDocxHardPageBreakStartNodeIndexes(model);
  const overflowBreaks = collectDocxEstimatedOverflowBreakStartNodeIndexes(
    model,
    hardBreaks,
    pageContentHeightPx,
    pageContentWidthPx,
    numberingDefinitions,
    paginationMetricsBySection,
    options
  );
  return new Set<number>([...hardBreaks, ...overflowBreaks]);
}
