// pagination-plan-core — column render, segment clip/boundary helpers, and
// page segmentation plan orchestration.
//
// Migrated from @extend-ai/react-docx editor.tsx lines:
//   11481-11561  (paragraphBeforeSpacingPx, paragraphAfterSpacingPx,
//                  paragraphWidowControlEnabled, paragraphIsOnlyExplicitPageBreak,
//                  resolveParagraphBeforeSpacingPx, paragraphCanSplitAcrossPages)
//   11990-12464  (resolveLineRangeWithinVerticalSlice … splitParagraphSegmentForColumnRender)
//   12466-12652  (buildRenderColumnSegmentsForPageSection)
//   12714-12757  (buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints)
//   13986-14006  (mergeTrailingPagesToTargetCount)

import type {
  DocModel,
  NumberingDefinitionSet,
  ParagraphNode,
  TableNode,
} from "../../engine/types";
import {
  paragraphHasLastRenderedPageBreak,
  paragraphHasPageBreakBefore,
} from "../../layout/pagination";
import type {
  DocumentPageNodeSegment,
  ParagraphLineRange,
} from "../../layout/page-segmentation-core";
import { sumEstimatedTableRowHeightsPx } from "../../layout/page-segmentation-table";
import {
  PAGE_OVERFLOW_TOLERANCE_PX,
  MIN_PARAGRAPH_LINE_HEIGHT_PX,
  TABLE_CELL_SLICE_FULLY_VISIBLE_BOTTOM_BUFFER_PX,
  PARAGRAPH_SEGMENT_TOP_BLEED_PX,
  PARAGRAPH_SEGMENT_DESCENDER_BLEED_PX,
  PARAGRAPH_SEGMENT_VISUAL_SAFETY_PX,
  PARAGRAPH_SEGMENT_FALLBACK_TOP_BLEED_MAX_PX,
  PARAGRAPH_SEGMENT_FALLBACK_BOTTOM_BLEED_MAX_PX,
  PARAGRAPH_SEGMENT_FALLBACK_VISUAL_SAFETY_PX,
} from "./constants";
import {
  estimateParagraphLineHeightPx,
  resolvePretextLineRangeContentHeightPx,
  paragraphBaseFontSizePx,
  resolveMeasureFontSizePx,
} from "./line-height";
import {
  estimateDocNodeHeightPx,
  estimateTableRowHeightsPx,
  estimateTableHeightPx,
  paragraphCanSplitAcrossPages,
} from "./line-height-table";
import { paragraphHasImage } from "./paragraph-inspect";
import {
  buildParagraphPretextLayoutSource,
  layoutParagraphPretextSource,
} from "./pretext-build";
import {
  effectiveParagraphAfterSpacingPx,
  effectiveParagraphBeforeSpacingPx,
} from "./pagination-plan-iterate";
import { paragraphHasExplicitColumnBreak } from "./paragraph-tracked";
import { paragraphBorderInsetPx } from "./style-block-css";
import { resolveListParagraphIndent } from "./xml-parsing-extra";
import { twipsToSignedPixels } from "./ooxml-helpers";

// ---------------------------------------------------------------------------
// Locally ported helpers (not yet available as standalone exports in helpers/)
// ---------------------------------------------------------------------------

interface TableRowSlice {
  rowIndex: number;
  startOffsetPx: number;
  sliceHeightPx: number;
  totalRowHeightPx: number;
}

/** `DocumentPageNodeSegment` with optional `tableRowSlice` field from the
 *  upstream editor. The layout layer currently does not produce row slices,
 *  but column-render paths may encounter them when segments cross module
 *  boundaries. */
type Seg = DocumentPageNodeSegment & { tableRowSlice?: TableRowSlice };

function paragraphMaxFontSizePx(paragraph: ParagraphNode): number {
  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  let maxFontSizePx = paragraphBaseFontPx;

  paragraph.children.forEach((child) => {
    if (child.type !== "text" && child.type !== "form-field") {
      return;
    }
    maxFontSizePx = Math.max(
      maxFontSizePx,
      resolveMeasureFontSizePx(child.style, paragraphBaseFontPx)
    );
  });

  return Math.max(1, Math.round(maxFontSizePx));
}

function paragraphAvailableTextWidthPx(
  paragraph: ParagraphNode,
  availableWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet
): number {
  const safeAvailableWidthPx = Math.max(24, Math.round(availableWidthPx));
  const resolvedIndent = resolveListParagraphIndent(
    paragraph,
    numberingDefinitions
  );
  const leftIndentPx = Math.max(
    0,
    twipsToSignedPixels(resolvedIndent?.leftTwips) ?? 0
  );
  const rightIndentPx = Math.max(
    0,
    twipsToSignedPixels(paragraph.style?.indent?.rightTwips) ?? 0
  );
  const firstLineIndentPx = twipsToSignedPixels(resolvedIndent?.firstLineTwips);
  const hangingIndentPx = twipsToSignedPixels(resolvedIndent?.hangingTwips);
  const firstLineDeltaPx =
    firstLineIndentPx ?? (hangingIndentPx ? -hangingIndentPx : 0);
  const textIndentReductionPx =
    Number.isFinite(firstLineDeltaPx) && (firstLineDeltaPx as number) > 0
      ? (firstLineDeltaPx as number)
      : 0;
  const leftBorderInsetPx = paragraphBorderInsetPx(
    paragraph.style?.borders?.left
  );
  const rightBorderInsetPx = paragraphBorderInsetPx(
    paragraph.style?.borders?.right
  );

  return Math.max(
    24,
    Math.round(
      safeAvailableWidthPx -
        leftIndentPx -
        rightIndentPx -
        textIndentReductionPx -
        leftBorderInsetPx -
        rightBorderInsetPx
    )
  );
}

function estimateParagraphLineCountWithinWidth(
  paragraph: ParagraphNode,
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet
): number {
  const textContent = paragraph.children
    .map((child) => {
      if (child.type === "text") {
        return child.text;
      }
      if (child.type === "form-field") {
        return "____";
      }
      return "";
    })
    .join("");
  if (!textContent) {
    return 1;
  }

  if (
    !Number.isFinite(availableWidthPx) ||
    (availableWidthPx as number) <= 0
  ) {
    return Math.max(1, textContent.split(/\r?\n/).length);
  }

  const effectiveWidthPx = paragraphAvailableTextWidthPx(
    paragraph,
    availableWidthPx as number,
    numberingDefinitions
  );
  const lineHeightPx = estimateParagraphLineHeightPx(paragraph);
  const roughCharsPerLine = Math.max(
    1,
    Math.floor(effectiveWidthPx / Math.max(1, lineHeightPx * 0.55))
  );
  let lines = 1;
  let lineCharCount = 0;
  for (const char of textContent) {
    if (char === "\n" || char === "\r") {
      lines += 1;
      lineCharCount = 0;
      continue;
    }
    lineCharCount += 1;
    if (lineCharCount > roughCharsPerLine) {
      lines += 1;
      lineCharCount = 1;
    }
  }
  return Math.max(1, lines);
}

// ---------------------------------------------------------------------------
// Paragraph segment identity / partial-line helpers
// ---------------------------------------------------------------------------

function paragraphSegmentHasPartialLineRange(
  paragraphLineRange?: ParagraphLineRange
): boolean {
  if (!paragraphLineRange) {
    return false;
  }

  return (
    paragraphLineRange.startLineIndex > 0 ||
    paragraphLineRange.endLineIndex < paragraphLineRange.totalLineCount
  );
}

function paragraphSegmentIdentityMatches(
  segment:
    | { nodeIndex: number; startLineIndex: number; endLineIndex: number }
    | undefined,
  nodeIndex: number,
  paragraphLineRange?: ParagraphLineRange
): boolean {
  if (!segment || !paragraphLineRange) {
    return false;
  }

  return (
    segment.nodeIndex === nodeIndex &&
    segment.startLineIndex === paragraphLineRange.startLineIndex &&
    segment.endLineIndex === paragraphLineRange.endLineIndex
  );
}

// ---------------------------------------------------------------------------
// Line range within a vertical slice
// ---------------------------------------------------------------------------

export function resolveLineRangeWithinVerticalSlice(
  lineTopOffsetsPx: number[],
  lineHeightPx: number,
  sliceTopPx: number,
  sliceBottomPx: number
): ParagraphLineRange | undefined {
  if (
    lineTopOffsetsPx.length === 0 ||
    !Number.isFinite(lineHeightPx) ||
    lineHeightPx <= 0
  ) {
    return undefined;
  }

  const safeSliceTopPx = Math.max(0, sliceTopPx);
  const safeSliceBottomPx = Math.max(safeSliceTopPx, sliceBottomPx);
  const sliceHasHeight = safeSliceBottomPx > safeSliceTopPx;
  let startLineIndex: number | undefined;
  let endLineIndex: number | undefined;

  for (let lineIndex = 0; lineIndex < lineTopOffsetsPx.length; lineIndex += 1) {
    const lineTopPx = lineTopOffsetsPx[lineIndex] ?? lineIndex * lineHeightPx;
    const lineBottomPx = lineTopPx + lineHeightPx;
    const lineBelongsToSlice =
      sliceHasHeight &&
      lineBottomPx > safeSliceTopPx + PAGE_OVERFLOW_TOLERANCE_PX &&
      lineBottomPx <= safeSliceBottomPx + PAGE_OVERFLOW_TOLERANCE_PX;
    if (lineBelongsToSlice) {
      if (startLineIndex === undefined) {
        startLineIndex = lineIndex;
      }
      endLineIndex = lineIndex + 1;
    }
  }

  if (
    startLineIndex === undefined ||
    endLineIndex === undefined ||
    endLineIndex <= startLineIndex
  ) {
    return undefined;
  }

  return {
    startLineIndex,
    endLineIndex,
    totalLineCount: lineTopOffsetsPx.length,
    lineHeightPx,
  };
}

// ---------------------------------------------------------------------------
// Table cell paragraph boundary helpers
// ---------------------------------------------------------------------------

export function resolveTableCellParagraphVisualBottomPx(params: {
  paragraphTopPx: number;
  paragraphHeightPx: number;
  textBottomPx: number;
}): number {
  return Math.max(
    Math.round(params.paragraphTopPx + params.paragraphHeightPx),
    Math.round(params.textBottomPx)
  );
}

export function tableCellParagraphFitsFullyWithinSlice(params: {
  sliceStartPx: number;
  sliceBottomPx: number;
  paragraphTopPx: number;
  paragraphBottomPx: number;
}): boolean {
  return (
    params.sliceStartPx <= params.paragraphTopPx + PAGE_OVERFLOW_TOLERANCE_PX &&
    params.sliceBottomPx >=
      params.paragraphBottomPx +
        TABLE_CELL_SLICE_FULLY_VISIBLE_BOTTOM_BUFFER_PX -
        PAGE_OVERFLOW_TOLERANCE_PX
  );
}

// ---------------------------------------------------------------------------
// Paragraph segment clip bleed — visual buffer for partial line segments
// ---------------------------------------------------------------------------

export function resolveParagraphSegmentClipBleedPx(
  paragraphLineRange?: ParagraphLineRange
): {
  topPx: number;
  bottomPx: number;
} {
  if (!paragraphSegmentHasPartialLineRange(paragraphLineRange)) {
    return {
      topPx: 0,
      bottomPx: 0,
    };
  }

  return {
    topPx:
      paragraphLineRange && paragraphLineRange.startLineIndex > 0
        ? Math.max(0, PARAGRAPH_SEGMENT_TOP_BLEED_PX)
        : 0,
    bottomPx:
      paragraphLineRange &&
      paragraphLineRange.endLineIndex < paragraphLineRange.totalLineCount
        ? Math.max(0, PARAGRAPH_SEGMENT_DESCENDER_BLEED_PX)
        : 0,
  };
}

export function resolveFallbackParagraphSegmentClipBleedPx(
  paragraph: ParagraphNode,
  paragraphLineRange?: ParagraphLineRange
): {
  topPx: number;
  bottomPx: number;
} {
  if (!paragraphSegmentHasPartialLineRange(paragraphLineRange)) {
    return {
      topPx: 0,
      bottomPx: 0,
    };
  }

  const lineHeightPx = Math.max(1, paragraphLineRange?.lineHeightPx ?? 0);
  const maxFontSizePx = paragraphMaxFontSizePx(paragraph);
  const glyphOvershootPx = Math.max(
    0,
    Math.ceil((maxFontSizePx - lineHeightPx) / 2)
  );
  const ascenderSafetyPx = Math.max(
    glyphOvershootPx,
    Math.ceil(lineHeightPx * 0.22)
  );

  return {
    topPx:
      paragraphLineRange && paragraphLineRange.startLineIndex > 0
        ? Math.min(
            PARAGRAPH_SEGMENT_FALLBACK_TOP_BLEED_MAX_PX,
            ascenderSafetyPx
          )
        : 0,
    bottomPx:
      paragraphLineRange &&
      paragraphLineRange.endLineIndex < paragraphLineRange.totalLineCount
        ? Math.min(
            PARAGRAPH_SEGMENT_FALLBACK_BOTTOM_BLEED_MAX_PX,
            glyphOvershootPx
          )
        : 0,
  };
}

// ---------------------------------------------------------------------------
// Non-flow reserve — safety margin for partial line segments
// ---------------------------------------------------------------------------

export function resolveParagraphSegmentNonFlowReservePx(
  paragraphLineRange?: ParagraphLineRange
): number {
  const bleed = resolveParagraphSegmentClipBleedPx(paragraphLineRange);
  if (bleed.topPx <= 0 && bleed.bottomPx <= 0) {
    return 0;
  }

  const lineHeightSafetyPx = Math.max(
    0,
    Math.ceil((paragraphLineRange?.lineHeightPx ?? 0) * 0.9)
  );
  return (
    Math.max(0, bleed.topPx) +
    Math.max(0, bleed.bottomPx) +
    Math.max(0, PARAGRAPH_SEGMENT_VISUAL_SAFETY_PX, lineHeightSafetyPx)
  );
}

function resolveFallbackParagraphSegmentNonFlowReservePx(
  paragraph: ParagraphNode,
  paragraphLineRange?: ParagraphLineRange
): number {
  const bleed = resolveFallbackParagraphSegmentClipBleedPx(
    paragraph,
    paragraphLineRange
  );
  if (bleed.topPx <= 0 && bleed.bottomPx <= 0) {
    return 0;
  }

  const lineHeightSafetyPx = Math.max(
    0,
    Math.ceil((paragraphLineRange?.lineHeightPx ?? 0) * 0.15)
  );
  return (
    Math.max(0, bleed.topPx) +
    Math.max(0, bleed.bottomPx) +
    Math.max(1, PARAGRAPH_SEGMENT_FALLBACK_VISUAL_SAFETY_PX, lineHeightSafetyPx)
  );
}

// ---------------------------------------------------------------------------
// Rendered segment height estimation
// ---------------------------------------------------------------------------

function estimateRenderedPageSegmentHeightPx(
  node: DocModel["nodes"][number],
  segment: Seg,
  model: DocModel,
  availableWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number,
  options?: {
    excludeWrappedFloatingImageFootprint?: boolean;
  }
): number {
  if (node.type === "paragraph") {
    const paragraphLineRange = segment.paragraphLineRange;
    if (paragraphLineRange) {
      const beforeSpacingPx =
        paragraphLineRange.startLineIndex === 0
          ? effectiveParagraphBeforeSpacingPx(
              model,
              segment.nodeIndex,
              node,
              segment.nodeIndex > 0 ? 1 : 0,
              false
            )
          : 0;
      const afterSpacingPx =
        paragraphLineRange.endLineIndex >= paragraphLineRange.totalLineCount
          ? effectiveParagraphAfterSpacingPx(
              model,
              segment.nodeIndex,
              node
            )
          : 0;
      const paragraphPretextSource = buildParagraphPretextLayoutSource(node, {
        allowExplicitLineBreakText: true,
        expandTabsForLayout: true,
      });
      const paragraphPretextLayout = paragraphPretextSource
        ? layoutParagraphPretextSource(
            node,
            paragraphPretextSource,
            paragraphAvailableTextWidthPx(
              node,
              availableWidthPx,
              numberingDefinitions
            ),
            Math.max(1, paragraphLineRange.lineHeightPx),
            []
          )
        : undefined;
      const segmentContentHeightPx =
        paragraphPretextLayout && paragraphPretextLayout.lineCount > 0
          ? resolvePretextLineRangeContentHeightPx(
              paragraphPretextLayout,
              paragraphLineRange.startLineIndex,
              paragraphLineRange.endLineIndex
            )
          : Math.max(
              1,
              paragraphLineRange.endLineIndex -
                paragraphLineRange.startLineIndex
            ) * Math.max(1, paragraphLineRange.lineHeightPx);
      return Math.max(
        1,
        beforeSpacingPx + segmentContentHeightPx + afterSpacingPx
      );
    }

    return Math.max(
      1,
      estimateDocNodeHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx
      )
    );
  }

  if (segment.tableRowRange) {
    if (segment.tableRowSlice) {
      return Math.max(
        MIN_PARAGRAPH_LINE_HEIGHT_PX,
        Math.round(segment.tableRowSlice.sliceHeightPx)
      );
    }
    const rowHeightsPx = estimateTableRowHeightsPx(
      node as TableNode,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    );
    return Math.max(
      MIN_PARAGRAPH_LINE_HEIGHT_PX,
      sumEstimatedTableRowHeightsPx(
        rowHeightsPx,
        segment.tableRowRange?.startRowIndex ?? 0,
        segment.tableRowRange?.endRowIndex ?? 0,
        MIN_PARAGRAPH_LINE_HEIGHT_PX
      )
    );
  }

  return Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX,
    estimateTableHeightPx(
      node as TableNode,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    )
  );
}

// ---------------------------------------------------------------------------
// Column render line range resolution
// ---------------------------------------------------------------------------

function resolveParagraphColumnRenderLineRange(
  paragraph: ParagraphNode,
  segment: DocumentPageNodeSegment,
  availableWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): ParagraphLineRange {
  const lineHeightPx = Math.max(
    1,
    segment.paragraphLineRange?.lineHeightPx ??
      estimateParagraphLineHeightPx(paragraph, docGridLinePitchPx)
  );
  if (segment.paragraphLineRange) {
    return {
      ...segment.paragraphLineRange,
      lineHeightPx,
    };
  }

  const paragraphTextWidthPx = paragraphAvailableTextWidthPx(
    paragraph,
    availableWidthPx,
    numberingDefinitions
  );
  const pretextSource = buildParagraphPretextLayoutSource(paragraph, {
    allowExplicitLineBreakText: true,
    expandTabsForLayout: true,
  });
  const pretextLayout = pretextSource
    ? layoutParagraphPretextSource(
        paragraph,
        pretextSource,
        paragraphTextWidthPx,
        lineHeightPx,
        []
      )
    : undefined;
  const totalLineCount =
    pretextLayout && pretextLayout.lineCount > 0
      ? pretextLayout.lineCount
      : estimateParagraphLineCountWithinWidth(
          paragraph,
          availableWidthPx,
          numberingDefinitions
        );

  return {
    startLineIndex: 0,
    endLineIndex: Math.max(1, totalLineCount),
    totalLineCount: Math.max(1, totalLineCount),
    lineHeightPx,
  };
}

// ---------------------------------------------------------------------------
// Paragraph segment split for column rendering
// ---------------------------------------------------------------------------

function splitParagraphSegmentForColumnRender(params: {
  paragraph: ParagraphNode;
  segment: Seg;
  model: DocModel;
  availableWidthPx: number;
  availableHeightPx: number;
  numberingDefinitions?: NumberingDefinitionSet;
  docGridLinePitchPx?: number;
}):
  | {
      currentSegment: Seg;
      currentHeightPx: number;
      remainderSegment: Seg;
    }
  | undefined {
  const {
    paragraph,
    segment,
    model,
    availableWidthPx,
    availableHeightPx,
    numberingDefinitions,
    docGridLinePitchPx,
  } = params;
  if (
    segment.tableRowRange ||
    segment.tableRowSlice ||
    paragraphHasExplicitColumnBreak(paragraph)
  ) {
    return undefined;
  }

  const fullLineRange = resolveParagraphColumnRenderLineRange(
    paragraph,
    segment,
    availableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx
  );
  const startLineIndex = Math.max(0, fullLineRange.startLineIndex);
  const endLineIndex = Math.max(startLineIndex, fullLineRange.endLineIndex);
  if (
    endLineIndex - startLineIndex < 2 ||
    !paragraphCanSplitAcrossPages(paragraph, fullLineRange.totalLineCount)
  ) {
    return undefined;
  }

  const safeAvailableHeightPx = Math.max(0, Math.round(availableHeightPx));
  let bestSegment: DocumentPageNodeSegment | undefined;
  let bestHeightPx = 0;

  for (
    let candidateEndLineIndex = startLineIndex + 1;
    candidateEndLineIndex < endLineIndex;
    candidateEndLineIndex += 1
  ) {
    const candidateSegment: DocumentPageNodeSegment = {
      ...segment,
      paragraphLineRange: {
        startLineIndex,
        endLineIndex: candidateEndLineIndex,
        totalLineCount: fullLineRange.totalLineCount,
        lineHeightPx: fullLineRange.lineHeightPx,
      },
    };
    const candidateHeightPx = estimateRenderedPageSegmentHeightPx(
      paragraph,
      candidateSegment,
      model,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    );
    if (
      candidateHeightPx >
      safeAvailableHeightPx + PAGE_OVERFLOW_TOLERANCE_PX
    ) {
      break;
    }

    bestSegment = candidateSegment;
    bestHeightPx = candidateHeightPx;
  }

  if (!bestSegment?.paragraphLineRange) {
    return undefined;
  }

  return {
    currentSegment: bestSegment,
    currentHeightPx: bestHeightPx,
    remainderSegment: {
      ...segment,
      paragraphLineRange: {
        startLineIndex: bestSegment.paragraphLineRange.endLineIndex,
        endLineIndex,
        totalLineCount: fullLineRange.totalLineCount,
        lineHeightPx: fullLineRange.lineHeightPx,
      },
    },
  };
}

// ---------------------------------------------------------------------------
// Column segment distribution
// ---------------------------------------------------------------------------

export function buildRenderColumnSegmentsForPageSection(
  model: DocModel,
  flowSegments: DocumentPageNodeSegment[],
  columnWidthsPx: number[],
  columnHeightPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPxByNodeIndex?: Map<number, number | undefined>,
  measuredParagraphOuterHeightsPxByNodeIndex?: Map<number, number>,
  balanceColumns = false,
  forceColumnBreakNodeIndexes?: Set<number>
): DocumentPageNodeSegment[][] {
  const columnCount = Math.max(1, columnWidthsPx.length);
  const columns = Array.from(
    { length: columnCount },
    () => [] as DocumentPageNodeSegment[]
  );
  const maxColumnHeightPx = Math.max(120, Math.round(columnHeightPx));
  const resolveSegmentHeightPx = (
    segment: Seg,
    columnWidthPx: number
  ): number => {
    const segmentNode = model.nodes[segment.nodeIndex];
    if (!segmentNode) {
      return MIN_PARAGRAPH_LINE_HEIGHT_PX;
    }

    const docGridLinePitchPx = docGridLinePitchPxByNodeIndex?.get(
      segment.nodeIndex
    );
    const measuredSegmentHeightPx =
      segmentNode.type === "paragraph" &&
      !segment.paragraphLineRange &&
      !segment.tableRowRange &&
      !segment.tableRowSlice
        ? measuredParagraphOuterHeightsPxByNodeIndex?.get(segment.nodeIndex)
        : undefined;
    return Number.isFinite(measuredSegmentHeightPx) &&
      (measuredSegmentHeightPx as number) > 0
      ? Math.max(1, Math.round(measuredSegmentHeightPx as number))
      : estimateRenderedPageSegmentHeightPx(
          segmentNode,
          segment,
          model,
          columnWidthPx,
          numberingDefinitions,
          docGridLinePitchPx
        );
  };
  const safeColumnHeightPx =
    balanceColumns && columnCount > 1
      ? Math.min(
          maxColumnHeightPx,
          Math.max(
            MIN_PARAGRAPH_LINE_HEIGHT_PX * 4,
            Math.ceil(
              flowSegments.reduce((totalHeightPx, segment) => {
                const columnWidthPx = Math.max(
                  120,
                  Math.round(columnWidthsPx[0] ?? 120)
                );
                return (
                  totalHeightPx +
                  resolveSegmentHeightPx(segment, columnWidthPx)
                );
              }, 0) / columnCount
            ) + PAGE_OVERFLOW_TOLERANCE_PX
          )
        )
      : maxColumnHeightPx;
  let columnIndex = 0;
  let consumedHeightPx = 0;

  const moveToNextColumn = (): boolean => {
    if (columnIndex + 1 >= columnCount) {
      return false;
    }

    columnIndex += 1;
    consumedHeightPx = 0;
    return true;
  };

  const pushSegment = (
    segment: DocumentPageNodeSegment,
    heightPx: number
  ): void => {
    columns[columnIndex]?.push(segment);
    consumedHeightPx += Math.max(1, Math.round(heightPx));
  };

  for (const rawSegment of flowSegments) {
    const flowSegment = rawSegment as Seg;
    const isNodeStartSegment =
      (flowSegment.paragraphLineRange?.startLineIndex ?? 0) === 0 &&
      (flowSegment.tableRowRange?.startRowIndex ?? 0) === 0 &&
      !flowSegment.tableRowSlice;
    if (
      isNodeStartSegment &&
      forceColumnBreakNodeIndexes?.has(flowSegment.nodeIndex) &&
      consumedHeightPx > 0
    ) {
      moveToNextColumn();
    }
    let pendingSegment: Seg | undefined = flowSegment;
    let splitGuard = 0;

    while (pendingSegment && splitGuard < 256) {
      splitGuard += 1;
      const currentSegment: Seg = pendingSegment;
      const segmentNode: DocModel["nodes"][number] | undefined =
        model.nodes[currentSegment.nodeIndex];
      if (!segmentNode) {
        columns[columnIndex]?.push(currentSegment);
        break;
      }

      const columnWidthPx = Math.max(
        120,
        Math.round(columnWidthsPx[columnIndex] ?? columnWidthsPx[0] ?? 120)
      );
      const docGridLinePitchPx = docGridLinePitchPxByNodeIndex?.get(
        currentSegment.nodeIndex
      );
      const segmentHeightPx = resolveSegmentHeightPx(
        currentSegment,
        columnWidthPx
      );
      const remainingHeightPx = Math.max(
        0,
        safeColumnHeightPx - consumedHeightPx
      );

      if (
        segmentHeightPx <= remainingHeightPx + PAGE_OVERFLOW_TOLERANCE_PX ||
        columnIndex + 1 >= columnCount
      ) {
        pushSegment(currentSegment, segmentHeightPx);
        pendingSegment = undefined;
        break;
      }

      const splitSegment:
        | {
            currentSegment: Seg;
            currentHeightPx: number;
            remainderSegment: Seg;
          }
        | undefined =
        segmentNode.type === "paragraph"
          ? splitParagraphSegmentForColumnRender({
              paragraph: segmentNode,
              segment: currentSegment,
              model,
              availableWidthPx: columnWidthPx,
              availableHeightPx: remainingHeightPx,
              numberingDefinitions,
              docGridLinePitchPx,
            })
          : undefined;
      if (splitSegment) {
        pushSegment(splitSegment.currentSegment, splitSegment.currentHeightPx);
        pendingSegment = splitSegment.remainderSegment;
        if (!moveToNextColumn()) {
          const remainderSegment = splitSegment.remainderSegment;
          const remainderHeightPx = estimateRenderedPageSegmentHeightPx(
            segmentNode,
            remainderSegment,
            model,
            columnWidthPx,
            numberingDefinitions,
            docGridLinePitchPx
          );
          pushSegment(remainderSegment, remainderHeightPx);
          pendingSegment = undefined;
        }
        continue;
      }

      if (!moveToNextColumn()) {
        pushSegment(currentSegment, segmentHeightPx);
        pendingSegment = undefined;
      }
    }
  }

  return columns;
}

// ---------------------------------------------------------------------------
// Page segment reconstruction from last-rendered break hints
// ---------------------------------------------------------------------------

export function buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints(
  model: DocModel
): DocumentPageNodeSegment[][] {
  if (model.nodes.length === 0) {
    return [];
  }

  const pageStartNodeIndexes: number[] = [0];
  model.nodes.forEach((node, nodeIndex) => {
    if (nodeIndex === 0 || node.type !== "paragraph") {
      return;
    }

    if (
      paragraphHasLastRenderedPageBreak(node) ||
      paragraphHasPageBreakBefore(node)
    ) {
      pageStartNodeIndexes.push(nodeIndex);
    }
  });

  const uniquePageStartNodeIndexes = [...new Set(pageStartNodeIndexes)].sort(
    (left, right) => left - right
  );
  const pages: DocumentPageNodeSegment[][] = [];

  uniquePageStartNodeIndexes.forEach((startNodeIndex, pageIndex) => {
    const nextStartNodeIndex =
      uniquePageStartNodeIndexes[pageIndex + 1] ?? model.nodes.length;
    const pageSegments: DocumentPageNodeSegment[] = [];
    for (
      let nodeIndex = startNodeIndex;
      nodeIndex < nextStartNodeIndex;
      nodeIndex += 1
    ) {
      pageSegments.push({ nodeIndex });
    }
    if (pageSegments.length > 0) {
      pages.push(pageSegments);
    }
  });

  return pages;
}

// ---------------------------------------------------------------------------
// Trailing page merge
// ---------------------------------------------------------------------------

export function mergeTrailingPagesToTargetCount(
  pages: DocumentPageNodeSegment[][],
  targetPageCount: number
): DocumentPageNodeSegment[][] {
  const safeTargetPageCount = Math.max(1, Math.round(targetPageCount));
  if (pages.length <= safeTargetPageCount) {
    return pages;
  }

  const merged = pages.map((pageSegments) => [...pageSegments]);
  while (merged.length > safeTargetPageCount) {
    const trailingPage = merged.pop();
    if (!trailingPage || merged.length === 0) {
      break;
    }

    merged[merged.length - 1] = [
      ...merged[merged.length - 1],
      ...trailingPage,
    ];
  }

  return merged;
}
