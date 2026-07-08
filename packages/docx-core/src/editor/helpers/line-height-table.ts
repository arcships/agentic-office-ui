// Table row/cell height estimation, table row slice boundary resolution,
// paragraph height estimation, and paragraph spacing helpers for pagination.
// Upstream editor.tsx: lines 10056-11590.

import type {
  DocModel,
  NumberingDefinitionSet,
  ParagraphNode,
  TableCellContentNode,
  TableNode,
} from "../../engine/types";
import { twipsToPixels } from "../../viewer/section-layout";
import {
  paragraphHasExplicitPageBreak,
} from "../../layout/pagination";
import type { TableSpacingTwips } from "./cache-utils";
import {
  heightEstimateCacheKeyPx,
  paragraphEstimatedHeightBySourceXml,
  setCacheEntry,
  tableEstimatedHeightBySourceXml,
  tableEstimatedRowHeightsByNode,
} from "./cache-utils";
import {
  EMPTY_PARAGRAPH_EXTRA_HEIGHT_PX,
  MIN_PARAGRAPH_LINE_HEIGHT_PX,
  MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX,
  PAGE_OVERFLOW_TOLERANCE_PX,
  SPLITTABLE_TABLE_ROW_DEEP_CONTENT_NODE_THRESHOLD,
  SPLITTABLE_TABLE_ROW_ESTIMATE_EXTRA_LINE_COUNT,
  TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_MIN_ROWS,
  TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_RATIO,
  TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX,
  WORD_TABLE_CELL_PARAGRAPH_AFTER_TWIPS,
  WORD_TABLE_CELL_PARAGRAPH_AUTO_LINE_TWIPS,
  WORD_TABLE_CELL_PARAGRAPH_BEFORE_TWIPS,
} from "./constants";
import { isParagraphCellContentNode, paragraphHasImage } from "./paragraph-inspect";
import {
  paragraphHasFormField,
  paragraphHasVisibleText,
  paragraphIsAbsoluteFloatingImageAnchorOnly,
  paragraphAvailableTextWidthPx,
  shouldRenderAbsoluteFloatingImage,
  shouldRenderWrappedFloatingImage,
} from "./paragraph-geometry";
import {
  estimateParagraphLineHeightPx,
  paragraphDocGridSnapState,
  wrappedPretextParagraphBlockHeightPx,
} from "./line-height";
import { paragraphLineCountWithinWidth } from "./line-height-wrap";
import {
  buildParagraphPretextLayoutSource,
  layoutParagraphPretextSource,
} from "./pretext-build";
import { resolveParagraphDualWrappedTextLayout } from "./pretext-measure";
import {
  clampTableWidthPx,
  columnWidthsFromTableDefinition,
  defaultColumnWidthsPx,
  fitColumnWidthsToWidth,
  mergeTableSpacing,
  normalizeColumnWidthsPx,
  resolveTableSpacingPaddingPx,
  tableColumnCount,
} from "./table-utils";
import { resolveCollapsedTableHorizontalOuterBleedPx } from "./table-utils-extra";
import {
  paragraphIsEffectivelyEmpty,
  paragraphIsSectionBreakAnchorCarryover,
} from "./paragraph-tracked";
import { paragraphBorderInsetPx } from "./style-block-css";

// -- Paragraph height estimation (upstream 10056-10094, 10391-10539) --

function estimateWrappedFloatingImageFootprintPx(
  paragraph: ParagraphNode,
  image: import("../../engine/types").ImageRunNode
): number {
  if (!shouldRenderWrappedFloatingImage(image)) {
    return 0;
  }

  const wrapType = image.floating?.wrapType;
  const isTopAndBottomWrap = wrapType === "topAndBottom";
  const isImageOnlyAnchorParagraph = !paragraphHasVisibleText(paragraph);
  if (!isTopAndBottomWrap && !isImageOnlyAnchorParagraph) {
    return 0;
  }

  if (isTopAndBottomWrap && !isImageOnlyAnchorParagraph) {
    return 0;
  }

  const floating = image.floating;
  const imageHeightPx =
    Number.isFinite(image.heightPx) && (image.heightPx as number) > 0
      ? Math.round(image.heightPx as number)
      : Number.isFinite(image.widthPx) && (image.widthPx as number) > 0
      ? Math.round(image.widthPx as number)
      : MIN_PARAGRAPH_LINE_HEIGHT_PX;
  const distTPx = Math.max(0, Math.round(floating?.distTPx ?? 0));
  const distBPx = Math.max(0, Math.round(floating?.distBPx ?? 0));
  const verticalOffsetPx = Math.max(0, Math.round(floating?.yPx ?? 0));

  return Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX,
    imageHeightPx + distTPx + distBPx + verticalOffsetPx
  );
}

export function estimateParagraphHeightPx(
  paragraph: ParagraphNode,
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number {
  const sourceXml = paragraph.sourceXml;
  const baseWidthKey = heightEstimateCacheKeyPx(
    availableWidthPx,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  const widthKey: number | string = baseWidthKey;
  if (sourceXml) {
    const cachedByWidth = paragraphEstimatedHeightBySourceXml.get(sourceXml);
    const cached = cachedByWidth?.get(widthKey);
    if (cached !== undefined) {
      return cached;
    }
  }

  const beforeSpacing =
    twipsToPixels(paragraph.style?.spacing?.beforeTwips) ?? 0;
  const afterSpacing = twipsToPixels(paragraph.style?.spacing?.afterTwips) ?? 0;
  const lineHeightPx = estimateParagraphLineHeightPx(
    paragraph,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  const effectiveWidthPx =
    Number.isFinite(availableWidthPx) && (availableWidthPx as number) > 0
      ? paragraphAvailableTextWidthPx(
          paragraph,
          availableWidthPx as number,
          numberingDefinitions
        )
      : undefined;
  const dualWrappedLayout =
    effectiveWidthPx !== undefined
      ? resolveParagraphDualWrappedTextLayout(
          paragraph,
          effectiveWidthPx,
          lineHeightPx
        )
      : undefined;
  const lineCount = paragraphLineCountWithinWidth(
    paragraph,
    availableWidthPx,
    numberingDefinitions
  );
  const absoluteFloatingAnchorOnlyParagraph =
    paragraphIsAbsoluteFloatingImageAnchorOnly(paragraph);
  const collapsibleAbsoluteFloatingAnchorOnlyParagraph =
    absoluteFloatingAnchorOnlyParagraph &&
    paragraphIsSectionBreakAnchorCarryover(paragraph);
  const inlineImageHeightPx = paragraph.children.reduce((largest, child) => {
    if (child.type !== "image") {
      return largest;
    }
    if (
      shouldRenderAbsoluteFloatingImage(child) ||
      shouldRenderWrappedFloatingImage(child)
    ) {
      return largest;
    }
    return Math.max(largest, child.heightPx ?? 0);
  }, 0);
  const wrappedFloatingImageHeightPx = paragraph.children.reduce(
    (largest, child) => {
      if (child.type !== "image") {
        return largest;
      }
      return Math.max(
        largest,
        estimateWrappedFloatingImageFootprintPx(paragraph, child)
      );
    },
    0
  );
  const emptyParagraphHeightPx = paragraphIsEffectivelyEmpty(paragraph)
    ? lineHeightPx + EMPTY_PARAGRAPH_EXTRA_HEIGHT_PX
    : 0;
  const topBorderInsetPx = paragraphBorderInsetPx(
    paragraph.style?.borders?.top
  );
  const bottomBorderInsetPx = paragraphBorderInsetPx(
    paragraph.style?.borders?.bottom
  );
  const textFlowHeightPx = collapsibleAbsoluteFloatingAnchorOnlyParagraph
    ? 0
    : dualWrappedLayout
    ? wrappedPretextParagraphBlockHeightPx(dualWrappedLayout.layout)
    : lineHeightPx * lineCount;

  const contentHeightPx = Math.max(
    collapsibleAbsoluteFloatingAnchorOnlyParagraph ? 0 : lineHeightPx,
    textFlowHeightPx,
    inlineImageHeightPx,
    wrappedFloatingImageHeightPx,
    emptyParagraphHeightPx
  );
  const estimatedHeightPx = Math.max(
    1,
    beforeSpacing +
      afterSpacing +
      topBorderInsetPx +
      bottomBorderInsetPx +
      contentHeightPx
  );
  if (sourceXml) {
    const cachedByWidth =
      paragraphEstimatedHeightBySourceXml.get(sourceXml) ??
      new Map<number, number>();
    cachedByWidth.set(widthKey, estimatedHeightPx);
    setCacheEntry(
      paragraphEstimatedHeightBySourceXml,
      sourceXml,
      cachedByWidth
    );
  }
  return estimatedHeightPx;
}

// -- Paragraph spacing introspection --

function paragraphHasExplicitBeforeSpacing(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) return false;
  const spacingTag = xml.match(/<w:spacing\b[^>]*\/?>/i)?.[0];
  if (!spacingTag) return false;
  return /\bw:before(?:\s*=|Lines\s*=|Autospacing\s*=)/i.test(spacingTag);
}
export function paragraphHasExplicitSpacing(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) return false;
  const paragraphPropertiesXml =
    xml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/i)?.[0] ??
    xml.match(/<w:pPr\b[^>]*\/>/i)?.[0] ??
    "";
  if (!paragraphPropertiesXml) return false;
  return /<w:spacing\b[^>]*\/?>/i.test(paragraphPropertiesXml);
}

// -- Word-like table cell paragraph defaults --

export function wordLikeTableCellParagraph(
  paragraph: ParagraphNode,
  applyWordTableDefaults: boolean
): ParagraphNode {
  if (
    !applyWordTableDefaults ||
    !paragraph.sourceXml ||
    paragraphHasExplicitSpacing(paragraph)
  ) {
    return paragraph;
  }
  return {
    ...paragraph,
    sourceXml: undefined,
    style: {
      ...(paragraph.style ?? {}),
      spacing: {
        ...(paragraph.style?.spacing ?? {}),
        beforeTwips: WORD_TABLE_CELL_PARAGRAPH_BEFORE_TWIPS,
        afterTwips: WORD_TABLE_CELL_PARAGRAPH_AFTER_TWIPS,
        lineTwips: WORD_TABLE_CELL_PARAGRAPH_AUTO_LINE_TWIPS,
        lineRule: "auto",
      },
    },
  };
}
function suppressFirstTableCellParagraphTopSpacing(
  paragraph: ParagraphNode
): boolean {
  if (!paragraph.sourceXml) {
    const beforeTwips = paragraph.style?.spacing?.beforeTwips;
    return !(Number.isFinite(beforeTwips) && (beforeTwips as number) > 0);
  }
  return !paragraphHasExplicitBeforeSpacing(paragraph);
}

// -- Table cell content height estimation --

export function estimateTableCellContentHeightPx(
  nodeContent: TableCellContentNode[],
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  applyWordTableDefaults = false,
  docGridLinePitchPx?: number
): number {
  let paragraphIndex = 0;
  let expandedWithPretextLayout = false;
  let totalHeightPx = 0;

  for (const contentNode of nodeContent) {
    if (!isParagraphCellContentNode(contentNode)) {
      totalHeightPx += estimateTableHeightPx(
        contentNode,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx
      );
      continue;
    }

    const disableDocGridSnap =
      paragraphDocGridSnapState(contentNode) !== "snap";
    const paragraphForLayout = wordLikeTableCellParagraph(
      contentNode,
      applyWordTableDefaults
    );
    const baseHeight = estimateParagraphHeightPx(
      paragraphForLayout,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx,
      disableDocGridSnap
    );
    const lineHeightPx = Math.max(
      MIN_PARAGRAPH_LINE_HEIGHT_PX,
      estimateParagraphLineHeightPx(
        paragraphForLayout,
        docGridLinePitchPx,
        disableDocGridSnap
      )
    );
    const pretextSource = buildParagraphPretextLayoutSource(
      paragraphForLayout,
      { allowExplicitLineBreakText: true, expandTabsForLayout: true }
    );
    const paragraphTextWidthPx =
      typeof availableWidthPx === "number" && availableWidthPx > 0
        ? paragraphAvailableTextWidthPx(
            paragraphForLayout,
            availableWidthPx,
            numberingDefinitions
          )
        : undefined;
    const pretextLayout =
      pretextSource &&
      typeof paragraphTextWidthPx === "number" &&
      paragraphTextWidthPx > 0
        ? layoutParagraphPretextSource(
            paragraphForLayout,
            pretextSource,
            paragraphTextWidthPx,
            lineHeightPx,
            []
          )
        : undefined;
    const suppressTopSpacing =
      paragraphIndex === 0 &&
      suppressFirstTableCellParagraphTopSpacing(contentNode);
    paragraphIndex += 1;
    const beforeSpacing = suppressTopSpacing
      ? 0
      : twipsToPixels(paragraphForLayout.style?.spacing?.beforeTwips) ?? 0;
    const afterSpacing =
      twipsToPixels(paragraphForLayout.style?.spacing?.afterTwips) ?? 0;
    const topBorderInsetPx = paragraphBorderInsetPx(
      paragraphForLayout.style?.borders?.top
    );
    const bottomBorderInsetPx = paragraphBorderInsetPx(
      paragraphForLayout.style?.borders?.bottom
    );
    const pretextHeightPx = pretextLayout
      ? beforeSpacing +
        afterSpacing +
        topBorderInsetPx +
        bottomBorderInsetPx +
        wrappedPretextParagraphBlockHeightPx(pretextLayout)
      : 0;
    const resolvedBaseHeight =
      pretextHeightPx > 0 ? Math.max(baseHeight, pretextHeightPx) : baseHeight;
    if (pretextHeightPx > baseHeight + lineHeightPx / 2) {
      expandedWithPretextLayout = true;
    }
    if (!suppressTopSpacing) {
      totalHeightPx += resolvedBaseHeight;
      continue;
    }
    totalHeightPx += Math.max(1, resolvedBaseHeight - beforeSpacing);
  }

  return (
    totalHeightPx +
    (expandedWithPretextLayout ? Math.max(1, MIN_PARAGRAPH_LINE_HEIGHT_PX) : 0)
  );
}

// -- Table row splitting / flow introspection --

export function rowAllowsPageSplit(row: TableNode["rows"][number]): boolean {
  return row.style?.cantSplit !== true && row.style?.heightRule !== "exact";
}
export function rowHasDeepFlowContent(row: TableNode["rows"][number]): boolean {
  let blockNodeCount = 0;
  let nestedTableCount = 0;
  for (const cell of row.cells) {
    blockNodeCount += cell.nodes.length;
    for (const contentNode of cell.nodes) {
      if (contentNode.type === "table") nestedTableCount += 1;
    }
  }
  return (
    nestedTableCount > 0 ||
    blockNodeCount >= SPLITTABLE_TABLE_ROW_DEEP_CONTENT_NODE_THRESHOLD
  );
}
export function rowHasNestedTableContent(
  row: TableNode["rows"][number]
): boolean {
  return row.cells.some((cell) =>
    cell.nodes.some((contentNode) => contentNode.type === "table")
  );
}
export function capSplitFriendlyTableRowEstimatePx(
  row: TableNode["rows"][number],
  estimatedRowHeightPx: number,
  explicitHeightPx?: number,
  pageContentHeightPx?: number
): number {
  if (!rowAllowsPageSplit(row)) return estimatedRowHeightPx;
  if (!Number.isFinite(explicitHeightPx) || (explicitHeightPx as number) <= 0)
    return estimatedRowHeightPx;
  if (!rowHasDeepFlowContent(row)) return estimatedRowHeightPx;
  const safeExplicitHeightPx = Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX * 2,
    Math.round(explicitHeightPx as number)
  );
  const safePageContentHeightPx =
    Number.isFinite(pageContentHeightPx) && (pageContentHeightPx as number) > 0
      ? Math.max(
          MIN_PARAGRAPH_LINE_HEIGHT_PX * 4,
          Math.round(pageContentHeightPx as number)
        )
      : undefined;
  if (
    safePageContentHeightPx !== undefined &&
    estimatedRowHeightPx > safePageContentHeightPx + PAGE_OVERFLOW_TOLERANCE_PX
  ) {
    return estimatedRowHeightPx;
  }
  const cappedHeightPx =
    safeExplicitHeightPx +
    MIN_PARAGRAPH_LINE_HEIGHT_PX *
      SPLITTABLE_TABLE_ROW_ESTIMATE_EXTRA_LINE_COUNT;
  return Math.min(estimatedRowHeightPx, cappedHeightPx);
}

// -- Table style introspection --

export function tableStyleIdFromSourceXml(
  table: TableNode
): string | undefined {
  const sourceXml = table.sourceXml ?? "";
  if (!sourceXml) return undefined;
  const styleMatch = sourceXml.match(/<w:tblStyle\b[^>]*w:val="([^"]+)"/i);
  const styleId = styleMatch?.[1]?.trim();
  return styleId ? styleId : undefined;
}
export function tableHasVisibleBorders(table: TableNode): boolean {
  const borders = table.style?.borders;
  if (!borders) return false;
  return Object.values(borders).some(
    (border) => border && border.type !== "none" && border.type !== "nil"
  );
}
export function tableContainsParagraphsWithoutExplicitSpacing(
  table: TableNode
): boolean {
  return table.rows.some((row) =>
    row.cells.some((cell) =>
      cell.nodes.some(
        (node) =>
          node.type === "paragraph" && !paragraphHasExplicitSpacing(node)
      )
    )
  );
}
export function tableUsesWordLikeParagraphDefaults(
  table: TableNode
): boolean {
  const styleId = tableStyleIdFromSourceXml(table)?.toLowerCase();
  if (styleId === "tablegrid") return true;
  return (
    table.style?.layout === "fixed" &&
    tableHasVisibleBorders(table) &&
    tableContainsParagraphsWithoutExplicitSpacing(table)
  );
}

// -- Table row height estimation --

export function estimateTableRowHeightsPx(
  table: TableNode,
  maxAvailableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number,
  pageContentHeightPx?: number
): number[] {
  const baseCacheKey = heightEstimateCacheKeyPx(
    maxAvailableWidthPx,
    docGridLinePitchPx
  );
  const cachedByKey = tableEstimatedRowHeightsByNode.get(table);
  let baseRowHeights = cachedByKey?.get(baseCacheKey);

  if (!baseRowHeights) {
    baseRowHeights = computeTableCellDerivedRowHeightsPx(
      table,
      maxAvailableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    );
    const cacheByKey = cachedByKey ?? new Map<number, number[]>();
    cacheByKey.set(baseCacheKey, baseRowHeights);
    tableEstimatedRowHeightsByNode.set(table, cacheByKey);
  }

  return table.rows.map((row, rowIndex) => {
    let rowHeightPx = baseRowHeights[rowIndex] ?? 0;
    const explicitHeightPx = twipsToPixels(row.style?.heightTwips);
    if (explicitHeightPx && explicitHeightPx > 0) {
      rowHeightPx =
        row.style?.heightRule === "exact"
          ? explicitHeightPx
          : Math.max(rowHeightPx, explicitHeightPx);
    }
    rowHeightPx = capSplitFriendlyTableRowEstimatePx(
      row,
      rowHeightPx,
      explicitHeightPx,
      pageContentHeightPx
    );
    const paginationPaddingRatio =
      table.rows.length >= 35
        ? 1.32
        : table.rows.length >=
          TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_MIN_ROWS
        ? TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_RATIO
        : 1;
    const paddedRowHeightPx =
      paginationPaddingRatio > 1
        ? Math.round(rowHeightPx * paginationPaddingRatio)
        : rowHeightPx;
    return Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, paddedRowHeightPx);
  });
}
function computeTableCellDerivedRowHeightsPx(
  table: TableNode,
  maxAvailableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number[] {
  const defaultCellMargin = table.style?.cellMarginTwips;
  const columnCount = tableColumnCount(table);
  const tableWidthPx = twipsToPixels(table.style?.widthTwips);
  const rawTableColumnWidthsPx = (() => {
    const definedWidthsTwips = columnWidthsFromTableDefinition(table, columnCount);
    if (!definedWidthsTwips || definedWidthsTwips.length === 0) {
      return defaultColumnWidthsPx(columnCount, tableWidthPx);
    }
    const widthsPx = definedWidthsTwips.map(
      (widthTwips) => twipsToPixels(widthTwips) ?? 0
    );
    return normalizeColumnWidthsPx(widthsPx, columnCount, tableWidthPx, 1);
  })();
  const rawResolvedTableWidthPx =
    tableWidthPx ??
    rawTableColumnWidthsPx.reduce((sum, widthPx) => sum + widthPx, 0);
  const collapsedHorizontalBorderBleedPx =
    resolveCollapsedTableHorizontalOuterBleedPx(table, columnCount);
  const maxTableWidthPx =
    Number.isFinite(maxAvailableWidthPx) && (maxAvailableWidthPx as number) > 0
      ? Math.max(
          120,
          (maxAvailableWidthPx as number) - collapsedHorizontalBorderBleedPx
        )
      : undefined;
  const resolvedTableWidthPx = clampTableWidthPx(
    rawResolvedTableWidthPx,
    maxTableWidthPx
  );
  const tableColumnWidthsPx = fitColumnWidthsToWidth(
    rawTableColumnWidthsPx,
    resolvedTableWidthPx
  );
  const applyWordTableDefaults = tableUsesWordLikeParagraphDefaults(table);

  return table.rows.map((row) => {
    let columnCursor = 0;
    const rowHeightPx = row.cells.reduce((largest, cell) => {
      const columnSpan =
        cell.style?.gridSpan && cell.style.gridSpan > 1 ? cell.style.gridSpan : 1;
      const startColumnIndex = columnCursor;
      const endColumnIndex = Math.min(
        columnCount - 1,
        startColumnIndex + columnSpan - 1
      );
      columnCursor += columnSpan;
      const spanWidthPx = tableColumnWidthsPx
        .slice(startColumnIndex, endColumnIndex + 1)
        .reduce((sum, widthPx) => sum + widthPx, 0);
      const fallbackCellWidthPx =
        (resolvedTableWidthPx / Math.max(1, columnCount)) * columnSpan;
      const cellWidthPx = spanWidthPx > 0 ? spanWidthPx : fallbackCellWidthPx;
      const margin = cell.style?.marginTwips ?? defaultCellMargin;
      const resolvedPaddingPx = resolveTableSpacingPaddingPx(margin);
      const verticalPaddingPx = resolvedPaddingPx.top + resolvedPaddingPx.bottom;
      const horizontalPaddingPx =
        resolvedPaddingPx.left + resolvedPaddingPx.right;
      const contentWidthPx = Math.max(
        24,
        Math.round(cellWidthPx - horizontalPaddingPx)
      );
      const paragraphHeightPx = estimateTableCellContentHeightPx(
        cell.nodes,
        contentWidthPx,
        numberingDefinitions,
        applyWordTableDefaults,
        docGridLinePitchPx
      );
      return Math.max(largest, paragraphHeightPx + verticalPaddingPx);
    }, 0);
    return rowHeightPx;
  });
}

// -- Table row height CSS --

export function resolveTableRowHeightCss(
  row: TableNode["rows"][number],
  rowHeightPx?: number
): Record<string, string | number | undefined> | undefined {
  if (!Number.isFinite(rowHeightPx) || (rowHeightPx as number) <= 0)
    return undefined;
  const resolvedHeightPx = Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX,
    Math.round(rowHeightPx as number)
  );
  if (row.style?.heightRule === "exact") {
    return { height: `${resolvedHeightPx}px` };
  }
  return { height: `${resolvedHeightPx}px` };
}

// -- Table cell slice boundary layout --

export interface TableCellSliceBoundaryLayout {
  safeBoundariesPx: number[];
  contentBottomPx: number;
}
export function uniqueSortedPixelBoundaries(values: number[]): number[] {
  const sorted = values
    .filter((value) => Number.isFinite(value))
    .map((value) => Math.max(0, Math.round(value)))
    .sort((left, right) => left - right);
  const unique: number[] = [];
  for (const value of sorted) {
    const previous = unique[unique.length - 1];
    if (
      previous === undefined ||
      Math.abs(previous - value) > TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
    ) {
      unique.push(value);
    }
  }
  return unique;
}
function estimateParagraphBoundaryOffsetsPx(
  paragraph: ParagraphNode,
  availableWidthPx: number,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  applyWordTableDefaults: boolean,
  docGridLinePitchPx: number | undefined,
  paragraphIndex: number
): { heightPx: number; safeBoundariesPx: number[] } {
  const paragraphForLayout = wordLikeTableCellParagraph(
    paragraph,
    applyWordTableDefaults
  );
  const disableDocGridSnap = paragraphDocGridSnapState(paragraph) === "disable";
  const paragraphHeightPx = estimateParagraphHeightPx(
    paragraphForLayout,
    availableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  const suppressTopSpacing =
    paragraphIndex === 0 &&
    suppressFirstTableCellParagraphTopSpacing(paragraph);
  const beforeSpacingPx = suppressTopSpacing
    ? 0
    : twipsToPixels(paragraphForLayout.style?.spacing?.beforeTwips) ?? 0;
  const afterSpacingPx =
    twipsToPixels(paragraphForLayout.style?.spacing?.afterTwips) ?? 0;
  const topBorderInsetPx = paragraphBorderInsetPx(
    paragraphForLayout.style?.borders?.top
  );
  const bottomBorderInsetPx = paragraphBorderInsetPx(
    paragraphForLayout.style?.borders?.bottom
  );
  const lineHeightPx = Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX,
    estimateParagraphLineHeightPx(
      paragraphForLayout,
      docGridLinePitchPx,
      disableDocGridSnap
    )
  );
  const pretextSource = buildParagraphPretextLayoutSource(paragraphForLayout, {
    allowExplicitLineBreakText: true,
    expandTabsForLayout: true,
  });
  const paragraphTextWidthPx = paragraphAvailableTextWidthPx(
    paragraphForLayout,
    availableWidthPx,
    numberingDefinitions
  );
  const pretextLayout = pretextSource
    ? layoutParagraphPretextSource(
        paragraphForLayout,
        pretextSource,
        paragraphTextWidthPx,
        lineHeightPx,
        []
      )
    : undefined;
  const lineTopOffsetsPx = pretextLayout
    ? pretextLayout.lines.map((line) => Math.max(0, Math.round(line.y)))
    : Array.from(
        {
          length: Math.max(
            1,
            paragraphLineCountWithinWidth(
              paragraphForLayout,
              availableWidthPx,
              numberingDefinitions
            )
          ),
        },
        (_, lineIndex) => lineIndex * lineHeightPx
      );
  const textTopPx = beforeSpacingPx + topBorderInsetPx;
  const textHeightPx = pretextLayout
    ? wrappedPretextParagraphBlockHeightPx(pretextLayout)
    : lineTopOffsetsPx.length * lineHeightPx;
  const visualHeightPx = Math.max(
    1,
    beforeSpacingPx +
      topBorderInsetPx +
      textHeightPx +
      bottomBorderInsetPx +
      afterSpacingPx
  );
  const heightPx = Math.max(1, paragraphHeightPx, visualHeightPx);
  const lineBoundariesPx = lineTopOffsetsPx.map(
    (lineTopPx) => textTopPx + lineTopPx + lineHeightPx
  );
  return {
    heightPx,
    safeBoundariesPx: uniqueSortedPixelBoundaries([
      ...lineBoundariesPx,
      heightPx,
    ]),
  };
}
function estimateNestedTableBoundaryOffsetsPx(
  table: TableNode,
  availableWidthPx: number,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  docGridLinePitchPx: number | undefined
): { heightPx: number; safeBoundariesPx: number[] } {
  const rowHeightsPx = estimateTableRowHeightsPx(
    table,
    availableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx
  );
  const boundariesPx: number[] = [];
  let cursorPx = 0;
  for (const rowHeightPx of rowHeightsPx) {
    cursorPx += Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, rowHeightPx);
    boundariesPx.push(cursorPx);
  }
  return {
    heightPx: Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, cursorPx),
    safeBoundariesPx: uniqueSortedPixelBoundaries(boundariesPx),
  };
}
export function estimateTableCellSliceBoundaryLayoutPx(params: {
  cell: TableNode["rows"][number]["cells"][number];
  rowHeightPx: number;
  contentWidthPx: number;
  tableCellMarginTwips?: TableSpacingTwips;
  numberingDefinitions?: NumberingDefinitionSet;
  applyWordTableDefaults: boolean;
  docGridLinePitchPx?: number;
}): TableCellSliceBoundaryLayout {
  const {
    cell,
    rowHeightPx,
    contentWidthPx,
    tableCellMarginTwips,
    numberingDefinitions,
    applyWordTableDefaults,
    docGridLinePitchPx,
  } = params;
  const paddingPx = resolveTableSpacingPaddingPx(
    mergeTableSpacing(tableCellMarginTwips, cell.style?.marginTwips)
  );
  const localBoundariesPx = [0, paddingPx.top];
  let contentCursorPx = paddingPx.top;
  let paragraphIndex = 0;

  for (const contentNode of cell.nodes) {
    const layout =
      contentNode.type === "paragraph"
        ? estimateParagraphBoundaryOffsetsPx(
            contentNode,
            contentWidthPx,
            numberingDefinitions,
            applyWordTableDefaults,
            docGridLinePitchPx,
            paragraphIndex++
          )
        : estimateNestedTableBoundaryOffsetsPx(
            contentNode,
            contentWidthPx,
            numberingDefinitions,
            docGridLinePitchPx
          );
    localBoundariesPx.push(
      ...layout.safeBoundariesPx.map((boundaryPx) => contentCursorPx + boundaryPx)
    );
    contentCursorPx += layout.heightPx;
  }

  const contentBottomPx = contentCursorPx + paddingPx.bottom;
  const contentFlowHeightPx = Math.max(0, contentCursorPx - paddingPx.top);
  const availableContentHeightPx = Math.max(
    0,
    rowHeightPx - paddingPx.top - paddingPx.bottom
  );
  const extraVerticalSpacePx = Math.max(
    0,
    availableContentHeightPx - contentFlowHeightPx
  );
  const verticalOffsetPx =
    cell.style?.verticalAlign === "center"
      ? Math.round(extraVerticalSpacePx / 2)
      : cell.style?.verticalAlign === "bottom"
      ? extraVerticalSpacePx
      : 0;
  return {
    safeBoundariesPx: uniqueSortedPixelBoundaries(
      localBoundariesPx.map((boundaryPx) =>
        Math.min(rowHeightPx, boundaryPx + verticalOffsetPx)
      )
    ),
    contentBottomPx: Math.min(rowHeightPx, contentBottomPx + verticalOffsetPx),
  };
}
export function tableCellSliceBoundaryIsSafe(
  layout: TableCellSliceBoundaryLayout,
  boundaryPx: number
): boolean {
  if (boundaryPx <= TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX) return true;
  if (
    boundaryPx >=
    layout.contentBottomPx - TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
  )
    return true;
  return layout.safeBoundariesPx.some(
    (safeBoundaryPx) =>
      Math.abs(safeBoundaryPx - boundaryPx) <=
      TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
  );
}

export function resolveTableRowSliceHeightOnSafeBoundaryPx(params: {
  table: TableNode;
  rowIndex: number;
  rowHeightPx: number;
  rowSliceOffsetPx: number;
  preferredSliceHeightPx: number;
  maxAvailableWidthPx?: number;
  numberingDefinitions?: NumberingDefinitionSet;
  docGridLinePitchPx?: number;
}): number | undefined {
  const {
    table,
    rowIndex,
    rowHeightPx,
    rowSliceOffsetPx,
    preferredSliceHeightPx,
    maxAvailableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx,
  } = params;
  const row = table.rows[rowIndex];
  if (!row || !rowHasNestedTableContent(row)) return preferredSliceHeightPx;

  const sliceStartPx = Math.max(0, Math.round(rowSliceOffsetPx));
  const preferredSliceEndPx = Math.min(
    rowHeightPx,
    sliceStartPx + Math.max(0, Math.round(preferredSliceHeightPx))
  );
  if (
    preferredSliceEndPx >=
    rowHeightPx - TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
  ) {
    return Math.max(0, rowHeightPx - sliceStartPx);
  }

  const columnCount = tableColumnCount(table);
  const tableWidthPx = twipsToPixels(table.style?.widthTwips);
  const rawTableColumnWidthsPx = (() => {
    const definedWidthsTwips = columnWidthsFromTableDefinition(table, columnCount);
    if (!definedWidthsTwips || definedWidthsTwips.length === 0) {
      return defaultColumnWidthsPx(columnCount, tableWidthPx);
    }
    const widthsPx = definedWidthsTwips.map(
      (widthTwips) => twipsToPixels(widthTwips) ?? 0
    );
    return normalizeColumnWidthsPx(widthsPx, columnCount, tableWidthPx, 1);
  })();
  const rawResolvedTableWidthPx =
    tableWidthPx ??
    rawTableColumnWidthsPx.reduce((sum, widthPx) => sum + widthPx, 0);
  const collapsedHorizontalBorderBleedPx =
    resolveCollapsedTableHorizontalOuterBleedPx(table, columnCount);
  const maxTableWidthPx =
    Number.isFinite(maxAvailableWidthPx) && (maxAvailableWidthPx as number) > 0
      ? Math.max(
          120,
          (maxAvailableWidthPx as number) - collapsedHorizontalBorderBleedPx
        )
      : undefined;
  const resolvedTableWidthPx = clampTableWidthPx(
    rawResolvedTableWidthPx,
    maxTableWidthPx
  );
  const tableColumnWidthsPx = fitColumnWidthsToWidth(
    rawTableColumnWidthsPx,
    resolvedTableWidthPx
  );
  const applyWordTableDefaults = tableUsesWordLikeParagraphDefaults(table);
  const tableCellMarginTwips = table.style?.cellMarginTwips;
  const cellLayouts: TableCellSliceBoundaryLayout[] = [];
  const candidateBoundariesPx = [preferredSliceEndPx];
  let columnCursor = 0;

  for (const cell of row.cells) {
    const colSpanValue =
      cell.style?.gridSpan && cell.style.gridSpan > 1 ? cell.style.gridSpan : 1;
    const startColumnIndex = columnCursor;
    const endColumnIndex = Math.min(
      columnCount - 1,
      startColumnIndex + colSpanValue - 1
    );
    columnCursor += colSpanValue;
    const spannedWidthPx = tableColumnWidthsPx
      .slice(startColumnIndex, endColumnIndex + 1)
      .reduce((sum, widthPx) => sum + widthPx, 0);
    const fallbackCellWidthPx =
      (resolvedTableWidthPx / Math.max(1, columnCount)) * colSpanValue;
    const cellRenderedWidthPx =
      twipsToPixels(cell.style?.widthTwips) ??
      (spannedWidthPx > 0 ? spannedWidthPx : fallbackCellWidthPx);
    const cellPaddingPx = resolveTableSpacingPaddingPx(
      mergeTableSpacing(tableCellMarginTwips, cell.style?.marginTwips)
    );
    const cellContentWidthPx = Math.max(
      1,
      cellRenderedWidthPx - cellPaddingPx.left - cellPaddingPx.right
    );
    const cellLayout = estimateTableCellSliceBoundaryLayoutPx({
      cell,
      rowHeightPx,
      contentWidthPx: cellContentWidthPx,
      tableCellMarginTwips,
      numberingDefinitions,
      applyWordTableDefaults,
      docGridLinePitchPx,
    });
    cellLayouts.push(cellLayout);
    candidateBoundariesPx.push(...cellLayout.safeBoundariesPx);
  }

  const minimumSliceEndPx =
    sliceStartPx + Math.max(1, MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX);
  const candidatesPx = uniqueSortedPixelBoundaries(candidateBoundariesPx)
    .filter(
      (boundaryPx) =>
        boundaryPx >= minimumSliceEndPx &&
        boundaryPx <=
          preferredSliceEndPx + TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
    )
    .sort((left, right) => right - left);

  for (const candidatePx of candidatesPx) {
    if (
      cellLayouts.every((layout) =>
        tableCellSliceBoundaryIsSafe(layout, candidatePx)
      )
    ) {
      return Math.max(0, candidatePx - sliceStartPx);
    }
  }
  return undefined;
}

// -- Table height estimation --

export function estimateTableHeightPx(
  table: TableNode,
  maxAvailableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number {
  const sourceXml = table.sourceXml;
  const widthKey = heightEstimateCacheKeyPx(maxAvailableWidthPx, docGridLinePitchPx);
  if (sourceXml) {
    const cachedByWidth = tableEstimatedHeightBySourceXml.get(sourceXml);
    const cached = cachedByWidth?.get(widthKey);
    if (cached !== undefined) return cached;
  }
  const estimatedRowsHeightPx = estimateTableRowHeightsPx(
    table,
    maxAvailableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx
  ).reduce(
    (sum, rowHeightPx) =>
      sum + Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, rowHeightPx),
    0
  );
  const estimatedHeightPx = Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX * 2,
    estimatedRowsHeightPx
  );
  if (sourceXml) {
    const cachedByWidth =
      tableEstimatedHeightBySourceXml.get(sourceXml) ?? new Map<number, number>();
    cachedByWidth.set(widthKey, estimatedHeightPx);
    setCacheEntry(tableEstimatedHeightBySourceXml, sourceXml, cachedByWidth);
  }
  return estimatedHeightPx;
}

export function estimateDocNodeHeightPx(
  node: DocModel["nodes"][number],
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number {
  return node.type === "paragraph"
    ? estimateParagraphHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx,
        undefined
      )
    : estimateTableHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx
      );
}

// -- Paragraph spacing helpers (shared with pagination-plan) --
// paragraphBeforeSpacingPx / paragraphAfterSpacingPx / resolveParagraphBeforeSpacingPx
// live in ../../layout/pagination. Only helpers missing from pagination are here.

export function paragraphWidowControlEnabled(
  paragraph: ParagraphNode
): boolean {
  return paragraph.style?.widowControl !== false;
}

export function paragraphIsOnlyExplicitPageBreak(
  paragraph: ParagraphNode
): boolean {
  if (!paragraphHasExplicitPageBreak(paragraph)) return false;
  return (
    !paragraphHasVisibleText(paragraph) &&
    !paragraphHasImage(paragraph) &&
    !paragraphHasFormField(paragraph)
  );
}

export function paragraphCanSplitAcrossPages(
  paragraph: ParagraphNode,
  lineCount: number,
  options?: {
    allowKeepLinesOverflow?: boolean;
    allowKeepNextOverflow?: boolean;
    allowImageParagraphSplit?: boolean;
  }
): boolean {
  if (lineCount < 2) return false;
  if (paragraph.style?.keepLines === true && !options?.allowKeepLinesOverflow)
    return false;
  if (paragraph.style?.keepNext === true && !options?.allowKeepNextOverflow)
    return false;
  if (paragraphHasImage(paragraph) && !options?.allowImageParagraphSplit)
    return false;
  if (paragraphHasFormField(paragraph)) return false;
  return true;
}
