// Table row slice resolution, table height estimation, and paragraph spacing
// helpers for pagination. Split from line-height-table.ts to stay within the
// ≤1000-line constraint.
// Upstream editor.tsx: lines 10768-10997, 11514-11590.

import type {
  DocModel,
  NumberingDefinitionSet,
  ParagraphNode,
  TableNode,
} from "../../engine/types";
import { twipsToPixels } from "../../viewer/section-layout";
import {
  paragraphHasExplicitPageBreak,
} from "../../layout/pagination";
import type { TableSpacingTwips } from "./cache-utils";
import {
  heightEstimateCacheKeyPx,
  setCacheEntry,
  tableEstimatedHeightBySourceXml,
} from "./cache-utils";
import {
  tabGridLayoutDefaults,
  type DocumentLayoutDefaults,
} from "../../engine/layout-defaults";
import {
  MIN_PARAGRAPH_LINE_HEIGHT_PX,
  MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX,
  TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX,
} from "./constants";
import {
  paragraphHasImage,
} from "./paragraph-inspect";
import {
  paragraphHasFormField,
  paragraphHasVisibleText,
} from "./paragraph-geometry";
import {
  estimateTableCellSliceBoundaryLayoutPx,
  tableCellSliceBoundaryIsSafe,
  type TableCellSliceBoundaryLayout,
  tableUsesWordLikeParagraphDefaults,
  rowHasNestedTableContent,
  estimateTableRowHeightsPx,
  estimateParagraphHeightPx,
  uniqueSortedPixelBoundaries,
} from "./line-height-table";
import {
  mergeTableSpacing,
  resolveTableSpacingPaddingPx,
} from "./table-utils";
import { resolveTableRenderMetricsPx } from "./table-metrics";

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

  const {
    columnCount,
    tableWidthPx: resolvedTableWidthPx,
    columnWidthsPx: tableColumnWidthsPx,
  } = resolveTableRenderMetricsPx(table, maxAvailableWidthPx);
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
      spannedWidthPx > 0
        ? spannedWidthPx
        : twipsToPixels(cell.style?.widthTwips) ?? fallbackCellWidthPx;
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
  docGridLinePitchPx?: number,
  layoutDefaults?: DocumentLayoutDefaults
): number {
  const sourceXml = table.sourceXml;
  // Table cells render zero-margin paragraphs: only the tab grid applies.
  const tableDefaults = tabGridLayoutDefaults(layoutDefaults);
  const widthKey = heightEstimateCacheKeyPx(
    maxAvailableWidthPx,
    docGridLinePitchPx,
    false,
    tableDefaults
  );
  if (sourceXml) {
    const cachedByWidth = tableEstimatedHeightBySourceXml.get(sourceXml);
    const cached = cachedByWidth?.get(widthKey);
    if (cached !== undefined) return cached;
  }
  const estimatedRowsHeightPx = estimateTableRowHeightsPx(
    table,
    maxAvailableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx,
    undefined,
    tableDefaults
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
      tableEstimatedHeightBySourceXml.get(sourceXml) ??
      new Map<number | string, number>();
    cachedByWidth.set(widthKey, estimatedHeightPx);
    setCacheEntry(tableEstimatedHeightBySourceXml, sourceXml, cachedByWidth);
  }
  return estimatedHeightPx;
}

export function estimateDocNodeHeightPx(
  node: DocModel["nodes"][number],
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number,
  layoutDefaults?: DocumentLayoutDefaults
): number {
  return node.type === "paragraph"
    ? estimateParagraphHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx,
        undefined,
        layoutDefaults
      )
    : estimateTableHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx,
        layoutDefaults
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
