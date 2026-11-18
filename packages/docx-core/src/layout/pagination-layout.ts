// Layout-layer pagination adapter: turns the real page-segmentation engine
// (`buildDocumentPageNodeSegments`) into page geometry that the viewer can
// consume.
//
// Why this module exists
// ----------------------
// `layout-engine.ts:layoutDocument()` is a coarse one-line-per-paragraph
// estimator: it never wraps lines, never honours explicit page breaks /
// section breaks / keepNext / widow control, never splits table rows and uses a
// single page geometry for the whole document. It stays exported for the public
// `core` entry, but it must not be the pagination source of truth.
//
// `page-segmentation-core.ts:buildDocumentPageNodeSegments()` is the real
// engine (line ranges, table row ranges, keep-next chains, hard breaks,
// per-section metrics). It needs a callback bundle plus per-section metrics;
// nothing in the runtime used to assemble them, which is why the engine had no
// caller. This module assembles both and projects the resulting segments into
// page/block frames.
//
// Remaining known gap (see issue #25): the Vue renderer still renders whole
// paragraphs per segment, so paragraph line splitting stays disabled here
// (`allowParagraphLineSplitting: false`) until `DocxParagraphHost` can clip to a
// line range. Table row ranges are already rendered by `DocxTableHost`.

import type {
  DocModel,
  NumberingDefinitionSet,
  ParagraphNode,
  TableNode
} from "../engine/types";
import { estimateParagraphLineHeightPx } from "../editor/helpers/line-height";
import {
  parseSectionColumns,
  resolveFooterPaginationReservePx,
  resolveHeaderPaginationReservePx,
  sectionHasVisibleFooterContent,
  sectionHasVisibleHeaderContent
} from "../editor/helpers/header-footer";
import {
  estimateParagraphHeightPx,
  estimateTableRowHeightsPx
} from "../editor/helpers/line-height-table";
import {
  estimateDocNodeHeightPx,
  paragraphCanSplitAcrossPages,
  paragraphWidowControlEnabled
} from "../editor/helpers/line-height-table-extra";
import { resolvePageNoteReservePx } from "../editor/helpers/note-numbering";
import { paragraphLineCountWithinWidth } from "../editor/helpers/line-height-wrap";
import {
  resolveDocumentLayoutDefaults,
  tabGridLayoutDefaults,
  type DocumentLayoutDefaults,
} from "../engine/layout-defaults";
import { resolveParagraphLineMetrics } from "../editor/helpers/paragraph-line-metrics";
import { paragraphLastRenderedPageBreakLineIndex } from "../editor/helpers/pagination-line-index";
import { paragraphHasVisibleText } from "../editor/helpers/paragraph-geometry";
import { paragraphIsStructuralSectionBreakSpacer } from "../editor/helpers/paragraph-tracked";
import { paragraphHasExplicitColumnBreak } from "../editor/helpers/paragraph-tracked";
import { buildRenderColumnSegmentsForPageSection } from "../editor/helpers/pagination-plan-core";
import { documentPageNodeSegmentsIdentityKey } from "../editor/helpers/pagination-plan-stabilize";
import { parseSectionLayout } from "../viewer/section-layout";
import {
  buildDocumentPageNodeSegments,
  type DocumentPageNodeSegment,
  type PageSegmentationCallbacks,
  type ParagraphLineRange,
  type TableRowRange
} from "./page-segmentation-core";
import {
  paragraphAfterSpacingPx,
  paragraphBeforeSpacingPx,
  resolveDocumentSectionsFromMetadata,
  resolveParagraphConsumedHeightPx,
  resolveSectionIndexForNodeIndex,
  type PaginationSectionMetrics
} from "./pagination";

const MIN_SECTION_CONTENT_PX = 120;
const DEFAULT_MIN_ROW_HEIGHT_PX = 14;

export interface PaginationRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PaginationMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface PaginationSectionGeometry {
  startNodeIndex: number;
  pageWidthPx: number;
  pageHeightPx: number;
  marginsPx: PaginationMargins;
  contentBoxPx: PaginationRect;
  docGridLinePitchPx?: number;
  /** Real section column count (`w:cols w:num`), 1 when the body is single-column. */
  columnCount: number;
  /** Gap between adjacent columns in CSS pixels. */
  columnGapPx: number;
  /** Width of each column in CSS pixels (length === columnCount). */
  columnWidthsPx: number[];
}

export interface PaginatedPageBlock {
  kind: "paragraph" | "table";
  nodeIndex: number;
  framePx: PaginationRect;
  paragraphLineRange?: ParagraphLineRange;
  tableRowRange?: TableRowRange;
  /** Zero-based column within the page body (0 for single-column sections). */
  columnIndex: number;
}

export interface PaginatedPage {
  index: number;
  sectionIndex: number;
  geometry: PaginationSectionGeometry;
  segments: DocumentPageNodeSegment[];
  /** Segments per body column; a single entry for single-column sections. */
  columns: DocumentPageNodeSegment[][];
  blocks: PaginatedPageBlock[];
}

export interface PaginatedDocumentLayout {
  pages: PaginatedPage[];
  segmentsByPage: DocumentPageNodeSegment[][];
  sectionGeometries: PaginationSectionGeometry[];
  sectionMetrics: PaginationSectionMetrics[];
}

/**
 * Feedback from a previous render/measure pass. Paragraph and table-row
 * heights are measured in the DOM; page content heights are only applied when
 * the page identity keys still match the re-estimated pagination, which is the
 * oscillation guard for the measure→repaginate loop.
 */
export interface PaginationMeasurementFeedback {
  paragraphHeightsPxByNodeIndex?: Record<number, number>;
  tableRowHeightsPxByNodeIndex?: Record<number, number[]>;
  pageContentHeightsPxByPageIndex?: number[];
  pageIdentityKeys?: string[];
}

export interface PaginateDocumentLayoutOptions {
  pageWidthPx: number;
  pageHeightPx: number;
  marginPx: number;
  /** Defaults to false until paragraph line ranges can be clipped when rendering. */
  allowParagraphLineSplitting?: boolean;
  /** Defaults to false; mirrored from the document compatibility settings by callers. */
  suppressSpacingBeforeAfterPageBreak?: boolean;
  /**
   * Per-page content height overrides (page index → px). Used by the note-height
   * reserve: a page that prints footnotes/endnotes has less room for body text.
   */
  pageContentHeightOverridesPx?: readonly number[];
  /** Honour Word's cached `w:lastRenderedPageBreak` markers. Defaults to true. */
  useLastRenderedPageBreakHints?: boolean;
  /** Measured-height feedback from the previous render pass (#99). */
  measurementFeedback?: PaginationMeasurementFeedback;
}

function positiveRounded(value: number, fallback: number): number {
  const rounded = Math.round(value);
  return Number.isFinite(rounded) && rounded > 0 ? rounded : fallback;
}

function resolveSectionContentBox(params: {
  pageWidthPx: number;
  pageHeightPx: number;
  marginsPx: PaginationMargins;
  headerReservePx?: number;
  footerReservePx?: number;
}): PaginationRect {
  const headerReservePx = Math.max(0, Math.round(params.headerReservePx ?? 0));
  const footerReservePx = Math.max(0, Math.round(params.footerReservePx ?? 0));
  return {
    x: params.marginsPx.left,
    y: params.marginsPx.top + headerReservePx,
    width: Math.max(
      MIN_SECTION_CONTENT_PX,
      params.pageWidthPx - params.marginsPx.left - params.marginsPx.right
    ),
    height: Math.max(
      MIN_SECTION_CONTENT_PX,
      params.pageHeightPx -
        params.marginsPx.top -
        params.marginsPx.bottom -
        headerReservePx -
        footerReservePx
    )
  };
}

/**
 * Per-section page geometry. Sections without explicit `<w:sectPr>` page setup
 * inherit the caller-supplied document layout options.
 */
export function resolvePaginationSectionGeometry(
  model: DocModel,
  options: PaginateDocumentLayoutOptions
): PaginationSectionGeometry[] {
  const defaultMargins = Math.max(0, Math.round(options.marginPx));
  const defaultGeometry: Omit<PaginationSectionGeometry, "startNodeIndex"> = {
    pageWidthPx: positiveRounded(options.pageWidthPx, 816),
    pageHeightPx: positiveRounded(options.pageHeightPx, 1056),
    marginsPx: {
      top: defaultMargins,
      right: defaultMargins,
      bottom: defaultMargins,
      left: defaultMargins
    },
    contentBoxPx: { x: 0, y: 0, width: 0, height: 0 },
    columnCount: 1,
    columnGapPx: 0,
    columnWidthsPx: [0]
  };
  defaultGeometry.contentBoxPx = resolveSectionContentBox(defaultGeometry);
  defaultGeometry.columnWidthsPx = [defaultGeometry.contentBoxPx.width];

  const resolvedSections = resolveDocumentSectionsFromMetadata(
    model.metadata
  );
  const layoutDefaults = resolveDocumentLayoutDefaults(model.metadata);

  return resolvedSections.map((section) => {
    const startNodeIndex =
      Number.isFinite(section.startNodeIndex) && section.startNodeIndex >= 0
        ? Math.round(section.startNodeIndex)
        : 0;
    const hasSectionProperties = Boolean(section.sectionPropertiesXml);
    const parsedLayout = parseSectionLayout(section.sectionPropertiesXml);
    const marginsPx: PaginationMargins = hasSectionProperties
      ? {
          top: Math.max(0, Math.round(parsedLayout.marginsPx.top)),
          right: Math.max(0, Math.round(parsedLayout.marginsPx.right)),
          bottom: Math.max(0, Math.round(parsedLayout.marginsPx.bottom)),
          left: Math.max(0, Math.round(parsedLayout.marginsPx.left))
        }
      : defaultGeometry.marginsPx;
    const geometry = {
      pageWidthPx: hasSectionProperties
        ? positiveRounded(parsedLayout.pageWidthPx, defaultGeometry.pageWidthPx)
        : defaultGeometry.pageWidthPx,
      pageHeightPx: hasSectionProperties
        ? positiveRounded(parsedLayout.pageHeightPx, defaultGeometry.pageHeightPx)
        : defaultGeometry.pageHeightPx,
      marginsPx,
      docGridLinePitchPx: parsedLayout.docGridLinePitchPx
    };
    const reserveOptions = {
      estimateDocNodeHeightPx: (
        node: DocModel["nodes"][number],
        availableWidthPx?: number,
        definitions?: NumberingDefinitionSet,
        docGridLinePitchPx?: number
      ) =>
        estimateDocNodeHeightPx(
          node,
          availableWidthPx,
          definitions,
          docGridLinePitchPx,
          layoutDefaults
        ),
      // Pagination should reserve only the real overlap with the body, not the
      // renderer's small safety strip for a footer that fits in the margin.
      minimumVisibleFlowFooterReservePx: 0
    };
    const headerReservePx = sectionHasVisibleHeaderContent(section)
      ? resolveHeaderPaginationReservePx(
          section.headerSections,
          {
            ...geometry,
            headerDistancePx: parsedLayout.headerDistancePx
          },
          reserveOptions
        )
      : 0;
    const footerReservePx = sectionHasVisibleFooterContent(section)
      ? resolveFooterPaginationReservePx(
          section.footerSections,
          {
            ...geometry,
            footerDistancePx: parsedLayout.footerDistancePx
          },
          reserveOptions
        )
      : 0;

    const contentBoxPx = resolveSectionContentBox({
      ...geometry,
      headerReservePx,
      footerReservePx
    });
    const parsedColumns = parseSectionColumns(section.sectionPropertiesXml);
    const columnCount = Math.max(1, parsedColumns?.count ?? 1);
    const columnGapPx = columnCount > 1 ? Math.max(0, parsedColumns?.gapPx ?? 0) : 0;
    const equalColumnWidthPx = Math.max(
      24,
      Math.round(
        (contentBoxPx.width - columnGapPx * (columnCount - 1)) / columnCount
      )
    );
    const columnWidthsPx =
      columnCount === 1
        ? [contentBoxPx.width]
        : parsedColumns?.widthsPx?.length === columnCount
        ? [...(parsedColumns.widthsPx ?? [])]
        : Array.from({ length: columnCount }, () => equalColumnWidthPx);

    return {
      startNodeIndex,
      ...geometry,
      contentBoxPx,
      columnCount,
      columnGapPx,
      columnWidthsPx
    };
  });
}

function createSegmentationCallbacks(
  model: DocModel,
  measuredParagraphHeightsPxByNodeIndex?: Record<number, number>
): PageSegmentationCallbacks {
  const layoutDefaults = resolveDocumentLayoutDefaults(model.metadata);
  const nodeIndexByParagraph = new WeakMap<ParagraphNode, number>();
  if (measuredParagraphHeightsPxByNodeIndex) {
    model.nodes.forEach((node, nodeIndex) => {
      if (node.type === "paragraph") {
        nodeIndexByParagraph.set(node, nodeIndex);
      }
    });
  }

  const estimateParagraphHeightWithMeasurement = (
    paragraph: ParagraphNode,
    availableWidthPx?: number,
    numberingDefinitions?: NumberingDefinitionSet,
    docGridLinePitchPx?: number
  ): number => {
    if (measuredParagraphHeightsPxByNodeIndex) {
      const nodeIndex = nodeIndexByParagraph.get(paragraph);
      const measuredHeightPx =
        nodeIndex === undefined
          ? undefined
          : measuredParagraphHeightsPxByNodeIndex[nodeIndex];
      if (Number.isFinite(measuredHeightPx) && (measuredHeightPx as number) > 0) {
        return Math.round(measuredHeightPx as number);
      }
    }
    return estimateParagraphHeightPx(
      paragraph,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx,
      undefined,
      layoutDefaults
    );
  };

  return {
    estimateDocNodeHeightPx: (
      node: DocModel["nodes"][number],
      availableWidthPx?: number,
      definitions?: NumberingDefinitionSet,
      docGridLinePitchPx?: number
    ) =>
      estimateDocNodeHeightPx(
        node,
        availableWidthPx,
        definitions,
        docGridLinePitchPx,
        layoutDefaults
      ),
    paragraphHasVisibleText,
    estimateParagraphHeightPx: estimateParagraphHeightWithMeasurement,
    estimateParagraphLineHeightPx,
    estimateTableRowHeightsPx: (
      table: TableNode,
      availableWidthPx?: number,
      definitions?: NumberingDefinitionSet,
      docGridLinePitchPx?: number
    ) =>
      estimateTableRowHeightsPx(
        table,
        availableWidthPx,
        definitions,
        docGridLinePitchPx,
        undefined,
        layoutDefaults
      ),
    paragraphLineCountWithinWidth: (
      paragraph: ParagraphNode,
      availableWidthPx?: number,
      definitions?: NumberingDefinitionSet
    ) =>
      paragraphLineCountWithinWidth(
        paragraph,
        availableWidthPx,
        definitions,
        layoutDefaults
      ),
    paragraphWidowControlEnabled,
    paragraphCanSplitAcrossPages,
    paragraphIsStructuralSectionBreakSpacer,
    resolveParagraphLineMetrics: (params) =>
      resolveParagraphLineMetrics(params.paragraph, {
        availableWidthPx: params.availableWidthPx,
        numberingDefinitions: params.numberingDefinitions,
        lineHeightPx: params.lineHeightPx,
        docGridLinePitchPx: params.docGridLinePitchPx,
        expectedLineCount: params.expectedLineCount,
        layoutDefaults
      })
  };
}

function sectionMetricsFromGeometry(
  geometries: PaginationSectionGeometry[]
): PaginationSectionMetrics[] {
  return geometries.map((geometry) => {
    const columnCount = Math.max(1, geometry.columnCount);
    const columnWidthPx = Math.max(
      MIN_SECTION_CONTENT_PX,
      Math.round(
        Math.min(
          ...geometry.columnWidthsPx,
          geometry.contentBoxPx.width
        )
      )
    );
    return {
      startNodeIndex: geometry.startNodeIndex,
      // Text wraps inside one column; the flow budget spans every column on
      // the page (word processors fill column 1, then column 2, ...).
      pageContentWidthPx: columnWidthPx,
      pageContentHeightPx: Math.max(
        MIN_SECTION_CONTENT_PX,
        Math.round(geometry.contentBoxPx.height) * columnCount
      ),
      pageContentHeightMultiplier: columnCount,
      docGridLinePitchPx: geometry.docGridLinePitchPx
    };
  });
}

function paragraphSegmentHeightPx(params: {
  segment: DocumentPageNodeSegment;
  paragraph: ParagraphNode;
  metrics: PaginationSectionMetrics;
  numberingDefinitions?: NumberingDefinitionSet;
  pageConsumedHeightPx: number;
  availableWidthPx?: number;
  layoutDefaults?: DocumentLayoutDefaults;
}): { heightPx: number; afterSpacingPx: number } {
  const beforeSpacingPx = paragraphBeforeSpacingPx(
    params.paragraph,
    params.layoutDefaults
  );
  const afterSpacingPx = paragraphAfterSpacingPx(
    params.paragraph,
    params.layoutDefaults
  );
  const lineRange = params.segment.paragraphLineRange;

  if (lineRange) {
    const lineHeightPx = Math.max(
      1,
      Math.round(
        lineRange.lineHeightPx ||
          estimateParagraphLineHeightPx(
            params.paragraph,
            params.metrics.docGridLinePitchPx
          )
      )
    );
    const lineCount = Math.max(
      1,
      lineRange.endLineIndex - lineRange.startLineIndex
    );
    // OOXML keeps both spacing values; only a page-top paragraph drops before.
    const topSpacingPx =
      lineRange.startLineIndex === 0 && params.pageConsumedHeightPx > 0
        ? beforeSpacingPx
        : 0;
    const bottomSpacingPx =
      lineRange.endLineIndex >= lineRange.totalLineCount ? afterSpacingPx : 0;

    return {
      heightPx: topSpacingPx + lineCount * lineHeightPx + bottomSpacingPx,
      afterSpacingPx: bottomSpacingPx > 0 ? afterSpacingPx : 0
    };
  }

  const rawHeightPx = Math.max(
    1,
    Math.round(
      estimateParagraphHeightPx(
        params.paragraph,
        params.availableWidthPx ?? params.metrics.pageContentWidthPx,
        params.numberingDefinitions,
        params.metrics.docGridLinePitchPx,
        undefined,
        params.layoutDefaults
      )
    )
  );

  return {
    heightPx: resolveParagraphConsumedHeightPx(
      rawHeightPx,
      beforeSpacingPx,
      params.pageConsumedHeightPx <= 0
    ),
    afterSpacingPx
  };
}

function tableSegmentHeightPx(params: {
  table: TableNode;
  segment: DocumentPageNodeSegment;
  metrics: PaginationSectionMetrics;
  numberingDefinitions?: NumberingDefinitionSet;
  availableWidthPx?: number;
  layoutDefaults?: DocumentLayoutDefaults;
}): number {
  const rowHeightsPx = estimateTableRowHeightsPx(
    params.table,
    params.availableWidthPx ?? params.metrics.pageContentWidthPx,
    params.numberingDefinitions,
    params.metrics.docGridLinePitchPx
  );
  if (rowHeightsPx.length === 0) {
    return DEFAULT_MIN_ROW_HEIGHT_PX;
  }

  const startRowIndex = Math.max(
    0,
    Math.min(
      rowHeightsPx.length,
      Math.round(params.segment.tableRowRange?.startRowIndex ?? 0)
    )
  );
  const endRowIndex = Math.max(
    startRowIndex,
    Math.min(
      rowHeightsPx.length,
      Math.round(params.segment.tableRowRange?.endRowIndex ?? rowHeightsPx.length)
    )
  );

  let heightPx = 0;
  const repeatHeaderRowCount = Math.max(
    0,
    Math.round(params.segment.tableRowRange?.repeatHeaderRowCount ?? 0)
  );
  for (let rowIndex = 0; rowIndex < repeatHeaderRowCount; rowIndex += 1) {
    heightPx += Math.max(
      DEFAULT_MIN_ROW_HEIGHT_PX,
      Math.round(rowHeightsPx[rowIndex] ?? DEFAULT_MIN_ROW_HEIGHT_PX)
    );
  }
  const rowSlice = params.segment.tableRowSlice;
  if (rowSlice) {
    heightPx += Math.round(rowSlice.sliceHeightPx);
  } else {
    for (let rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex += 1) {
      heightPx += Math.max(
        DEFAULT_MIN_ROW_HEIGHT_PX,
        Math.round(rowHeightsPx[rowIndex] ?? DEFAULT_MIN_ROW_HEIGHT_PX)
      );
    }
  }

  return Math.max(DEFAULT_MIN_ROW_HEIGHT_PX, heightPx);
}

function projectPageBlocks(params: {
  columns: DocumentPageNodeSegment[][];
  model: DocModel;
  geometry: PaginationSectionGeometry;
  metrics: PaginationSectionMetrics;
  measuredParagraphHeightsPxByNodeIndex?: Map<number, number>;
}): PaginatedPageBlock[] {
  const blocks: PaginatedPageBlock[] = [];
  const contentBox = params.geometry.contentBoxPx;
  const numberingDefinitions = params.model.metadata.numberingDefinitions;
  const layoutDefaults = resolveDocumentLayoutDefaults(params.model.metadata);
  const columnWidthsPx =
    params.geometry.columnWidthsPx.length > 0
      ? params.geometry.columnWidthsPx
      : [contentBox.width];
  let columnX = contentBox.x;

  params.columns.forEach((columnSegments, columnIndex) => {
    const columnWidthPx = Math.max(
      24,
      Math.round(columnWidthsPx[columnIndex] ?? columnWidthsPx[0] ?? contentBox.width)
    );
    let pageConsumedHeightPx = 0;

    for (const segment of columnSegments) {
      const node = params.model.nodes[segment.nodeIndex];
      if (!node) {
        continue;
      }

      if (node.type === "paragraph") {
        const measuredHeightPx = params.measuredParagraphHeightsPxByNodeIndex?.get(
          segment.nodeIndex
        );
        const height =
          Number.isFinite(measuredHeightPx) && (measuredHeightPx as number) > 0
            ? {
                heightPx: Math.round(measuredHeightPx as number),
                afterSpacingPx: paragraphAfterSpacingPx(node, layoutDefaults)
              }
            : paragraphSegmentHeightPx({
                segment,
                paragraph: node,
                metrics: params.metrics,
                numberingDefinitions,
                pageConsumedHeightPx,
                availableWidthPx: columnWidthPx,
                layoutDefaults
              });
        blocks.push({
          kind: "paragraph",
          nodeIndex: segment.nodeIndex,
          framePx: {
            x: columnX,
            y: contentBox.y + pageConsumedHeightPx,
            width: columnWidthPx,
            height: height.heightPx
          },
          paragraphLineRange: segment.paragraphLineRange,
          columnIndex
        });
        pageConsumedHeightPx += height.heightPx;
        continue;
      }

      const heightPx = tableSegmentHeightPx({
        table: node,
        segment,
        metrics: params.metrics,
        numberingDefinitions,
        availableWidthPx: columnWidthPx,
        layoutDefaults
      });
      blocks.push({
        kind: "table",
        nodeIndex: segment.nodeIndex,
        framePx: {
          x: columnX,
          y: contentBox.y + pageConsumedHeightPx,
          width: columnWidthPx,
          height: heightPx
        },
        tableRowRange: segment.tableRowRange,
        columnIndex
      });
      pageConsumedHeightPx += heightPx;
    }

    columnX += columnWidthPx + params.geometry.columnGapPx;
  });

  return blocks;
}

function resolveKeepNextGroupEnds(model: DocModel): Map<number, number> {
  const groupEndByStartNodeIndex = new Map<number, number>();
  for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
    const node = model.nodes[nodeIndex];
    if (
      node.type !== "paragraph" ||
      node.style?.keepNext !== true ||
      !paragraphHasVisibleText(node)
    ) {
      continue;
    }
    let chainEndNodeIndex = nodeIndex;
    while (chainEndNodeIndex < model.nodes.length - 1) {
      const current = model.nodes[chainEndNodeIndex];
      if (
        current.type !== "paragraph" ||
        current.style?.keepNext !== true ||
        !paragraphHasVisibleText(current)
      ) {
        break;
      }
      const next = model.nodes[chainEndNodeIndex + 1];
      if (next.type !== "paragraph") {
        chainEndNodeIndex += 1;
        break;
      }
      chainEndNodeIndex += 1;
    }
    if (chainEndNodeIndex > nodeIndex) {
      groupEndByStartNodeIndex.set(nodeIndex, chainEndNodeIndex);
    }
  }
  return groupEndByStartNodeIndex;
}

/**
 * Paginate a document with the real segmentation engine and project the result
 * into per-page blocks with absolute frames.
 */
export function paginateDocumentLayout(
  model: DocModel,
  options: PaginateDocumentLayoutOptions
): PaginatedDocumentLayout {
  const sectionGeometries = resolvePaginationSectionGeometry(model, options);
  const sectionMetrics = sectionMetricsFromGeometry(sectionGeometries);
  const fallbackMetrics = sectionMetrics[0];
  const feedback = options.measurementFeedback;

  const buildSegments = (measuredPageContentHeightsPx?: number[]) =>
    model.nodes.length === 0
      ? []
      : buildDocumentPageNodeSegments(
          model,
          fallbackMetrics.pageContentHeightPx,
          fallbackMetrics.pageContentWidthPx,
          createSegmentationCallbacks(
            model,
            feedback?.paragraphHeightsPxByNodeIndex
          ),
          model.metadata.numberingDefinitions,
          sectionMetrics,
          {
            allowParagraphLineSplitting: options.allowParagraphLineSplitting ?? false,
            suppressSpacingBeforeAfterPageBreak:
              options.suppressSpacingBeforeAfterPageBreak ?? false,
            preferLastRenderedParagraphStartBreaks:
              options.useLastRenderedPageBreakHints ?? true,
            paragraphLastRenderedPageBreakLineIndex: (
              paragraph,
              lineCount,
              widthPx,
              definitions
            ) =>
              paragraphLastRenderedPageBreakLineIndex(
                paragraph,
                lineCount,
                widthPx,
                definitions,
                resolveDocumentLayoutDefaults(model.metadata)
              ),
            measuredPageContentHeightsPxByPageIndex:
              options.pageContentHeightOverridesPx
                ? [...options.pageContentHeightOverridesPx]
                : measuredPageContentHeightsPx,
            measuredTableRowHeightsByNodeIndex:
              feedback?.tableRowHeightsPxByNodeIndex
          }
        );

  let segmentsByPage = buildSegments();
  if (
    feedback?.pageContentHeightsPxByPageIndex?.some((heightPx) =>
      Number.isFinite(heightPx) && heightPx > 0
    )
  ) {
    // Apply measured page budgets only where the page still contains the same
    // segments that were measured; otherwise the value belongs to a different
    // pagination and would oscillate.
    const currentIdentityKeys = segmentsByPage.map(
      documentPageNodeSegmentsIdentityKey
    );
    const measuredHeights = feedback.pageContentHeightsPxByPageIndex.map(
      (heightPx, pageIndex) =>
        feedback.pageIdentityKeys?.[pageIndex] === currentIdentityKeys[pageIndex]
          ? heightPx
          : -1
    );
    if (measuredHeights.some((heightPx) => heightPx > 0)) {
      segmentsByPage = buildSegments(measuredHeights);
    }
  }

  const pagesBySegment: DocumentPageNodeSegment[][] =
    segmentsByPage.length > 0 ? segmentsByPage : [[]];

  let sectionIndex = 0;
  const pageSections = pagesBySegment.map((pageSegments) => {
    const firstNodeIndex = pageSegments[0]?.nodeIndex;
    sectionIndex =
      firstNodeIndex === undefined
        ? sectionIndex
        : resolveSectionIndexForNodeIndex(sectionGeometries, firstNodeIndex, sectionIndex);
    return sectionIndex;
  });
  const forceColumnBreakNodeIndexes = new Set<number>();
  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph" && paragraphHasExplicitColumnBreak(node)) {
      forceColumnBreakNodeIndexes.add(nodeIndex);
    }
  });
  const keepTogetherNodeRangeByStartNodeIndex = resolveKeepNextGroupEnds(model);
  const measuredParagraphHeightsPxByNodeIndex =
    feedback?.paragraphHeightsPxByNodeIndex
      ? new Map(
          Object.entries(feedback.paragraphHeightsPxByNodeIndex).map(
            ([nodeIndex, heightPx]) => [Number(nodeIndex), heightPx] as const
          )
        )
      : undefined;

  const pages: PaginatedPage[] = pagesBySegment.map((pageSegments, pageIndex) => {
    const section = pageSections[pageIndex] ?? 0;
    const geometry = sectionGeometries[section] ?? sectionGeometries[0];
    const metrics = sectionMetrics[section] ?? fallbackMetrics;
    const nextSection = pageSections[pageIndex + 1];
    const isLastPageOfSection =
      nextSection === undefined || nextSection !== section;
    const columns =
      geometry.columnCount > 1 && pageSegments.length > 0
        ? buildRenderColumnSegmentsForPageSection(
            model,
            pageSegments,
            geometry.columnWidthsPx,
            geometry.contentBoxPx.height,
            model.metadata.numberingDefinitions,
            undefined,
            measuredParagraphHeightsPxByNodeIndex,
            isLastPageOfSection,
            forceColumnBreakNodeIndexes,
            keepTogetherNodeRangeByStartNodeIndex
          )
        : [pageSegments];
    const flowSegments = columns.flat();

    return {
      index: pageIndex,
      sectionIndex: section,
      geometry,
      segments: flowSegments,
      columns,
      blocks: projectPageBlocks({
        columns,
        model,
        geometry,
        metrics,
        measuredParagraphHeightsPxByNodeIndex
      })
    };
  });

  return {
    pages,
    segmentsByPage: pages.map((page) => page.segments),
    sectionGeometries,
    sectionMetrics
  };
}

const MAX_NOTE_RESERVE_ITERATIONS = 3;

function samePageSegments(
  left: DocumentPageNodeSegment[][],
  right: DocumentPageNodeSegment[][]
): boolean {
  if (left.length !== right.length) {
    return false;
  }
  return left.every((segments, pageIndex) => {
    const other = right[pageIndex] ?? [];
    if (segments.length !== other.length) {
      return false;
    }
    return segments.every((segment, segmentIndex) => {
      const candidate = other[segmentIndex];
      return (
        candidate?.nodeIndex === segment.nodeIndex &&
        candidate?.tableRowRange?.startRowIndex ===
          segment.tableRowRange?.startRowIndex &&
        candidate?.tableRowRange?.endRowIndex ===
          segment.tableRowRange?.endRowIndex &&
        candidate?.paragraphLineRange?.startLineIndex ===
          segment.paragraphLineRange?.startLineIndex &&
        candidate?.paragraphLineRange?.endLineIndex ===
          segment.paragraphLineRange?.endLineIndex
      );
    });
  });
}

/**
 * Paginate while reserving the height of the notes printed at the bottom of
 * each page.
 *
 * A page that owns footnotes (or endnotes placed at document/section end) has
 * less room for body content. Pagination is re-run with per-page content-height
 * overrides until the page segmentation is stable, so a footnote carrying a tall
 * table can no longer overflow the last page.
 */
export function paginateDocumentLayoutWithNoteReserve(
  model: DocModel,
  options: PaginateDocumentLayoutOptions
): PaginatedDocumentLayout {
  const baseOptions: PaginateDocumentLayoutOptions = {
    ...options,
    pageContentHeightOverridesPx: undefined
  };
  let paginated = paginateDocumentLayout(model, baseOptions);
  if (paginated.pages.length === 0) {
    return paginated;
  }

  for (let iteration = 0; iteration < MAX_NOTE_RESERVE_ITERATIONS; iteration += 1) {
    const reservePx = resolvePageNoteReservePx(model, paginated.segmentsByPage, {
      estimateDocNodeHeightPx: (
        node: DocModel["nodes"][number],
        availableWidthPx?: number,
        definitions?: NumberingDefinitionSet,
        docGridLinePitchPx?: number
      ) =>
        estimateDocNodeHeightPx(
          node,
          availableWidthPx,
          definitions,
          docGridLinePitchPx,
          tabGridLayoutDefaults(resolveDocumentLayoutDefaults(model.metadata))
        ),
      numberingDefinitions: model.metadata.numberingDefinitions,
      resolvePageContentWidthPx: (pageIndex) =>
        Math.max(
          MIN_SECTION_CONTENT_PX,
          Math.round(
            paginated.pages[pageIndex]?.geometry.contentBoxPx.width ??
              MIN_SECTION_CONTENT_PX
          )
        )
    });
    if (!reservePx.some((value) => value > 0)) {
      break;
    }

    const overrides = paginated.pages.map((page, pageIndex) =>
      Math.max(
        MIN_SECTION_CONTENT_PX,
        Math.round(page.geometry.contentBoxPx.height) -
          Math.round(reservePx[pageIndex] ?? 0)
      )
    );
    const next = paginateDocumentLayout(model, {
      ...baseOptions,
      pageContentHeightOverridesPx: overrides
    });
    const stable = samePageSegments(paginated.segmentsByPage, next.segmentsByPage);
    paginated = next;
    if (stable) {
      break;
    }
  }

  return paginated;
}
