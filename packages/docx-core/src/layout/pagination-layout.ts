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
  estimateParagraphHeightPx,
  estimateTableRowHeightsPx
} from "../editor/helpers/line-height-table";
import {
  estimateDocNodeHeightPx,
  paragraphCanSplitAcrossPages,
  paragraphWidowControlEnabled
} from "../editor/helpers/line-height-table-extra";
import { paragraphLineCountWithinWidth } from "../editor/helpers/line-height-wrap";
import { paragraphHasVisibleText } from "../editor/helpers/paragraph-geometry";
import { paragraphIsStructuralSectionBreakSpacer } from "../editor/helpers/paragraph-tracked";
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
}

export interface PaginatedPageBlock {
  kind: "paragraph" | "table";
  nodeIndex: number;
  framePx: PaginationRect;
  paragraphLineRange?: ParagraphLineRange;
  tableRowRange?: TableRowRange;
}

export interface PaginatedPage {
  index: number;
  sectionIndex: number;
  geometry: PaginationSectionGeometry;
  segments: DocumentPageNodeSegment[];
  blocks: PaginatedPageBlock[];
}

export interface PaginatedDocumentLayout {
  pages: PaginatedPage[];
  segmentsByPage: DocumentPageNodeSegment[][];
  sectionGeometries: PaginationSectionGeometry[];
  sectionMetrics: PaginationSectionMetrics[];
}

export interface PaginateDocumentLayoutOptions {
  pageWidthPx: number;
  pageHeightPx: number;
  marginPx: number;
  /** Defaults to false until paragraph line ranges can be clipped when rendering. */
  allowParagraphLineSplitting?: boolean;
  /** Defaults to false; mirrored from the document compatibility settings by callers. */
  suppressSpacingBeforeAfterPageBreak?: boolean;
}

function positiveRounded(value: number, fallback: number): number {
  const rounded = Math.round(value);
  return Number.isFinite(rounded) && rounded > 0 ? rounded : fallback;
}

function resolveSectionContentBox(params: {
  pageWidthPx: number;
  pageHeightPx: number;
  marginsPx: PaginationMargins;
}): PaginationRect {
  return {
    x: params.marginsPx.left,
    y: params.marginsPx.top,
    width: Math.max(
      MIN_SECTION_CONTENT_PX,
      params.pageWidthPx - params.marginsPx.left - params.marginsPx.right
    ),
    height: Math.max(
      MIN_SECTION_CONTENT_PX,
      params.pageHeightPx - params.marginsPx.top - params.marginsPx.bottom
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
    contentBoxPx: { x: 0, y: 0, width: 0, height: 0 }
  };
  defaultGeometry.contentBoxPx = resolveSectionContentBox(defaultGeometry);

  const sections = model.metadata.sections ?? [];
  if (sections.length === 0) {
    return [{ startNodeIndex: 0, ...defaultGeometry }];
  }

  return sections.map((section, index) => {
    const startNodeIndex =
      Number.isFinite(section.startNodeIndex) && section.startNodeIndex >= 0
        ? Math.round(section.startNodeIndex)
        : 0;

    if (index === 0 && !section.sectionPropertiesXml) {
      return { startNodeIndex, ...defaultGeometry };
    }

    if (!section.sectionPropertiesXml) {
      return { startNodeIndex, ...defaultGeometry };
    }

    const layout = parseSectionLayout(section.sectionPropertiesXml);
    const marginsPx: PaginationMargins = {
      top: Math.max(0, Math.round(layout.marginsPx.top)),
      right: Math.max(0, Math.round(layout.marginsPx.right)),
      bottom: Math.max(0, Math.round(layout.marginsPx.bottom)),
      left: Math.max(0, Math.round(layout.marginsPx.left))
    };
    const geometry = {
      pageWidthPx: positiveRounded(layout.pageWidthPx, defaultGeometry.pageWidthPx),
      pageHeightPx: positiveRounded(layout.pageHeightPx, defaultGeometry.pageHeightPx),
      marginsPx,
      docGridLinePitchPx: layout.docGridLinePitchPx
    };

    return {
      startNodeIndex,
      ...geometry,
      contentBoxPx: resolveSectionContentBox(geometry)
    };
  });
}

function createSegmentationCallbacks(): PageSegmentationCallbacks {
  return {
    estimateDocNodeHeightPx,
    paragraphHasVisibleText,
    estimateParagraphHeightPx,
    estimateParagraphLineHeightPx,
    estimateTableRowHeightsPx,
    paragraphLineCountWithinWidth,
    paragraphWidowControlEnabled,
    paragraphCanSplitAcrossPages,
    paragraphIsStructuralSectionBreakSpacer
  };
}

function sectionMetricsFromGeometry(
  geometries: PaginationSectionGeometry[]
): PaginationSectionMetrics[] {
  return geometries.map((geometry) => ({
    startNodeIndex: geometry.startNodeIndex,
    pageContentWidthPx: Math.max(
      MIN_SECTION_CONTENT_PX,
      Math.round(geometry.contentBoxPx.width)
    ),
    pageContentHeightPx: Math.max(
      MIN_SECTION_CONTENT_PX,
      Math.round(geometry.contentBoxPx.height)
    ),
    pageContentHeightMultiplier: 1,
    docGridLinePitchPx: geometry.docGridLinePitchPx
  }));
}

function paragraphSegmentHeightPx(params: {
  segment: DocumentPageNodeSegment;
  paragraph: ParagraphNode;
  metrics: PaginationSectionMetrics;
  numberingDefinitions?: NumberingDefinitionSet;
  pageConsumedHeightPx: number;
  previousParagraphAfterPx: number;
}): { heightPx: number; afterSpacingPx: number } {
  const beforeSpacingPx = paragraphBeforeSpacingPx(params.paragraph);
  const afterSpacingPx = paragraphAfterSpacingPx(params.paragraph);
  const collapsedMarginPx =
    params.pageConsumedHeightPx > 0
      ? Math.min(params.previousParagraphAfterPx, beforeSpacingPx)
      : 0;
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
    const topSpacingPx =
      lineRange.startLineIndex === 0
        ? Math.max(0, beforeSpacingPx - collapsedMarginPx)
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
        params.metrics.pageContentWidthPx,
        params.numberingDefinitions,
        params.metrics.docGridLinePitchPx
      )
    )
  );

  return {
    heightPx: Math.max(1, rawHeightPx - collapsedMarginPx),
    afterSpacingPx
  };
}

function tableSegmentHeightPx(params: {
  table: TableNode;
  segment: DocumentPageNodeSegment;
  metrics: PaginationSectionMetrics;
  numberingDefinitions?: NumberingDefinitionSet;
}): number {
  const rowHeightsPx = estimateTableRowHeightsPx(
    params.table,
    params.metrics.pageContentWidthPx,
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
  for (let rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex += 1) {
    heightPx += Math.max(
      DEFAULT_MIN_ROW_HEIGHT_PX,
      Math.round(rowHeightsPx[rowIndex] ?? DEFAULT_MIN_ROW_HEIGHT_PX)
    );
  }

  return Math.max(DEFAULT_MIN_ROW_HEIGHT_PX, heightPx);
}

function projectPageBlocks(params: {
  pageSegments: DocumentPageNodeSegment[];
  model: DocModel;
  geometry: PaginationSectionGeometry;
  metrics: PaginationSectionMetrics;
}): PaginatedPageBlock[] {
  const blocks: PaginatedPageBlock[] = [];
  const contentBox = params.geometry.contentBoxPx;
  const numberingDefinitions = params.model.metadata.numberingDefinitions;
  let pageConsumedHeightPx = 0;
  let previousParagraphAfterPx = 0;

  for (const segment of params.pageSegments) {
    const node = params.model.nodes[segment.nodeIndex];
    if (!node) {
      continue;
    }

    if (node.type === "paragraph") {
      const height = paragraphSegmentHeightPx({
        segment,
        paragraph: node,
        metrics: params.metrics,
        numberingDefinitions,
        pageConsumedHeightPx,
        previousParagraphAfterPx
      });
      blocks.push({
        kind: "paragraph",
        nodeIndex: segment.nodeIndex,
        framePx: {
          x: contentBox.x,
          y: contentBox.y + pageConsumedHeightPx,
          width: contentBox.width,
          height: height.heightPx
        },
        paragraphLineRange: segment.paragraphLineRange
      });
      pageConsumedHeightPx += height.heightPx;
      previousParagraphAfterPx = height.afterSpacingPx;
      continue;
    }

    const heightPx = tableSegmentHeightPx({
      table: node,
      segment,
      metrics: params.metrics,
      numberingDefinitions
    });
    blocks.push({
      kind: "table",
      nodeIndex: segment.nodeIndex,
      framePx: {
        x: contentBox.x,
        y: contentBox.y + pageConsumedHeightPx,
        width: contentBox.width,
        height: heightPx
      },
      tableRowRange: segment.tableRowRange
    });
    pageConsumedHeightPx += heightPx;
    previousParagraphAfterPx = 0;
  }

  return blocks;
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

  const segmentsByPage =
    model.nodes.length === 0
      ? []
      : buildDocumentPageNodeSegments(
          model,
          fallbackMetrics.pageContentHeightPx,
          fallbackMetrics.pageContentWidthPx,
          createSegmentationCallbacks(),
          model.metadata.numberingDefinitions,
          sectionMetrics,
          {
            allowParagraphLineSplitting: options.allowParagraphLineSplitting ?? false,
            suppressSpacingBeforeAfterPageBreak:
              options.suppressSpacingBeforeAfterPageBreak ?? false
          }
        );

  const pagesBySegment: DocumentPageNodeSegment[][] =
    segmentsByPage.length > 0 ? segmentsByPage : [[]];

  let sectionIndex = 0;
  const pages: PaginatedPage[] = pagesBySegment.map((pageSegments, pageIndex) => {
    const firstNodeIndex = pageSegments[0]?.nodeIndex;
    sectionIndex =
      firstNodeIndex === undefined
        ? sectionIndex
        : resolveSectionIndexForNodeIndex(sectionGeometries, firstNodeIndex, sectionIndex);
    const geometry = sectionGeometries[sectionIndex] ?? sectionGeometries[0];
    const metrics = sectionMetrics[sectionIndex] ?? fallbackMetrics;

    return {
      index: pageIndex,
      sectionIndex,
      geometry,
      segments: pageSegments,
      blocks: projectPageBlocks({
        pageSegments,
        model,
        geometry,
        metrics
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
