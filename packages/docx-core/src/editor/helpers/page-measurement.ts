// Page-content height measurement, pagination section metrics, page-segment
// identity keys, and estimated footnote-reserve application.
// Upstream editor.tsx: lines 2433-3053.
//
// Provides the measurement-driven page-content-height entry points consumed by
// the layout layer and the viewer pagination pipeline. `estimateParagraphHeightPx`
// is supplied as a callback so this module stays decoupled from the table-height
// / line-height estimation cluster (mirrors the layout layer's callback-driven
// page-segmentation contract).
//
// `resolvePaginationSectionMetricsIndexForNodeIndex` and
// `scalePaginationSectionMetricsHeights` are re-exported from the layout layer
// (layout/pagination.ts) to avoid duplicating the canonical implementations.

import type {
  DocModel,
  DocumentNoteDefinition,
  NumberingDefinitionSet,
  ParagraphNode
} from "../../engine/types";
import type { DocumentLayoutMetrics } from "../../viewer/section-layout";
import { parseSectionLayout } from "../../viewer/section-layout";
import {
  type PaginationSectionMetrics,
  resolvePaginationSectionMetricsIndexForNodeIndex,
  scalePaginationSectionMetricsHeights
} from "../../layout/pagination";
import type { DocumentPageNodeSegment } from "../../layout/page-segmentation-core";
import {
  type HeaderFooterDocumentSection,
  type HeaderFooterReserveOptions,
  parseSectionColumns,
  resolveSectionPaginationContentWidthPx,
  resolveHeaderPaginationReservePx,
  resolveFooterPaginationReservePx,
  sectionHasVisibleHeaderContent,
  sectionHasVisibleFooterContent
} from "./header-footer";
import { nodeReferencedNoteIds } from "./paragraph-tracked";
import {
  MEASURED_PAGE_FOOTER_CLEARANCE_BUFFER_PX,
  MEASURED_PAGE_HEADER_CLEARANCE_BUFFER_PX,
  UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX
} from "./constants";

export {
  resolvePaginationSectionMetricsIndexForNodeIndex,
  scalePaginationSectionMetricsHeights
};

// Re-export the layout-layer segment type so consumers can import the full
// page-measurement surface from a single module.
export type { DocumentPageNodeSegment, PaginationSectionMetrics };

// A page node segment extended with the optional table-row-slice field present
// in the upstream DocumentPageNodeSegment (used by the identity-key serializer).
export interface MeasuredDocumentPageNodeSegment extends DocumentPageNodeSegment {
  tableRowSlice?: {
    rowIndex: number;
    startOffsetPx: number;
    sliceHeightPx: number;
  };
}

// Extended pagination-section metrics carrying the column multiplier used by
// the render-time page-content-height resolver. The layout-layer
// PaginationSectionMetrics omits this field; page-measurement augments it.
export interface PageMeasurementPaginationSectionMetrics
  extends PaginationSectionMetrics {
  pageContentHeightMultiplier: number;
}

// Estimate a paragraph's rendered height. Supplied by the table-height /
// line-height estimation cluster via callback to keep this module decoupled.
export type EstimateParagraphHeightPxCallback = (
  paragraph: ParagraphNode,
  pageContentWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet
) => number;

// --- Measured page-content height diagnostics (upstream 2433-2569) ---

function resolveMeasuredPageContentHeightDiagnostics(params: {
  pageLayout: Pick<
    DocumentLayoutMetrics,
    "pageHeightPx" | "marginsPx" | "footerDistancePx"
  >;
  fallbackHeightPx: number;
  headerHeightPx: number;
  currentMeasuredHeightPx?: number;
  bodyTopPx?: number;
  bodyRenderedBottomPx?: number;
  footerTopPx?: number;
  skipBodyBottomAdjustment?: boolean;
}): {
  heightPx: number;
  bodyOverrunsFooter: boolean;
} {
  const {
    pageLayout,
    fallbackHeightPx,
    headerHeightPx,
    currentMeasuredHeightPx,
    bodyTopPx,
    bodyRenderedBottomPx,
    footerTopPx,
    skipBodyBottomAdjustment = false,
  } = params;
  const effectiveBodyRenderedBottomPx = skipBodyBottomAdjustment
    ? undefined
    : bodyRenderedBottomPx;
  const effectiveCurrentMeasuredHeightPx = skipBodyBottomAdjustment
    ? undefined
    : currentMeasuredHeightPx;

  const nominalBodyBottomPx =
    pageLayout.pageHeightPx - pageLayout.marginsPx.bottom;
  const footerOverlapPx = Number.isFinite(footerTopPx)
    ? Math.max(0, Math.round(nominalBodyBottomPx - (footerTopPx as number)))
    : 0;
  const allowedBodyBottomPx = nominalBodyBottomPx - footerOverlapPx;
  const hardFooterBottomLimitPx =
    Number.isFinite(footerTopPx) && footerOverlapPx === 0
      ? Math.round(
          (footerTopPx as number) - UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX
        )
      : undefined;
  const guardedAllowedBodyBottomPx = Number.isFinite(hardFooterBottomLimitPx)
    ? Math.min(allowedBodyBottomPx, hardFooterBottomLimitPx as number)
    : allowedBodyBottomPx;
  const renderedBodyOverrunPx = Number.isFinite(effectiveBodyRenderedBottomPx)
    ? Math.max(
        0,
        Math.round(
          (effectiveBodyRenderedBottomPx as number) - guardedAllowedBodyBottomPx
        )
      )
    : 0;
  const measuredBodyToFooterGapPx = Number.isFinite(
    effectiveBodyRenderedBottomPx
  )
    ? Math.max(
        0,
        Math.round(
          guardedAllowedBodyBottomPx - (effectiveBodyRenderedBottomPx as number)
        )
      )
    : undefined;
  const measuredFooterClearanceBufferPx =
    Number.isFinite(footerTopPx) && Number.isFinite(measuredBodyToFooterGapPx)
      ? Math.max(
          0,
          MEASURED_PAGE_FOOTER_CLEARANCE_BUFFER_PX -
            (measuredBodyToFooterGapPx as number)
        )
      : 0;
  const correctedAllowedBodyBottomPx = Math.max(
    0,
    guardedAllowedBodyBottomPx - measuredFooterClearanceBufferPx
  );
  const iterativeMeasuredHeightPx =
    renderedBodyOverrunPx > 0 &&
    !Number.isFinite(bodyTopPx) &&
    Number.isFinite(effectiveCurrentMeasuredHeightPx) &&
    (effectiveCurrentMeasuredHeightPx as number) > 0
      ? Math.max(
          120,
          Math.round(
            (effectiveCurrentMeasuredHeightPx as number) -
              renderedBodyOverrunPx -
              measuredFooterClearanceBufferPx
          )
        )
      : undefined;
  const bodyOverrunsFooter = renderedBodyOverrunPx > 0;

  if (
    Number.isFinite(bodyTopPx) &&
    correctedAllowedBodyBottomPx > (bodyTopPx as number)
  ) {
    const measuredHeaderClearanceBufferPx =
      headerHeightPx > 0 ? MEASURED_PAGE_HEADER_CLEARANCE_BUFFER_PX : 0;
    const correctedHeightPx = Math.max(
      120,
      Math.round(correctedAllowedBodyBottomPx - (bodyTopPx as number))
    );
    return {
      heightPx: Number.isFinite(iterativeMeasuredHeightPx)
        ? Math.min(
            Math.max(120, correctedHeightPx - measuredHeaderClearanceBufferPx),
            Math.max(
              120,
              (iterativeMeasuredHeightPx as number) -
                measuredHeaderClearanceBufferPx
            )
          )
        : Math.max(120, correctedHeightPx - measuredHeaderClearanceBufferPx),
      bodyOverrunsFooter,
    };
  }

  const measuredHeaderClearanceBufferPx =
    headerHeightPx > 0 ? MEASURED_PAGE_HEADER_CLEARANCE_BUFFER_PX : 0;
  const correctedFallbackHeightPx = Math.max(
    120,
    fallbackHeightPx -
      headerHeightPx -
      measuredHeaderClearanceBufferPx -
      footerOverlapPx -
      renderedBodyOverrunPx -
      measuredFooterClearanceBufferPx
  );
  return {
    heightPx: Number.isFinite(iterativeMeasuredHeightPx)
      ? Math.min(correctedFallbackHeightPx, iterativeMeasuredHeightPx as number)
      : correctedFallbackHeightPx,
    bodyOverrunsFooter,
  };
}

export function resolveMeasuredPageContentHeightPx(params: {
  pageLayout: Pick<
    DocumentLayoutMetrics,
    "pageHeightPx" | "marginsPx" | "footerDistancePx"
  >;
  fallbackHeightPx: number;
  headerHeightPx: number;
  currentMeasuredHeightPx?: number;
  bodyTopPx?: number;
  bodyRenderedBottomPx?: number;
  footerTopPx?: number;
  skipBodyBottomAdjustment?: boolean;
}): number {
  return resolveMeasuredPageContentHeightDiagnostics(params).heightPx;
}

export function resolveMeasuredBodyRenderedBottomPx(
  descendants: Array<{
    bottomPx: number;
    widthPx: number;
    heightPx: number;
    ignore?: boolean;
  }>
): number | undefined {
  let visualBottomPx: number | undefined;

  descendants.forEach((descendant) => {
    if (descendant.ignore) {
      return;
    }
    if (!Number.isFinite(descendant.bottomPx)) {
      return;
    }
    if (descendant.widthPx <= 0 && descendant.heightPx <= 0) {
      return;
    }

    visualBottomPx =
      visualBottomPx === undefined
        ? descendant.bottomPx
        : Math.max(visualBottomPx, descendant.bottomPx);
  });

  return visualBottomPx;
}

export function stabilizeMeasuredPageContentHeights(
  current: number[],
  next: number[],
  options?: {
    currentPageIdentityKeys?: string[];
    nextPageIdentityKeys?: string[];
  }
): number[] {
  return next.map((heightPx, pageIndex) => {
    const roundedNextHeightPx = Math.round(heightPx);
    const currentHeightPx = current[pageIndex];
    const currentPageIdentityKey =
      options?.currentPageIdentityKeys?.[pageIndex];
    const nextPageIdentityKey = options?.nextPageIdentityKeys?.[pageIndex];
    const canPreserveConservativeHeight =
      currentPageIdentityKey === undefined ||
      nextPageIdentityKey === undefined ||
      currentPageIdentityKey === nextPageIdentityKey;
    return Number.isFinite(currentHeightPx)
      ? canPreserveConservativeHeight
        ? Math.min(Math.round(currentHeightPx as number), roundedNextHeightPx)
        : roundedNextHeightPx
      : roundedNextHeightPx;
  });
}

// --- Page-segment identity keys (upstream 2643-2666) ---

export function documentPageNodeSegmentIdentityKey(
  segment: MeasuredDocumentPageNodeSegment
): string {
  const tableRowRangeKey = segment.tableRowRange
    ? `${segment.tableRowRange.startRowIndex}-${segment.tableRowRange.endRowIndex}`
    : "none";
  const tableRowSliceKey = segment.tableRowSlice
    ? `${segment.tableRowSlice.rowIndex}-${Math.round(
        segment.tableRowSlice.startOffsetPx
      )}-${Math.round(segment.tableRowSlice.sliceHeightPx)}`
    : "none";
  const paragraphLineRangeKey = segment.paragraphLineRange
    ? `${segment.paragraphLineRange.startLineIndex}-${segment.paragraphLineRange.endLineIndex}-${segment.paragraphLineRange.totalLineCount}`
    : "none";

  return `${segment.nodeIndex}|${tableRowRangeKey}|${tableRowSliceKey}|${paragraphLineRangeKey}`;
}

export function documentPageNodeSegmentsIdentityKey(
  pageSegments: MeasuredDocumentPageNodeSegment[]
): string {
  return pageSegments.map(documentPageNodeSegmentIdentityKey).join("::");
}

// --- Pagination section metrics (upstream 2667-2793) ---

export function buildPaginationSectionMetrics(
  sections: HeaderFooterDocumentSection[],
  fallbackLayout: DocumentLayoutMetrics,
  reserveOptions: HeaderFooterReserveOptions
): PageMeasurementPaginationSectionMetrics[] {
  const fallbackWidthPx =
    resolveSectionPaginationContentWidthPx(fallbackLayout);
  const fallbackHeightPx = Math.max(
    120,
    fallbackLayout.pageHeightPx -
      fallbackLayout.marginsPx.top -
      fallbackLayout.marginsPx.bottom
  );

  if (sections.length === 0) {
    return [
      {
        startNodeIndex: 0,
        pageContentWidthPx: fallbackWidthPx,
        pageContentHeightPx: fallbackHeightPx,
        pageContentHeightMultiplier: 1,
      },
    ];
  }

  return sections
    .map((section) => {
      const layout = parseSectionLayout(section.sectionPropertiesXml);
      const sectionColumns = parseSectionColumns(section.sectionPropertiesXml);
      const pageContentHeightMultiplier = Math.max(
        1,
        sectionColumns?.count ?? 1
      );
      const hasHeaderContent = sectionHasVisibleHeaderContent(section);
      const hasFooterContent = sectionHasVisibleFooterContent(section);
      const headerPaginationReservePx = hasHeaderContent
        ? resolveHeaderPaginationReservePx(
            section.headerSections ?? [],
            layout,
            reserveOptions
          )
        : 0;
      const footerPaginationReservePx = hasFooterContent
        ? resolveFooterPaginationReservePx(
            section.footerSections ?? [],
            layout,
            reserveOptions
          )
        : 0;
      return {
        startNodeIndex: Math.max(0, Math.round(section.startNodeIndex)),
        pageContentWidthPx: resolveSectionPaginationContentWidthPx(
          layout,
          section.sectionPropertiesXml
        ),
        pageContentHeightPx: Math.max(
          120,
          (layout.pageHeightPx -
            layout.marginsPx.top -
            layout.marginsPx.bottom -
            headerPaginationReservePx -
            footerPaginationReservePx) *
            pageContentHeightMultiplier
        ),
        pageContentHeightMultiplier,
        docGridLinePitchPx: layout.docGridLinePitchPx,
      };
    })
    .sort((left, right) => left.startNodeIndex - right.startNodeIndex);
}

export function scaleMeasuredPageContentHeights(
  measuredPageContentHeightsPxByPageIndex: number[] | undefined,
  heightScale: number
): number[] | undefined {
  if (
    !measuredPageContentHeightsPxByPageIndex ||
    measuredPageContentHeightsPxByPageIndex.length === 0
  ) {
    return measuredPageContentHeightsPxByPageIndex;
  }

  if (!Number.isFinite(heightScale) || Math.abs(heightScale - 1) < 0.001) {
    return measuredPageContentHeightsPxByPageIndex;
  }

  return measuredPageContentHeightsPxByPageIndex.map((heightPx) =>
    Math.max(120, Math.round(heightPx * heightScale))
  );
}

function paragraphSegmentHasPartialLineRange(
  paragraphLineRange?: DocumentPageNodeSegment["paragraphLineRange"]
): boolean {
  if (!paragraphLineRange) {
    return false;
  }

  return (
    paragraphLineRange.startLineIndex > 0 ||
    paragraphLineRange.endLineIndex < paragraphLineRange.totalLineCount
  );
}

export function documentPageContainsOnlySplitParagraphSegments(
  pageSegments: DocumentPageNodeSegment[]
): boolean {
  return (
    pageSegments.length > 0 &&
    pageSegments.every(
      (segment) =>
        !segment.tableRowRange &&
        paragraphSegmentHasPartialLineRange(segment.paragraphLineRange)
    )
  );
}

export function resolvePageContentHeightPxForPageSegments(
  pageSegments: DocumentPageNodeSegment[],
  pageIndex: number,
  defaultPageContentHeightPx: number,
  metricsBySection: PaginationSectionMetrics[],
  measuredPageContentHeightsPxByPageIndex?: number[],
  measuredPageContentIdentityKeysByPageIndex?: string[],
  pageIdentityKey?: string
): number {
  const pageContainsOnlySplitParagraphSegments =
    documentPageContainsOnlySplitParagraphSegments(pageSegments);
  const measuredHeightPx = measuredPageContentHeightsPxByPageIndex?.[pageIndex];
  const measuredHeightMatchesCurrentPage =
    pageIdentityKey === undefined ||
    measuredPageContentIdentityKeysByPageIndex?.[pageIndex] === undefined ||
    measuredPageContentIdentityKeysByPageIndex?.[pageIndex] === pageIdentityKey;
  if (
    Number.isFinite(measuredHeightPx) &&
    measuredHeightMatchesCurrentPage &&
    !pageContainsOnlySplitParagraphSegments
  ) {
    return Math.max(120, Math.round(measuredHeightPx as number));
  }

  const firstNodeIndex = pageSegments[0]?.nodeIndex ?? 0;
  const metricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
    metricsBySection,
    firstNodeIndex,
    0
  );

  return Math.max(
    120,
    Math.round(
      metricsBySection[metricsIndex]?.pageContentHeightPx ??
        defaultPageContentHeightPx
    )
  );
}

export function resolveRenderPageContentHeightPxForPageSegments(params: {
  pageSegments: DocumentPageNodeSegment[];
  pageIndex: number;
  defaultPageContentHeightPx: number;
  metricsBySection: PageMeasurementPaginationSectionMetrics[];
  measuredPageContentHeightsPxByPageIndex?: number[];
  measuredPageContentIdentityKeysByPageIndex?: string[];
  pageIdentityKey?: string;
  useMeasuredPageContentHeights?: boolean;
  pageContentHeightScale?: number;
}): number {
  const firstNodeIndex = params.pageSegments[0]?.nodeIndex ?? 0;
  const metricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
    params.metricsBySection,
    firstNodeIndex,
    0
  );
  const sectionHeightMultiplier = Math.max(
    1,
    Math.round(
      params.metricsBySection[metricsIndex]?.pageContentHeightMultiplier ?? 1
    )
  );
  const pageContainsOnlySplitParagraphSegments =
    documentPageContainsOnlySplitParagraphSegments(params.pageSegments);
  const measuredHeightPx =
    params.useMeasuredPageContentHeights === false
      ? undefined
      : params.measuredPageContentHeightsPxByPageIndex?.[params.pageIndex];
  const measuredHeightMatchesCurrentPage =
    params.pageIdentityKey === undefined ||
    params.measuredPageContentIdentityKeysByPageIndex?.[params.pageIndex] ===
      undefined ||
    params.measuredPageContentIdentityKeysByPageIndex?.[params.pageIndex] ===
      params.pageIdentityKey;
  const usesMeasuredVisualHeight =
    Number.isFinite(measuredHeightPx) &&
    measuredHeightMatchesCurrentPage &&
    !pageContainsOnlySplitParagraphSegments;
  const resolvedHeightPx = resolvePageContentHeightPxForPageSegments(
    params.pageSegments,
    params.pageIndex,
    params.defaultPageContentHeightPx,
    params.metricsBySection,
    params.useMeasuredPageContentHeights === false
      ? undefined
      : params.measuredPageContentHeightsPxByPageIndex,
    params.useMeasuredPageContentHeights === false
      ? undefined
      : params.measuredPageContentIdentityKeysByPageIndex,
    params.pageIdentityKey
  );
  if (usesMeasuredVisualHeight || sectionHeightMultiplier <= 1) {
    return resolvedHeightPx;
  }

  return Math.max(120, Math.round(resolvedHeightPx / sectionHeightMultiplier));
}

// --- Estimated footnote-reserve application (upstream 2906-3053) ---

function estimateDocumentNoteSectionHeightPx(
  notes: DocumentNoteDefinition[],
  pageContentWidthPx: number,
  estimateParagraphHeightPx: EstimateParagraphHeightPxCallback,
  numberingDefinitions?: NumberingDefinitionSet
): number {
  if (notes.length === 0) {
    return 0;
  }

  let totalHeightPx = 19;
  notes.forEach((note, noteIndex) => {
    if (noteIndex > 0) {
      totalHeightPx += 6;
    }

    const noteParagraphs =
      note.nodes?.filter(
        (node): node is ParagraphNode => node.type === "paragraph"
      ) ?? [];
    if (noteParagraphs.length === 0) {
      totalHeightPx += Math.round(12 * 1.35);
      return;
    }

    noteParagraphs.forEach((paragraph, paragraphIndex) => {
      if (paragraphIndex > 0) {
        totalHeightPx += 4;
      }
      totalHeightPx += estimateParagraphHeightPx(
        paragraph,
        pageContentWidthPx,
        numberingDefinitions
      );
    });
  });

  return Math.max(0, totalHeightPx);
}

function collectReferencedFootnotesForPageSegments(
  model: DocModel,
  pageSegments: DocumentPageNodeSegment[],
  footnotesById: Map<number, DocumentNoteDefinition>
): DocumentNoteDefinition[] {
  const referencedIds: number[] = [];
  const seen = new Set<number>();

  pageSegments.forEach((segment) => {
    const node = model.nodes[segment.nodeIndex];
    if (!node) {
      return;
    }

    nodeReferencedNoteIds(
      node,
      "footnote",
      segment.tableRowRange,
      segment.paragraphLineRange
    ).forEach((referenceId) => {
      if (seen.has(referenceId)) {
        return;
      }
      seen.add(referenceId);
      referencedIds.push(referenceId);
    });
  });

  return referencedIds
    .map((referenceId) => footnotesById.get(referenceId))
    .filter((note): note is DocumentNoteDefinition => Boolean(note));
}

export function applyEstimatedFootnoteReserveToPages(
  model: DocModel,
  pages: DocumentPageNodeSegment[][],
  defaultPageContentHeightPx: number,
  pageContentWidthPx: number,
  metricsBySection: PaginationSectionMetrics[],
  numberingDefinitions: NumberingDefinitionSet | undefined,
  footnotes: DocumentNoteDefinition[],
  estimateParagraphHeightPx: EstimateParagraphHeightPxCallback,
  buildPages: (
    pageContentHeightsOverride?: number[]
  ) => DocumentPageNodeSegment[][]
): DocumentPageNodeSegment[][] {
  if (pages.length === 0 || footnotes.length === 0) {
    return pages;
  }

  const footnotesById = new Map(footnotes.map((note) => [note.id, note]));
  let currentPages = pages;

  for (let iteration = 0; iteration < 3; iteration += 1) {
    const reserveOverrides = currentPages.map((pageSegments, pageIndex) => {
      const pageFootnotes = collectReferencedFootnotesForPageSegments(
        model,
        pageSegments,
        footnotesById
      );
      const baseHeightPx = resolvePageContentHeightPxForPageSegments(
        pageSegments,
        pageIndex,
        defaultPageContentHeightPx,
        metricsBySection
      );
      if (pageFootnotes.length === 0) {
        return baseHeightPx;
      }

      const reservePx = estimateDocumentNoteSectionHeightPx(
        pageFootnotes,
        pageContentWidthPx,
        estimateParagraphHeightPx,
        numberingDefinitions
      );
      return Math.max(120, baseHeightPx - reservePx);
    });

    const nextPages = buildPages(reserveOverrides);
    const stable =
      nextPages.length === currentPages.length &&
      nextPages.every((pageSegments, pageIndex) => {
        const previousSegments = currentPages[pageIndex] ?? [];
        if (pageSegments.length !== previousSegments.length) {
          return false;
        }
        return pageSegments.every((segment, segmentIndex) => {
          const previous = previousSegments[segmentIndex];
          return (
            previous?.nodeIndex === segment.nodeIndex &&
            previous?.tableRowRange?.startRowIndex ===
              segment.tableRowRange?.startRowIndex &&
            previous?.tableRowRange?.endRowIndex ===
              segment.tableRowRange?.endRowIndex &&
            previous?.paragraphLineRange?.startLineIndex ===
              segment.paragraphLineRange?.startLineIndex &&
            previous?.paragraphLineRange?.endLineIndex ===
              segment.paragraphLineRange?.endLineIndex
          );
        });
      });

    currentPages = nextPages;
    if (stable) {
      break;
    }
  }

  return currentPages;
}
