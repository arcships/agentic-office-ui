// pagination-plan-stabilize — measured page-content height stabilization,
// page-segment identity keys, pagination section metric construction, and
// measured-height scaling.
//
// Migrated from @extend-ai/react-docx editor.tsx lines:
//   2617-2641  (stabilizeMeasuredPageContentHeights)
//   2643-2666  (documentPageNodeSegmentIdentityKey, documentPageNodeSegmentsIdentityKey)
//   2667-2727  (buildPaginationSectionMetrics)
//   2774-2793  (scaleMeasuredPageContentHeights)

import type { DocumentLayoutMetrics } from "../../viewer/section-layout";
import { parseSectionLayout } from "../../viewer/section-layout";
import type { PaginationSectionMetrics } from "../../layout/pagination";
import type {
  DocumentPageNodeSegment,
} from "../../layout/page-segmentation-core";
import type {
  HeaderFooterDocumentSection,
  HeaderFooterReserveOptions
} from "./header-footer";
import {
  parseSectionColumns,
  resolveHeaderPaginationReservePx,
  resolveFooterPaginationReservePx,
  resolveSectionPaginationContentWidthPx,
  sectionHasVisibleHeaderContent,
  sectionHasVisibleFooterContent
} from "./header-footer";

// Re-export the layout-layer segment type so consumers can import the full
// pagination-plan-stabilize surface from a single module.
export type { DocumentPageNodeSegment } from "../../layout/page-segmentation-core";
export type { PaginationSectionMetrics } from "../../layout/pagination";

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

// --- Measured page-content height stabilization (upstream 2617-2641) ---

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

// --- Pagination section metrics (upstream 2667-2727) ---

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

// --- Measured page-content height scaling (upstream 2774-2793) ---

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
