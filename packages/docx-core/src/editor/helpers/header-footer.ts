// Section parsing, header/footer reference inheritance, and header/footer
// pagination reserve computation.
// Upstream editor.tsx: lines 1692-2431.
//
// Parses <w:sectPr> column layout, page-number start/format, and section start
// type; inherits header/footer references across sections; and estimates how
// much vertical space a section's header/footer content reserves on the page
// (including floating-image intrusion). The reserve entry points accept a
// `estimateDocNodeHeightPx` callback so this module stays decoupled from the
// table-height / line-height estimation cluster (mirrors the layout layer's
// callback-driven page-segmentation contract).

import type {
  DocModel,
  FooterSection,
  HeaderSection,
  ImageRunNode,
  NumberingDefinitionSet,
  ParagraphNode
} from "../../engine/types";
import type { DocumentLayoutMetrics } from "../../viewer/section-layout";
import { twipsToPixels } from "../../viewer/section-layout";
import {
  DEFAULT_PAGE_NUMBER_START,
  FLOATING_FOOTER_BASELINE_CLEARANCE_RESERVE_PX,
  FLOATING_HEADER_FOOTER_CLEARANCE_BUFFER_PX,
  MIN_PARAGRAPH_LINE_HEIGHT_PX,
  MIN_VISIBLE_FLOW_FOOTER_PAGINATION_RESERVE_PX,
  UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX
} from "./constants";
import {
  paragraphHasFormField,
  paragraphHasVisibleText,
  shouldRenderAbsoluteFloatingImage,
  shouldRenderWrappedFloatingImage,
  shouldReserveHeaderFooterFloatingImageSpace
} from "./paragraph-geometry";
import { paragraphHasImage } from "./paragraph-inspect";

interface SectionColumnLayout {
  count: number;
  gapPx: number;
  widthsPx?: number[];
}

export function parseSectionColumns(
  sectionPropertiesXml?: string
): SectionColumnLayout | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const columnsTag = sectionPropertiesXml.match(/<w:cols\b[^>]*\/?>/i)?.[0];
  if (!columnsTag) {
    return undefined;
  }

  const numberOfColumnsRaw = columnsTag.match(/w:num="([\d.]+)"/i)?.[1];
  const numberOfColumns = numberOfColumnsRaw
    ? Math.round(Number(numberOfColumnsRaw))
    : 1;
  if (!Number.isFinite(numberOfColumns) || numberOfColumns <= 1) {
    return undefined;
  }
  const columnCount = Math.max(2, numberOfColumns);

  const columnTags = [...sectionPropertiesXml.matchAll(/<w:col\b[^>]*\/>/gi)];
  const widthsTwips = columnTags
    .map((match) => Number(match[0].match(/w:w="([\d.]+)"/i)?.[1]))
    .filter((value) => Number.isFinite(value) && value > 0);
  const widthsPx = widthsTwips.map((value) =>
    Math.max(1, Math.round(twipsToPixels(value) ?? 0))
  );

  // Gap resolution: explicit space on the cols tag, then per-col space, then
  // — when every column width is declared — derived from the section body
  // width (Word keeps widths + gaps equal to the body), then Word's default.
  const colsTagSpaceRaw = columnsTag.match(/w:space="([\d.]+)"/i)?.[1];
  const firstColSpaceRaw = columnTags
    .map((match) => match[0].match(/w:space="([\d.]+)"/i)?.[1])
    .find((value) => value !== undefined);
  let columnGapTwips =
    colsTagSpaceRaw !== undefined
      ? Number(colsTagSpaceRaw)
      : firstColSpaceRaw !== undefined
      ? Number(firstColSpaceRaw)
      : undefined;
  if (columnGapTwips === undefined && widthsTwips.length === columnCount) {
    const pageWidthTwips = Number(
      sectionPropertiesXml.match(/<w:pgSz\b[^>]*w:w="([\d.]+)"/i)?.[1]
    );
    const marginTag = sectionPropertiesXml.match(/<w:pgMar\b[^>]*\/?>/i)?.[0];
    const marginLeftTwips = Number(marginTag?.match(/w:left="([\d.-]+)"/i)?.[1]);
    const marginRightTwips = Number(
      marginTag?.match(/w:right="([\d.-]+)"/i)?.[1]
    );
    if (
      Number.isFinite(pageWidthTwips) &&
      Number.isFinite(marginLeftTwips) &&
      Number.isFinite(marginRightTwips)
    ) {
      const bodyTwips = pageWidthTwips - marginLeftTwips - marginRightTwips;
      const totalWidthTwips = widthsTwips.reduce((sum, value) => sum + value, 0);
      columnGapTwips = Math.max(
        0,
        (bodyTwips - totalWidthTwips) / Math.max(1, columnCount - 1)
      );
    }
  }
  if (columnGapTwips === undefined || !Number.isFinite(columnGapTwips)) {
    columnGapTwips = 720;
  }
  const columnGapPx = twipsToPixels(columnGapTwips) ?? 24;

  return {
    count: columnCount,
    gapPx: Math.max(0, columnGapPx),
    ...(widthsPx.length === columnCount ? { widthsPx } : undefined),
  };
}

export function resolveSectionPaginationContentWidthPx(
  layout: Pick<DocumentLayoutMetrics, "pageWidthPx" | "marginsPx">,
  sectionPropertiesXml?: string
): number {
  const bodyWidthPx = Math.max(
    120,
    layout.pageWidthPx - layout.marginsPx.left - layout.marginsPx.right
  );
  const columns = parseSectionColumns(sectionPropertiesXml);
  if (!columns) {
    return bodyWidthPx;
  }

  const totalGapPx =
    Math.max(0, columns.gapPx) * Math.max(0, columns.count - 1);
  return Math.max(120, Math.round((bodyWidthPx - totalGapPx) / columns.count));
}

export function parseSectionPageNumberStart(sectionPropertiesXml?: string): number {
  if (!sectionPropertiesXml) {
    return DEFAULT_PAGE_NUMBER_START;
  }

  const pageNumberTag = sectionPropertiesXml.match(
    /<w:pgNumType\b[^>]*\/?>/i
  )?.[0];
  if (!pageNumberTag) {
    return DEFAULT_PAGE_NUMBER_START;
  }

  const startRaw = pageNumberTag.match(/\bw:start="(\d+)"/i)?.[1];
  if (!startRaw) {
    return DEFAULT_PAGE_NUMBER_START;
  }

  const start = Number(startRaw);
  if (!Number.isFinite(start) || start <= 0) {
    return DEFAULT_PAGE_NUMBER_START;
  }

  return Math.max(1, Math.round(start));
}

export function parseSectionPageNumberStartOverride(
  sectionPropertiesXml?: string
): number | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const pageNumberTag = sectionPropertiesXml.match(
    /<w:pgNumType\b[^>]*\/?>/i
  )?.[0];
  if (!pageNumberTag) {
    return undefined;
  }

  const startRaw = pageNumberTag.match(/\bw:start="(\d+)"/i)?.[1];
  if (!startRaw) {
    return undefined;
  }

  const start = Number(startRaw);
  if (!Number.isFinite(start) || start <= 0) {
    return undefined;
  }

  return Math.max(1, Math.round(start));
}

export function parseSectionPageNumberFormat(
  sectionPropertiesXml?: string
): string | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const pageNumberTag = sectionPropertiesXml.match(
    /<w:pgNumType\b[^>]*\/?>/i
  )?.[0];
  const format = pageNumberTag?.match(/\bw:fmt="([^"]+)"/i)?.[1]?.trim();
  return format && format.length > 0 ? format : undefined;
}

export function parseSectionStartType(
  sectionPropertiesXml?: string
): string | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const sectionType = sectionPropertiesXml
    .match(/<w:type\b[^>]*w:val="([^"]+)"/i)?.[1]
    ?.trim();
  return sectionType && sectionType.length > 0
    ? sectionType.toLowerCase()
    : undefined;
}

export interface HeaderFooterDocumentSection {
  startNodeIndex: number;
  sectionPropertiesXml?: string;
  headerSections: HeaderSection[];
  footerSections: FooterSection[];
}

export function normalizeSectionReferenceType(referenceType?: string): string {
  const normalized = referenceType?.trim().toLowerCase();
  return normalized && normalized.length > 0 ? normalized : "default";
}

function inheritSectionReferences<T extends HeaderSection | FooterSection>(
  sections: HeaderFooterDocumentSection[],
  sectionKey: "headerSections" | "footerSections"
): HeaderFooterDocumentSection[] {
  const inheritedByType = new Map<string, T>();

  return sections.map((section) => {
    const explicitSections = section[sectionKey] as T[];
    if (explicitSections.length > 0) {
      explicitSections.forEach((entry) => {
        inheritedByType.set(
          normalizeSectionReferenceType(entry.referenceType),
          entry
        );
      });
    }

    return {
      ...section,
      [sectionKey]: [...inheritedByType.values()],
    };
  });
}

function resolveInheritedSectionHeaderFooterReferences(
  sections: HeaderFooterDocumentSection[]
): HeaderFooterDocumentSection[] {
  if (sections.length === 0) {
    return sections;
  }

  return inheritSectionReferences(
    inheritSectionReferences(sections, "headerSections"),
    "footerSections"
  );
}

function resolveDocumentSectionsFromMetadata(
  metadata: DocModel["metadata"]
): HeaderFooterDocumentSection[] {
  const normalizedSections = (metadata.sections ?? [])
    .map(
      (section): HeaderFooterDocumentSection => ({
        startNodeIndex:
          Number.isFinite(section.startNodeIndex) &&
          (section.startNodeIndex as number) >= 0
            ? Math.round(section.startNodeIndex as number)
            : 0,
        sectionPropertiesXml: section.sectionPropertiesXml,
        headerSections: section.headerSections ?? [],
        footerSections: section.footerSections ?? [],
      })
    )
    .sort((left, right) => left.startNodeIndex - right.startNodeIndex);

  if (normalizedSections.length > 0) {
    if (normalizedSections[0].startNodeIndex > 0) {
      normalizedSections.unshift({
        startNodeIndex: 0,
        sectionPropertiesXml: normalizedSections[0].sectionPropertiesXml,
        headerSections: normalizedSections[0].headerSections,
        footerSections: normalizedSections[0].footerSections,
      });
    }
    return resolveInheritedSectionHeaderFooterReferences(normalizedSections);
  }

  return [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: metadata.sectionPropertiesXml,
      headerSections: metadata.headerSections ?? [],
      footerSections: metadata.footerSections ?? [],
    },
  ];
}

function paragraphHasHeaderFooterReserveRelevantContent(
  paragraph: ParagraphNode
): boolean {
  if (paragraphHasVisibleText(paragraph) || paragraphHasFormField(paragraph)) {
    return true;
  }

  if (!paragraphHasImage(paragraph)) {
    return false;
  }

  return paragraph.children.some((child) => {
    if (child.type !== "image") {
      return false;
    }

    // Inline/wrapped images are part of the header/footer flow and should keep
    // reserve. Absolute behind-text anchors (watermarks/background art) should not.
    if (!child.floating || shouldRenderWrappedFloatingImage(child)) {
      return true;
    }

    if (!shouldRenderAbsoluteFloatingImage(child)) {
      return true;
    }

    return shouldReserveHeaderFooterFloatingImageSpace(child);
  });
}

export function sectionHasVisibleHeaderContent(
  section: HeaderFooterDocumentSection
): boolean {
  const headerSections = section.headerSections ?? [];
  return headerSections.some((headerSection) =>
    (headerSection.nodes ?? []).some(
      (node) =>
        node.type === "table" ||
        paragraphHasHeaderFooterReserveRelevantContent(node)
    )
  );
}

export function sectionHasVisibleFooterContent(
  section: HeaderFooterDocumentSection
): boolean {
  const footerSections = section.footerSections ?? [];
  return footerSections.some((footerSection) =>
    (footerSection.nodes ?? []).some(
      (node) =>
        node.type === "table" ||
        paragraphHasHeaderFooterReserveRelevantContent(node)
    )
  );
}

function resolveHeaderFooterAbsoluteFloatingTopPx(
  image: ImageRunNode,
  layout: Pick<DocumentLayoutMetrics, "marginsPx">
): number | undefined {
  if (!shouldRenderAbsoluteFloatingImage(image) || !image.floating) {
    return undefined;
  }

  const floating = image.floating;
  const verticalRelativeTo = floating.verticalRelativeTo?.trim().toLowerCase();
  if (!Number.isFinite(floating.yPx)) {
    return undefined;
  }

  if (verticalRelativeTo === "page") {
    return Math.round(floating.yPx as number);
  }

  if (verticalRelativeTo === "margin") {
    return Math.round((floating.yPx as number) + layout.marginsPx.top);
  }

  return undefined;
}

function resolveFooterParagraphFloatingBoundaryTopPx(
  paragraph: ParagraphNode,
  layout: Pick<DocumentLayoutMetrics, "marginsPx">
): number | undefined {
  if (!paragraphHasImage(paragraph)) {
    return undefined;
  }

  let boundaryTopPx: number | undefined;

  paragraph.children.forEach((child) => {
    if (
      child.type !== "image" ||
      !child.floating ||
      !shouldReserveHeaderFooterFloatingImageSpace(child)
    ) {
      return;
    }

    const imageTopPx = resolveHeaderFooterAbsoluteFloatingTopPx(child, layout);
    if (!Number.isFinite(imageTopPx)) {
      return;
    }

    const distTPx = Math.max(0, Math.round(child.floating.distTPx ?? 0));
    const candidateTopPx = Math.max(
      0,
      Math.round((imageTopPx as number) - distTPx)
    );
    boundaryTopPx =
      boundaryTopPx === undefined
        ? candidateTopPx
        : Math.min(boundaryTopPx, candidateTopPx);
  });

  return boundaryTopPx;
}

function resolveFooterNodesFloatingBoundaryTopPx(
  footerNodes: DocModel["nodes"],
  layout: Pick<DocumentLayoutMetrics, "marginsPx">
): number | undefined {
  let boundaryTopPx: number | undefined;

  footerNodes.forEach((node) => {
    if (node.type !== "paragraph") {
      return;
    }

    const paragraphBoundaryTopPx = resolveFooterParagraphFloatingBoundaryTopPx(
      node,
      layout
    );
    if (!Number.isFinite(paragraphBoundaryTopPx)) {
      return;
    }

    boundaryTopPx =
      boundaryTopPx === undefined
        ? Math.round(paragraphBoundaryTopPx as number)
        : Math.min(boundaryTopPx, Math.round(paragraphBoundaryTopPx as number));
  });

  return boundaryTopPx;
}

function estimateHeaderFooterParagraphFloatingReservePx(
  paragraph: ParagraphNode,
  layout: Pick<
    DocumentLayoutMetrics,
    "pageWidthPx" | "pageHeightPx" | "marginsPx"
  > &
    Partial<Pick<DocumentLayoutMetrics, "headerDistancePx">>,
  region: "header" | "footer"
): number {
  if (!paragraphHasImage(paragraph)) {
    return 0;
  }

  const nominalBodyTopPx = layout.marginsPx.top;
  const nominalBodyBottomPx = layout.pageHeightPx - layout.marginsPx.bottom;

  if (region === "footer") {
    const floatingBoundaryTopPx = resolveFooterParagraphFloatingBoundaryTopPx(
      paragraph,
      layout
    );
    return Number.isFinite(floatingBoundaryTopPx)
      ? Math.max(
          0,
          nominalBodyBottomPx -
            Math.round(floatingBoundaryTopPx as number) +
            FLOATING_HEADER_FOOTER_CLEARANCE_BUFFER_PX
        )
      : 0;
  }

  return paragraph.children.reduce((largest, child) => {
    if (child.type !== "image" || !child.floating) {
      return largest;
    }

    const floating = child.floating;
    const imageHeightPx =
      Number.isFinite(child.heightPx) && (child.heightPx as number) > 0
        ? Math.round(child.heightPx as number)
        : Number.isFinite(child.widthPx) && (child.widthPx as number) > 0
        ? Math.round(child.widthPx as number)
        : MIN_PARAGRAPH_LINE_HEIGHT_PX;
    const distTPx = Math.max(0, Math.round(floating.distTPx ?? 0));
    const distBPx = Math.max(0, Math.round(floating.distBPx ?? 0));

    let topPx = resolveHeaderFooterAbsoluteFloatingTopPx(child, layout);
    if (!Number.isFinite(topPx)) {
      const verticalRelativeTo = floating.verticalRelativeTo
        ?.trim()
        .toLowerCase();
      const isParagraphRelativeAnchor =
        verticalRelativeTo === undefined ||
        verticalRelativeTo === "" ||
        verticalRelativeTo === "paragraph" ||
        verticalRelativeTo === "line";
      if (region === "header" && isParagraphRelativeAnchor) {
        const headerAnchorOriginPx = Math.max(
          0,
          Math.round(layout.headerDistancePx ?? layout.marginsPx.top)
        );
        topPx = headerAnchorOriginPx + Math.round(floating.yPx ?? 0);
      }
    }
    if (!Number.isFinite(topPx)) {
      return largest;
    }

    const resolvedTopPx = Math.round(topPx as number);
    const resolvedBottomPx = resolvedTopPx + imageHeightPx + distTPx + distBPx;
    const reserveBehindDocHeaderBand =
      region === "header" &&
      floating.behindDocument === true &&
      Math.max(1, Math.round(child.widthPx ?? 0)) >=
        Math.round(layout.pageWidthPx * 0.5) &&
      resolvedTopPx <= nominalBodyTopPx + 192 &&
      resolvedBottomPx > nominalBodyTopPx;
    if (
      !shouldReserveHeaderFooterFloatingImageSpace(child) &&
      !reserveBehindDocHeaderBand
    ) {
      return largest;
    }

    const bodyOverlapPx = Math.max(0, resolvedBottomPx - nominalBodyTopPx);
    if (bodyOverlapPx <= 0) {
      return largest;
    }

    return Math.max(
      largest,
      bodyOverlapPx + FLOATING_HEADER_FOOTER_CLEARANCE_BUFFER_PX
    );
  }, 0);
}

export interface HeaderFooterReserveOptions {
  estimateDocNodeHeightPx: (
    node: DocModel["nodes"][number],
    availableWidthPx: number,
    numberingDefinitions: NumberingDefinitionSet | undefined,
    docGridLinePitchPx?: number
  ) => number;
}

export function resolveHeaderPaginationReservePx(
  headerSections: HeaderSection[],
  layout: Pick<
    DocumentLayoutMetrics,
    | "pageWidthPx"
    | "pageHeightPx"
    | "marginsPx"
    | "headerDistancePx"
    | "docGridLinePitchPx"
  >,
  options: HeaderFooterReserveOptions
): number {
  const visibleHeaderSections = (headerSections ?? []).filter((headerSection) =>
    (headerSection.nodes ?? []).some(
      (node) =>
        node.type === "table" ||
        paragraphHasHeaderFooterReserveRelevantContent(node)
    )
  );
  if (visibleHeaderSections.length === 0) {
    return 0;
  }

  const availableWidthPx = Math.max(
    24,
    layout.pageWidthPx - layout.marginsPx.left - layout.marginsPx.right
  );
  const estimatedHeaderHeightPx = visibleHeaderSections.reduce(
    (largestHeightPx, headerSection) => {
      const visibleNodes = (headerSection.nodes ?? []).filter(
        (node) =>
          node.type === "table" ||
          paragraphHasHeaderFooterReserveRelevantContent(node)
      );
      if (visibleNodes.length === 0) {
        return largestHeightPx;
      }

      const nodeHeightsPx = visibleNodes.reduce((sum, node) => {
        const floatingReservePx =
          node.type === "paragraph"
            ? estimateHeaderFooterParagraphFloatingReservePx(
                node,
                layout,
                "header"
              )
            : 0;
        return (
          sum +
          Math.max(
            options.estimateDocNodeHeightPx(
              node,
              availableWidthPx,
              undefined,
              layout.docGridLinePitchPx
            ),
            floatingReservePx
          )
        );
      }, 0);
      const interParagraphGapPx = Math.max(0, visibleNodes.length - 1) * 8;
      return Math.max(
        largestHeightPx,
        Math.round(nodeHeightsPx + interParagraphGapPx)
      );
    },
    0
  );

  if (estimatedHeaderHeightPx <= 0) {
    return 0;
  }

  const headerBodyOverlapPx =
    estimatedHeaderHeightPx + layout.headerDistancePx - layout.marginsPx.top;
  return Math.max(0, Math.round(headerBodyOverlapPx));
}

export function resolveFooterPaginationReservePx(
  footerSections: FooterSection[],
  layout: Pick<
    DocumentLayoutMetrics,
    | "pageWidthPx"
    | "pageHeightPx"
    | "marginsPx"
    | "footerDistancePx"
    | "docGridLinePitchPx"
  >,
  options: HeaderFooterReserveOptions
): number {
  const visibleFooterSections = (footerSections ?? []).filter((footerSection) =>
    (footerSection.nodes ?? []).some(
      (node) =>
        node.type === "table" ||
        paragraphHasHeaderFooterReserveRelevantContent(node)
    )
  );
  if (visibleFooterSections.length === 0) {
    return 0;
  }

  const availableWidthPx = Math.max(
    24,
    layout.pageWidthPx - layout.marginsPx.left - layout.marginsPx.right
  );
  const estimatedFooterHeightPx = visibleFooterSections.reduce(
    (largestHeightPx, footerSection) => {
      const visibleNodes = (footerSection.nodes ?? []).filter(
        (node) =>
          node.type === "table" ||
          paragraphHasHeaderFooterReserveRelevantContent(node)
      );
      if (visibleNodes.length === 0) {
        return largestHeightPx;
      }

      const nodeHeightsPx = visibleNodes.reduce((sum, node) => {
        const floatingReservePx =
          node.type === "paragraph"
            ? estimateHeaderFooterParagraphFloatingReservePx(
                node,
                layout,
                "footer"
              )
            : 0;
        return (
          sum +
          Math.max(
            options.estimateDocNodeHeightPx(
              node,
              availableWidthPx,
              undefined,
              layout.docGridLinePitchPx
            ),
            floatingReservePx
          )
        );
      }, 0);
      const interParagraphGapPx = Math.max(0, visibleNodes.length - 1) * 8;
      return Math.max(
        largestHeightPx,
        Math.round(nodeHeightsPx + interParagraphGapPx)
      );
    },
    0
  );

  if (estimatedFooterHeightPx <= 0) {
    return 0;
  }

  const nominalBodyBottomPx = layout.pageHeightPx - layout.marginsPx.bottom;
  const floatingFooterBoundaryTopPx = visibleFooterSections.reduce<
    number | undefined
  >((smallestTopPx, footerSection) => {
    const sectionFloatingTopPx = resolveFooterNodesFloatingBoundaryTopPx(
      footerSection.nodes ?? [],
      layout
    );
    if (!Number.isFinite(sectionFloatingTopPx)) {
      return smallestTopPx;
    }

    return smallestTopPx === undefined
      ? Math.round(sectionFloatingTopPx as number)
      : Math.min(smallestTopPx, Math.round(sectionFloatingTopPx as number));
  }, undefined);
  const hasFloatingFooterBodyIntrusionRisk = visibleFooterSections.some(
    (footerSection) =>
      (footerSection.nodes ?? []).some(
        (node) =>
          node.type === "paragraph" &&
          estimateHeaderFooterParagraphFloatingReservePx(
            node,
            layout,
            "footer"
          ) > 0
      )
  );

  const footerBodyOverlapPx =
    estimatedFooterHeightPx + layout.footerDistancePx - layout.marginsPx.bottom;
  const explicitFooterBoundaryReservePx = Math.max(
    0,
    Math.round(footerBodyOverlapPx + UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX)
  );
  const floatingFooterBoundaryReservePx = Number.isFinite(
    floatingFooterBoundaryTopPx
  )
    ? Math.max(
        0,
        Math.round(
          nominalBodyBottomPx -
            Math.round(floatingFooterBoundaryTopPx as number) +
            UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX
        )
      )
    : 0;

  const resolvedReservePx = hasFloatingFooterBodyIntrusionRisk
    ? Math.max(
        explicitFooterBoundaryReservePx,
        floatingFooterBoundaryReservePx,
        FLOATING_FOOTER_BASELINE_CLEARANCE_RESERVE_PX
      )
    : Math.max(
        explicitFooterBoundaryReservePx,
        floatingFooterBoundaryReservePx
      );

  return Math.max(
    estimatedFooterHeightPx > 0
      ? MIN_VISIBLE_FLOW_FOOTER_PAGINATION_RESERVE_PX
      : 0,
    resolvedReservePx
  );
}
