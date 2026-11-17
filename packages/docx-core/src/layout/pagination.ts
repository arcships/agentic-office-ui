import type {
  DocModel,
  FooterSection,
  HeaderSection,
  ParagraphNode,
  TableCellContentNode,
  TableNode
} from "../engine/types";
import {
  resolveDocumentSectionsFromMetadata,
  resolveSectionIndexForNodeIndex,
  sectionBreakPropertiesStartNewPage,
  type ResolvedModelSection
} from "./sections";

export {
  resolveDocumentSectionsFromMetadata,
  resolveSectionIndexForNodeIndex,
  sectionBreakPropertiesStartNewPage,
  sectionBreakRequiredPageParity
} from "./sections";
export type {
  ResolvedModelSection,
  SectionStartPageParity
} from "./sections";

const TWIPS_PER_PIXEL = 15;
const PAGE_BREAK_XML_PATTERN = /<w:br\b[^>]*w:type="page"[^>]*\/?>/i;
const LAST_RENDERED_PAGE_BREAK_XML_PATTERN = /<w:lastRenderedPageBreak\b[^>]*\/?>/i;
const PAGE_BREAK_BEFORE_XML_PATTERN = /<w:pageBreakBefore\b[^>]*\/?>/i;
// Tokenizes the paragraph's inline story in document order: text runs, tabs,
// line breaks and page breaks. Used to map an inline page break to a text
// offset inside the paragraph instead of treating the whole paragraph as a
// node-level page start.
const PARAGRAPH_INLINE_TOKEN_PATTERN =
  /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:br\b[^>]*\/?>|<w:cr\b[^>]*\/?>|<w:tab\b[^>]*\/?>|<w:noBreakHyphen\b[^>]*\/?>/gi;
const XML_ENTITY_PATTERN = /&(?:#(\d+)|#x([0-9a-f]+)|(amp|lt|gt|quot|apos));/gi;
const SECTION_PROPERTIES_XML_PATTERN = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/i;
const paragraphBreakFlagsByParagraph = new WeakMap<
  ParagraphNode,
  {
    sourceXml: string;
    flags: {
      explicitPageBreak: boolean;
      lastRenderedPageBreak: boolean;
      pageBreakBefore: boolean;
      sectionBreakStartsNewPage: boolean;
    };
  }
>();
const tableExplicitPageBreakInfoByTable = new WeakMap<
  TableNode,
  {
    sourceXml: string;
    info: TableExplicitPageBreakInfo;
  }
>();

type ParagraphBreakFlags = {
  explicitPageBreak: boolean;
  lastRenderedPageBreak: boolean;
  pageBreakBefore: boolean;
  sectionBreakStartsNewPage: boolean;
};

export interface PaginationSectionMetrics {
  startNodeIndex: number;
  pageContentWidthPx: number;
  pageContentHeightPx: number;
  pageContentHeightMultiplier?: number;
  docGridLinePitchPx?: number;
}

export interface TableExplicitPageBreakInfo {
  startRowIndexes: number[];
  breakAfterTable: boolean;
}

export interface DocumentPageRange {
  startNodeIndex: number;
  endNodeIndex: number;
}

function twipsToPixels(twips?: number): number | undefined {
  if (!Number.isFinite(twips)) {
    return undefined;
  }

  return Math.max(0, Math.round((twips as number) / TWIPS_PER_PIXEL));
}

function isOnOffTagEnabled(tagXml: string | undefined): boolean {
  if (!tagXml) {
    return false;
  }

  const valueMatch = tagXml.match(/\bw:val="([^"]+)"/i)?.[1]?.trim().toLowerCase();
  if (!valueMatch) {
    return true;
  }

  return !["0", "false", "off", "no"].includes(valueMatch);
}

function paragraphHasImage(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => child.type === "image");
}

function paragraphHasFormField(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => child.type === "form-field");
}

function paragraphHasVisibleText(paragraph: ParagraphNode): boolean {
  return paragraph.children.some(
    (child) =>
      (child.type === "text" && child.text.trim().length > 0) ||
      (child.type === "form-field" && (child.value ?? "").trim().length > 0)
  );
}

function paragraphIsOnlyExplicitPageBreak(paragraph: ParagraphNode): boolean {
  if (!paragraphHasExplicitPageBreak(paragraph)) {
    return false;
  }

  return (
    !paragraphHasVisibleText(paragraph) &&
    !paragraphHasImage(paragraph) &&
    !paragraphHasFormField(paragraph)
  );
}

function tableCellDirectParagraphs(nodeContent: TableCellContentNode[]): ParagraphNode[] {
  return nodeContent.filter((entry): entry is ParagraphNode => entry.type === "paragraph");
}

function tableRowHasSubstantiveContentOutsideBreakOnlyParagraphs(
  row: TableNode["rows"][number]
): boolean {
  return row.cells.some((cell) =>
    tableCellDirectParagraphs(cell.nodes).some((paragraph) => {
      if (paragraphIsOnlyExplicitPageBreak(paragraph)) {
        return false;
      }

      return (
        paragraphHasVisibleText(paragraph) ||
        paragraphHasImage(paragraph) ||
        paragraphHasFormField(paragraph) ||
        paragraphHasExplicitPageBreak(paragraph)
      );
    })
  );
}

function tableRowSubstantiveCellIndexes(row: TableNode["rows"][number]): number[] {
  const indexes: number[] = [];
  row.cells.forEach((cell, cellIndex) => {
    const substantive = tableCellDirectParagraphs(cell.nodes).some((paragraph) => {
      if (paragraphIsOnlyExplicitPageBreak(paragraph)) {
        return false;
      }

      return (
        paragraphHasVisibleText(paragraph) ||
        paragraphHasImage(paragraph) ||
        paragraphHasFormField(paragraph) ||
        paragraphHasExplicitPageBreak(paragraph)
      );
    });
    if (substantive) {
      indexes.push(cellIndex);
    }
  });
  return indexes;
}

function tableRowUsesTrailingSignatureCellBreakPattern(
  table: TableNode,
  row: TableNode["rows"][number]
): boolean {
  if (table.rows.length !== 1 || row.cells.length < 4) {
    return false;
  }

  const substantiveIndexes = tableRowSubstantiveCellIndexes(row);
  return substantiveIndexes.length === 1 && substantiveIndexes[0] >= row.cells.length - 1;
}

function resolveParagraphBreakFlags(paragraph: ParagraphNode): ParagraphBreakFlags {
  const sourceXml = paragraph.sourceXml ?? "";
  const cached = paragraphBreakFlagsByParagraph.get(paragraph);
  if (cached?.sourceXml === sourceXml) {
    return cached.flags;
  }

  const flags = {
    explicitPageBreak: PAGE_BREAK_XML_PATTERN.test(sourceXml),
    lastRenderedPageBreak: LAST_RENDERED_PAGE_BREAK_XML_PATTERN.test(sourceXml),
    pageBreakBefore: isOnOffTagEnabled(sourceXml.match(PAGE_BREAK_BEFORE_XML_PATTERN)?.[0]),
    sectionBreakStartsNewPage: (() => {
      const sectionProperties = sourceXml.match(SECTION_PROPERTIES_XML_PATTERN)?.[0];
      if (!sectionProperties) {
        return false;
      }

      return sectionBreakPropertiesStartNewPage(sectionProperties);
    })()
  };
  paragraphBreakFlagsByParagraph.set(paragraph, { flags, sourceXml });
  return flags;
}

export function paragraphHasExplicitPageBreak(paragraph: ParagraphNode): boolean {
  if (!paragraph.sourceXml) {
    return false;
  }

  return resolveParagraphBreakFlags(paragraph).explicitPageBreak;
}

export function paragraphHasPageBreakBefore(paragraph: ParagraphNode): boolean {
  if (paragraph.style?.pageBreakBefore === true) {
    return true;
  }

  if (!paragraph.sourceXml) {
    return false;
  }

  return resolveParagraphBreakFlags(paragraph).pageBreakBefore;
}

export function sectionBreakAfterParagraphStartsNewPage(paragraph: ParagraphNode): boolean {
  if (!paragraph.sourceXml) {
    return false;
  }

  return resolveParagraphBreakFlags(paragraph).sectionBreakStartsNewPage;
}

export function paragraphHasLastRenderedPageBreak(paragraph: ParagraphNode): boolean {
  if (!paragraph.sourceXml) {
    return false;
  }

  return resolveParagraphBreakFlags(paragraph).lastRenderedPageBreak;
}

export type LastRenderedPageBreakPosition = "none" | "leading" | "internal";

/**
 * Where Word's cached `w:lastRenderedPageBreak` marker sits inside the
 * paragraph. Word writes the marker at the previous render's break position,
 * which is usually mid-paragraph; only a marker with nothing but run/paragraph
 * boilerplate before it denotes a paragraph that starts a page.
 */
export function paragraphLastRenderedPageBreakPosition(
  paragraph: ParagraphNode
): LastRenderedPageBreakPosition {
  if (!paragraphHasLastRenderedPageBreak(paragraph)) {
    return "none";
  }

  return paragraphStartsWithLastRenderedPageBreak(paragraph) ? "leading" : "internal";
}

export function paragraphStartsWithLastRenderedPageBreak(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml || !paragraphHasLastRenderedPageBreak(paragraph)) {
    return false;
  }

  const breakMatch = xml.match(LAST_RENDERED_PAGE_BREAK_XML_PATTERN);
  if (!breakMatch || breakMatch.index === undefined) {
    return false;
  }

  const leadingXml = xml
    .slice(0, breakMatch.index)
    .replace(/^<w:p\b[^>]*>/i, "")
    .replace(/<w:pPr\b(?:[^/>]*\/>|[\s\S]*?<\/w:pPr>)/i, "")
    .replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/gi, "")
    .replace(/<\/?w:r\b[^>]*>/gi, "")
    .replace(/<w:(?:proofErr|bookmarkStart|bookmarkEnd|permStart|permEnd)\b[^>]*\/?>/gi, "")
    .replace(/<\/?w:(?:ins|smartTag)\b[^>]*>/gi, "")
    .replace(/\s+/g, "");

  return leadingXml.length === 0;
}

function decodeXmlText(xmlText: string): string {
  return xmlText.replace(
    XML_ENTITY_PATTERN,
    (match, decimal: string | undefined, hex: string | undefined, named: string | undefined) => {
      if (decimal !== undefined) {
        const codePoint = Number.parseInt(decimal, 10);
        return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : match;
      }
      if (hex !== undefined) {
        const codePoint = Number.parseInt(hex, 16);
        return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : match;
      }
      switch (named) {
        case "amp":
          return "&";
        case "lt":
          return "<";
        case "gt":
          return ">";
        case "quot":
          return '"';
        case "apos":
          return "'";
        default:
          return match;
      }
    }
  );
}

/**
 * Text offsets (code units) of `<w:br w:type="page"/>` runs inside the
 * paragraph. Offsets are measured in the same text the model exposes through
 * `paragraph.children`, so they can be mapped back to a wrap line.
 */
export function collectParagraphExplicitPageBreakOffsets(
  paragraph: ParagraphNode
): number[] {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return [];
  }

  const offsets: number[] = [];
  let textOffset = 0;
  for (const match of xml.matchAll(PARAGRAPH_INLINE_TOKEN_PATTERN)) {
    const token = match[0];
    if (match[1] !== undefined) {
      textOffset += decodeXmlText(match[1]).length;
      continue;
    }

    if (/^<w:br\b/i.test(token)) {
      if (/\bw:type="page"/i.test(token)) {
        offsets.push(textOffset);
      } else {
        // A plain <w:br/> is a line break and occupies one text position.
        textOffset += 1;
      }
      continue;
    }

    // <w:cr/>, <w:tab/> and <w:noBreakHyphen/> each render as one character.
    textOffset += 1;
  }

  return offsets;
}

function paragraphModelTextLength(paragraph: ParagraphNode): number {
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

/**
 * True when the paragraph's explicit page break sits at (or after) the end of
 * its visible content, i.e. the break belongs between this paragraph and the
 * next node rather than inside this paragraph.
 */
export function paragraphHasTrailingExplicitPageBreak(paragraph: ParagraphNode): boolean {
  if (!paragraphHasExplicitPageBreak(paragraph)) {
    return false;
  }

  const offsets = collectParagraphExplicitPageBreakOffsets(paragraph);
  if (offsets.length === 0) {
    // The break lives in markup the offset scanner does not model (e.g. inside
    // a field result); keep the historical node-level behaviour.
    return true;
  }

  const textLength = paragraphModelTextLength(paragraph);
  return Math.max(...offsets) >= textLength;
}

/**
 * OOXML does not collapse `w:spacing w:after` into the next paragraph's
 * `w:before`; both apply. Only a paragraph that begins a page or column drops
 * its `w:before`. This converts a raw paragraph height (which already contains
 * both spacing values, see `estimateParagraphHeightPx`) into the height the
 * flow algorithm should consume.
 */
export function resolveParagraphConsumedHeightPx(
  rawNodeHeightPx: number,
  beforeSpacingPx: number,
  atPageTop: boolean
): number {
  const safeRawHeightPx = Math.max(1, rawNodeHeightPx);
  return atPageTop
    ? Math.max(1, safeRawHeightPx - Math.max(0, beforeSpacingPx))
    : safeRawHeightPx;
}

export function paragraphBeforeSpacingPx(paragraph: ParagraphNode): number {
  return twipsToPixels(paragraph.style?.spacing?.beforeTwips) ?? 0;
}

export function paragraphAfterSpacingPx(paragraph: ParagraphNode): number {
  return twipsToPixels(paragraph.style?.spacing?.afterTwips) ?? 0;
}

export function resolveParagraphBeforeSpacingPx(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode,
  pageConsumedHeightPx: number,
  suppressSpacingBeforeAfterPageBreak: boolean
): number {
  const beforeSpacingPx = paragraphBeforeSpacingPx(paragraph);
  if (!suppressSpacingBeforeAfterPageBreak) {
    return beforeSpacingPx;
  }

  if (pageConsumedHeightPx <= 0 && nodeIndex > 0) {
    const previousNode = model.nodes[nodeIndex - 1];
    if (previousNode?.type === "paragraph") {
      if (paragraphIsOnlyExplicitPageBreak(previousNode)) {
        return 0;
      }

      if (paragraphHasPageBreakBefore(previousNode)) {
        return 0;
      }
    }
  }

  return beforeSpacingPx;
}

export function sectionTitlePageEnabled(sectionPropertiesXml?: string): boolean {
  if (!sectionPropertiesXml) {
    return false;
  }

  const titlePageTag = sectionPropertiesXml.match(/<w:titlePg\b[^>]*\/?>/i)?.[0];
  return isOnOffTagEnabled(titlePageTag);
}

export function selectSectionVariantForPage<T extends HeaderSection | FooterSection>(
  sections: T[],
  sectionPropertiesXml: string | undefined,
  pageIndex: number,
  options?: {
    evenAndOddHeaders?: boolean;
    /** Zero-based physical page index used for even/odd selection. */
    documentPageIndex?: number;
  }
): T | undefined {
  if (sections.length === 0) {
    return undefined;
  }

  const titlePage = sectionTitlePageEnabled(sectionPropertiesXml);
  const normalizeType = (value: string | undefined): string => value?.trim().toLowerCase() ?? "";
  const first = sections.find((section) => normalizeType(section.referenceType) === "first");
  const defaultSection = sections.find((section) => {
    const referenceType = normalizeType(section.referenceType);
    return referenceType === "default" || referenceType === "";
  });
  const even = sections.find((section) => normalizeType(section.referenceType) === "even");
  const evenAndOddHeadersEnabled = options?.evenAndOddHeaders ?? true;

  const safePageIndex = Number.isFinite(pageIndex) ? Math.max(0, Math.round(pageIndex)) : 0;
  const physicalPageIndex = Number.isFinite(options?.documentPageIndex)
    ? Math.max(0, Math.round(options?.documentPageIndex as number))
    : safePageIndex;
  const oddPageNumber = physicalPageIndex % 2 === 0;

  if (safePageIndex === 0 && titlePage) {
    return first;
  }

  if (evenAndOddHeadersEnabled && !oddPageNumber && even) {
    return even;
  }

  if (defaultSection) {
    return defaultSection;
  }

  return first ?? even ?? sections[0];
}

export function resolveSectionPropertiesXmlForNodeIndex(
  sections: ResolvedModelSection[],
  nodeIndex: number,
  fallbackSectionPropertiesXml?: string
): string | undefined {
  if (sections.length === 0) {
    return fallbackSectionPropertiesXml;
  }

  const sectionIndex = resolveSectionIndexForNodeIndex(sections, nodeIndex, 0);
  return sections[sectionIndex]?.sectionPropertiesXml ?? fallbackSectionPropertiesXml;
}

export function resolvePaginationSectionMetricsIndexForNodeIndex(
  metricsBySection: PaginationSectionMetrics[],
  nodeIndex: number,
  previousSectionIndex: number
): number {
  if (metricsBySection.length === 0) {
    return 0;
  }

  const safePrevious = Math.max(0, Math.min(previousSectionIndex, metricsBySection.length - 1));
  let sectionIndex = safePrevious;
  if (nodeIndex < metricsBySection[sectionIndex].startNodeIndex) {
    sectionIndex = 0;
  }

  while (
    sectionIndex + 1 < metricsBySection.length &&
    metricsBySection[sectionIndex + 1].startNodeIndex <= nodeIndex
  ) {
    sectionIndex += 1;
  }

  return sectionIndex;
}

export function scalePaginationSectionMetricsHeights(
  metricsBySection: PaginationSectionMetrics[],
  heightScale: number
): PaginationSectionMetrics[] {
  if (!Number.isFinite(heightScale) || Math.abs(heightScale - 1) < 0.001) {
    return metricsBySection;
  }

  return metricsBySection.map((metrics) => ({
    ...metrics,
    pageContentHeightPx: Math.max(120, Math.round(metrics.pageContentHeightPx * heightScale))
  }));
}

export function collectTableExplicitPageBreakInfo(table: TableNode): TableExplicitPageBreakInfo {
  const sourceXml = table.sourceXml ?? "";
  if (sourceXml) {
    const cached = tableExplicitPageBreakInfoByTable.get(table);
    if (cached?.sourceXml === sourceXml) {
      return cached.info;
    }
  }

  const startRowIndexes = new Set<number>();
  let breakAfterTable = false;

  table.rows.forEach((row, rowIndex) => {
    let rowBreakTarget: number | undefined;

    row.cells.forEach((cell) => {
      tableCellDirectParagraphs(cell.nodes).forEach((paragraph) => {
        if (!paragraphHasExplicitPageBreak(paragraph)) {
          return;
        }

        if (
          paragraphIsOnlyExplicitPageBreak(paragraph) &&
          rowIndex === 0 &&
          tableRowHasSubstantiveContentOutsideBreakOnlyParagraphs(row) &&
          !tableRowUsesTrailingSignatureCellBreakPattern(table, row)
        ) {
          return;
        }

        const breakTarget = paragraphIsOnlyExplicitPageBreak(paragraph) ? rowIndex : rowIndex + 1;
        if (rowBreakTarget === undefined || breakTarget < rowBreakTarget) {
          rowBreakTarget = breakTarget;
        }
      });
    });

    if (rowBreakTarget === undefined) {
      return;
    }

    if (rowBreakTarget >= table.rows.length) {
      breakAfterTable = true;
      return;
    }

    startRowIndexes.add(Math.max(0, rowBreakTarget));
  });

  const info: TableExplicitPageBreakInfo = {
    startRowIndexes: [...startRowIndexes].sort((left, right) => left - right),
    breakAfterTable
  };

  if (sourceXml) {
    tableExplicitPageBreakInfoByTable.set(table, { info, sourceXml });
  }

  return info;
}

export function collectTopLevelExplicitPageBreakStartNodeIndexes(
  nodes: DocModel["nodes"]
): Set<number> {
  const breaks = new Set<number>();

  for (let nodeIndex = 0; nodeIndex < nodes.length; nodeIndex += 1) {
    const node = nodes[nodeIndex];
    const hasNextNode = nodeIndex + 1 < nodes.length;

    if (node.type === "paragraph") {
      if (paragraphHasPageBreakBefore(node)) {
        breaks.add(nodeIndex);
      }

      if (hasNextNode && paragraphHasTrailingExplicitPageBreak(node)) {
        breaks.add(nodeIndex + 1);
      }
      continue;
    }

    const tableBreakInfo = collectTableExplicitPageBreakInfo(node);
    if (tableBreakInfo.startRowIndexes.includes(0) && nodeIndex > 0) {
      breaks.add(nodeIndex);
    }

    if (hasNextNode && tableBreakInfo.breakAfterTable) {
      breaks.add(nodeIndex + 1);
    }
  }

  for (const breakIndex of [...breaks]) {
    if (breakIndex <= 0 || breakIndex >= nodes.length) {
      breaks.delete(breakIndex);
    }
  }

  return breaks;
}

export function collectDocxHardPageBreakStartNodeIndexes(model: DocModel): Set<number> {
  const breaks = collectTopLevelExplicitPageBreakStartNodeIndexes(model.nodes);

  const sections = resolveDocumentSectionsFromMetadata(model.metadata);
  for (let sectionIndex = 1; sectionIndex < sections.length; sectionIndex += 1) {
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

  for (const breakIndex of [...breaks]) {
    if (breakIndex <= 0 || breakIndex >= model.nodes.length) {
      breaks.delete(breakIndex);
    }
  }

  return breaks;
}

export function collectDocxLastRenderedPageBreakStartNodeIndexes(model: DocModel): number[] {
  const breaks: number[] = [];

  model.nodes.forEach((node, nodeIndex) => {
    if (
      node.type === "paragraph" &&
      paragraphLastRenderedPageBreakPosition(node) === "leading"
    ) {
      breaks.push(nodeIndex);
    }
  });

  return breaks;
}

export function buildDocumentPageRanges(
  nodeCount: number,
  pageBreakStartNodeIndexes: Iterable<number>
): DocumentPageRange[] {
  if (nodeCount <= 0) {
    return [];
  }

  const sortedBreakStartIndexes = [...pageBreakStartNodeIndexes]
    .filter((index) => index > 0 && index < nodeCount)
    .sort((left, right) => left - right);

  const ranges: DocumentPageRange[] = [];
  let startNodeIndex = 0;

  for (const breakStartIndex of sortedBreakStartIndexes) {
    if (breakStartIndex <= startNodeIndex) {
      continue;
    }

    ranges.push({
      startNodeIndex,
      endNodeIndex: breakStartIndex
    });
    startNodeIndex = breakStartIndex;
  }

  ranges.push({
    startNodeIndex,
    endNodeIndex: nodeCount
  });

  return ranges;
}
