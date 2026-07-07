// Paragraph structural analysis: empty/section-break/carryover spacer
// detection, contextual spacing suppression, bookmark/note reference
// extraction, and section-break anchor carryover classification.
// Upstream editor.tsx: lines 7936-8652.
//
// The explicit page-break / page-break-before / last-rendered-page-break /
// section-break-starts-new-page detectors and the before/after spacing
// helpers live in layout/pagination.ts (already migrated); this module
// re-uses the shared `paragraphBreakFlagsBySourceXml` cache for the
// break-flag read paths and supplies the higher-level structural predicates
// (spacer / carryover / contextual-spacing / bookmark / note-reference)
// that the layout layer consumes through callbacks.

import type {
  DocModel,
  ParagraphNode
} from "../../engine/types";
import type {
  DocumentPageNodeSegment,
  ParagraphLineRange,
  TableRowRange
} from "../../layout/page-segmentation-core";
import { paragraphBreakFlagsBySourceXml, setCacheEntry } from "./cache-utils";
import {
  BOOKMARK_START_XML_PATTERN,
  COLUMN_BREAK_XML_PATTERN,
  ENDNOTE_REFERENCE_XML_PATTERN,
  FOOTNOTE_REFERENCE_XML_PATTERN,
  LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_PX,
  LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_RATIO,
  LAST_RENDERED_PAGE_BREAK_HINT_MIN_REMAINING_SPACE_PX,
  LAST_RENDERED_PAGE_BREAK_XML_PATTERN,
  PAGE_BREAK_BEFORE_XML_PATTERN,
  PAGE_BREAK_XML_PATTERN,
  SECTION_PROPERTIES_XML_PATTERN,
  SECTION_TYPE_XML_PATTERN
} from "./constants";
import { clampNumber } from "./zoom-utils";
import {
  paragraphHasFormField,
  paragraphHasVisibleText,
  paragraphIsAbsoluteFloatingImageAnchorOnly
} from "./paragraph-geometry";
import { paragraphHasImage, tableCellParagraphsRecursively } from "./paragraph-inspect";
import { paragraphLetterheadSideFromIndent } from "./letterhead";
import {
  parseSectionStartType,
  type HeaderFooterDocumentSection
} from "./header-footer";
import { paragraphBorderVisible } from "./table-utils";

// --- Section-break property helpers (re-used via shared cache) ---

function sectionBreakPropertiesStartNewPage(
  sectionPropertiesXml: string
): boolean {
  const sectionType =
    sectionPropertiesXml
      .match(SECTION_TYPE_XML_PATTERN)?.[1]
      ?.trim()
      .toLowerCase() ?? "nextpage";

  if (sectionType === "continuous") {
    return false;
  }

  if (sectionType === "nextcolumn") {
    const columnsTag = sectionPropertiesXml.match(/<w:cols\b[^>]*\/?>/i)?.[0];
    const columnsCount = Number.parseInt(
      columnsTag?.match(/\bw:num="(\d+)"/i)?.[1] ?? "",
      10
    );
    // In single-column sections, "nextColumn" effectively behaves like "nextPage".
    return !Number.isFinite(columnsCount) || columnsCount <= 1;
  }

  return true;
}

export function isOnOffTagEnabled(tagXml: string | undefined): boolean {
  if (!tagXml) {
    return false;
  }

  const valueMatch = tagXml
    .match(/\bw:val="([^"]+)"/i)?.[1]
    ?.trim()
    .toLowerCase();
  if (!valueMatch) {
    return true;
  }

  return !["0", "false", "off", "no"].includes(valueMatch);
}

function paragraphHasExplicitPageBreak(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.explicitPageBreak;
  }

  const flags = {
    explicitPageBreak: PAGE_BREAK_XML_PATTERN.test(xml),
    explicitColumnBreak: COLUMN_BREAK_XML_PATTERN.test(xml),
    lastRenderedPageBreak: LAST_RENDERED_PAGE_BREAK_XML_PATTERN.test(xml),
    pageBreakBefore: isOnOffTagEnabled(
      xml.match(PAGE_BREAK_BEFORE_XML_PATTERN)?.[0]
    ),
    sectionBreakStartsNewPage: (() => {
      const sectionProperties = xml.match(SECTION_PROPERTIES_XML_PATTERN)?.[0];
      if (!sectionProperties) {
        return false;
      }
      // ECMA-376 §2.6.22: omitted <w:type> defaults to nextPage.
      return sectionBreakPropertiesStartNewPage(sectionProperties);
    })(),
  };
  setCacheEntry(paragraphBreakFlagsBySourceXml, xml, flags);
  return flags.explicitPageBreak;
}

export function paragraphHasExplicitColumnBreak(
  paragraph: ParagraphNode
): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.explicitColumnBreak;
  }

  paragraphHasExplicitPageBreak(paragraph);
  return paragraphBreakFlagsBySourceXml.get(xml)?.explicitColumnBreak ?? false;
}

function paragraphHasLastRenderedPageBreak(
  paragraph: ParagraphNode
): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.lastRenderedPageBreak;
  }

  paragraphHasExplicitPageBreak(paragraph);
  return (
    paragraphBreakFlagsBySourceXml.get(xml)?.lastRenderedPageBreak ?? false
  );
}

function paragraphStartsWithLastRenderedPageBreak(
  paragraph: ParagraphNode
): boolean {
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
    .replace(
      /<w:(?:proofErr|bookmarkStart|bookmarkEnd|permStart|permEnd)\b[^>]*\/?>/gi,
      ""
    )
    .replace(/<\/?w:(?:ins|smartTag)\b[^>]*>/gi, "")
    .replace(/\s+/g, "");

  return leadingXml.length === 0;
}

export function shouldHonorParagraphStartLastRenderedPageBreak(params: {
  pageConsumedHeightPx: number;
  pageContentHeightPx: number;
}): boolean {
  const pageConsumedHeightPx = Math.max(
    0,
    Math.round(params.pageConsumedHeightPx)
  );
  const pageContentHeightPx = Math.max(
    0,
    Math.round(params.pageContentHeightPx)
  );
  if (pageConsumedHeightPx <= 0 || pageContentHeightPx <= 0) {
    return false;
  }

  const remainingHeightPx = Math.max(
    0,
    pageContentHeightPx - pageConsumedHeightPx
  );
  const maxAllowedRemainingHeightPx = clampNumber(
    Math.round(
      pageContentHeightPx *
        LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_RATIO
    ),
    LAST_RENDERED_PAGE_BREAK_HINT_MIN_REMAINING_SPACE_PX,
    LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_PX
  );
  return remainingHeightPx <= maxAllowedRemainingHeightPx;
}

function paragraphHasPageBreakBefore(paragraph: ParagraphNode): boolean {
  if (paragraph.style?.pageBreakBefore === true) {
    return true;
  }

  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.pageBreakBefore;
  }

  paragraphHasExplicitPageBreak(paragraph);
  return paragraphBreakFlagsBySourceXml.get(xml)?.pageBreakBefore ?? false;
}

function sectionBreakAfterParagraphStartsNewPage(
  paragraph: ParagraphNode
): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.sectionBreakStartsNewPage;
  }

  // `paragraphHasExplicitPageBreak` populates the shared cache with every
  // break-related flag, including `sectionBreakStartsNewPage`. The cache
  // read below has to happen unconditionally — a section break paragraph
  // is allowed to start a new page (per ECMA-376 §2.6.22) even when it
  // carries no explicit `<w:br w:type="page"/>` run, and an earlier
  // version of this helper returned `false` in exactly that case.
  paragraphHasExplicitPageBreak(paragraph);
  return (
    paragraphBreakFlagsBySourceXml.get(xml)?.sectionBreakStartsNewPage ?? false
  );
}

// --- Structural predicates ---

export function paragraphIsEffectivelyEmpty(
  paragraph: ParagraphNode
): boolean {
  if (paragraphHasImage(paragraph) || paragraphHasFormField(paragraph)) {
    return false;
  }

  return paragraph.children.every(
    (child) => child.type === "text" && child.text.length === 0
  );
}

export function paragraphHasDeletedParagraphMark(
  paragraph: ParagraphNode
): boolean {
  if (paragraph.paragraphMarkDeleted === true) {
    return true;
  }

  const sourceXml = paragraph.sourceXml ?? "";
  return /<w:pPr\b[\s\S]*?<w:rPr\b[\s\S]*?<w:del\b/i.test(sourceXml);
}

export function paragraphCollapsesIntoPreviousParagraph(
  paragraph: ParagraphNode,
  previousNode?: DocModel["nodes"][number]
): boolean {
  return (
    paragraphIsEffectivelyEmpty(paragraph) &&
    previousNode?.type === "paragraph" &&
    paragraphHasDeletedParagraphMark(previousNode)
  );
}

export function paragraphIsStructuralSectionBreakSpacer(
  paragraph: ParagraphNode
): boolean {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml || !SECTION_PROPERTIES_XML_PATTERN.test(sourceXml)) {
    return false;
  }

  if (
    !paragraphIsEffectivelyEmpty(paragraph) ||
    paragraphHasExplicitPageBreak(paragraph) ||
    paragraphHasPageBreakBefore(paragraph) ||
    paragraphStartsWithLastRenderedPageBreak(paragraph)
  ) {
    return false;
  }

  if (paragraphLetterheadSideFromIndent(paragraph)) {
    return false;
  }

  return true;
}

export function paragraphActsAsSectionBreakCarryoverSpacer(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode
): boolean {
  if (
    !paragraphIsEffectivelyEmpty(paragraph) ||
    paragraphHasExplicitPageBreak(paragraph) ||
    paragraphHasPageBreakBefore(paragraph) ||
    paragraphStartsWithLastRenderedPageBreak(paragraph)
  ) {
    return false;
  }

  const nextNode = model.nodes[nodeIndex + 1];
  if (!nextNode || nextNode.type !== "paragraph") {
    return false;
  }

  return paragraphIsSectionBreakAnchorCarryover(nextNode);
}

export function paragraphActsAsTrailingRenderedPageBreakSpacer(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode
): boolean {
  if (
    !paragraphIsEffectivelyEmpty(paragraph) ||
    paragraphHasExplicitPageBreak(paragraph) ||
    paragraphHasPageBreakBefore(paragraph) ||
    paragraphStartsWithLastRenderedPageBreak(paragraph) ||
    paragraphHasImage(paragraph) ||
    paragraphHasFormField(paragraph)
  ) {
    return false;
  }

  let lookaheadIndex = nodeIndex + 1;
  while (lookaheadIndex < model.nodes.length) {
    const nextNode = model.nodes[lookaheadIndex];
    if (nextNode?.type !== "paragraph") {
      return false;
    }
    if (!paragraphIsEffectivelyEmpty(nextNode)) {
      return (
        paragraphHasPageBreakBefore(nextNode) ||
        paragraphStartsWithLastRenderedPageBreak(nextNode)
      );
    }
    if (
      paragraphHasExplicitPageBreak(nextNode) ||
      paragraphHasPageBreakBefore(nextNode) ||
      paragraphStartsWithLastRenderedPageBreak(nextNode) ||
      paragraphHasImage(nextNode) ||
      paragraphHasFormField(nextNode)
    ) {
      return false;
    }
    lookaheadIndex += 1;
  }

  return false;
}

export function paragraphContextualSpacingStyleKey(
  paragraph: ParagraphNode
): string | undefined {
  const styleId = paragraph.style?.styleId?.trim().toLowerCase();
  if (styleId) {
    return `id:${styleId}`;
  }

  const styleName = paragraph.style?.styleName?.trim().toLowerCase();
  if (styleName) {
    return `name:${styleName}`;
  }

  return undefined;
}

export function paragraphsSuppressInterParagraphSpacing(
  previousParagraph: ParagraphNode | undefined,
  currentParagraph: ParagraphNode
): boolean {
  if (!previousParagraph) {
    return false;
  }

  const previousStyleKey =
    paragraphContextualSpacingStyleKey(previousParagraph);
  const currentStyleKey = paragraphContextualSpacingStyleKey(currentParagraph);
  if (!previousStyleKey || previousStyleKey !== currentStyleKey) {
    return false;
  }

  return (
    previousParagraph.style?.contextualSpacing === true ||
    currentParagraph.style?.contextualSpacing === true
  );
}

export function nodeHasSubstantiveContentForPagination(
  node: DocModel["nodes"][number]
): boolean {
  if (node.type === "table") {
    return true;
  }

  if (
    paragraphHasVisibleText(node) ||
    paragraphHasImage(node) ||
    paragraphHasFormField(node)
  ) {
    return true;
  }

  return (
    paragraphHasPageBreakBefore(node) ||
    paragraphHasExplicitPageBreak(node) ||
    sectionBreakAfterParagraphStartsNewPage(node)
  );
}

export function paragraphBookmarkNames(paragraph: ParagraphNode): string[] {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml) {
    return [];
  }

  const names = [
    ...sourceXml.matchAll(new RegExp(BOOKMARK_START_XML_PATTERN.source, "gi")),
  ]
    .map((match) => match[1]?.trim())
    .filter((name): name is string =>
      Boolean(name && name.length > 0 && name !== "_GoBack")
    );
  return [...new Set(names)];
}

export function paragraphReferencedNoteIds(
  paragraph: ParagraphNode,
  noteType: "footnote" | "endnote"
): number[] {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml) {
    return [];
  }

  const pattern =
    noteType === "footnote"
      ? new RegExp(FOOTNOTE_REFERENCE_XML_PATTERN.source, "gi")
      : new RegExp(ENDNOTE_REFERENCE_XML_PATTERN.source, "gi");
  const references: number[] = [];

  for (const match of sourceXml.matchAll(pattern)) {
    const rawId = Number(match[1]);
    if (!Number.isFinite(rawId) || rawId < 0) {
      continue;
    }
    references.push(Math.round(rawId));
  }

  return references;
}

export function nodeReferencedNoteIds(
  node: DocModel["nodes"][number],
  noteType: "footnote" | "endnote",
  tableRowRange?: TableRowRange,
  paragraphLineRange?: ParagraphLineRange
): number[] {
  if (node.type === "paragraph") {
    if (paragraphLineRange && paragraphLineRange.startLineIndex > 0) {
      return [];
    }
    return paragraphReferencedNoteIds(node, noteType);
  }

  const references: number[] = [];
  const startRowIndex = Math.max(0, tableRowRange?.startRowIndex ?? 0);
  const endRowIndex = Math.min(
    node.rows.length,
    tableRowRange?.endRowIndex ?? node.rows.length
  );
  for (let rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex += 1) {
    const row = node.rows[rowIndex];
    row?.cells.forEach((cell) => {
      tableCellParagraphsRecursively(cell.nodes).forEach((paragraph) => {
        references.push(...paragraphReferencedNoteIds(paragraph, noteType));
      });
    });
  }
  return references;
}

export function eventTargetIsInteractiveControl(
  target: EventTarget | null
): boolean {
  if (!(target instanceof Element)) {
    return false;
  }

  return Boolean(
    target.closest("a[href],button,input,select,textarea,[role='checkbox']")
  );
}

export function eventTargetIsNestedTableParagraphEditor(
  target: EventTarget | null
): boolean {
  if (!(target instanceof Element)) {
    return false;
  }

  return Boolean(
    target.closest("[data-docx-table-cell-paragraph-host='true']")
  );
}

function sectionTitlePageEnabled(sectionPropertiesXml?: string): boolean {
  if (!sectionPropertiesXml) {
    return false;
  }

  const titlePageTag = sectionPropertiesXml.match(
    /<w:titlePg\b[^>]*\/?>/i
  )?.[0];
  return isOnOffTagEnabled(titlePageTag);
}

function selectSectionVariantForPage<
  T extends import("../../engine/types").HeaderSection | import("../../engine/types").FooterSection
>(
  sections: T[],
  sectionPropertiesXml: string | undefined,
  pageIndex: number,
  options?: {
    evenAndOddHeaders?: boolean;
  }
): T | undefined {
  if (sections.length === 0) {
    return undefined;
  }

  const titlePage = sectionTitlePageEnabled(sectionPropertiesXml);
  const normalizeType = (value: string | undefined): string =>
    value?.trim().toLowerCase() ?? "";
  const first = sections.find(
    (section) => normalizeType(section.referenceType) === "first"
  );
  const defaultSection = sections.find((section) => {
    const referenceType = normalizeType(section.referenceType);
    return referenceType === "default" || referenceType === "";
  });
  const even = sections.find(
    (section) => normalizeType(section.referenceType) === "even"
  );
  const evenAndOddHeadersEnabled = options?.evenAndOddHeaders ?? true;

  const safePageIndex = Number.isFinite(pageIndex)
    ? Math.max(0, Math.round(pageIndex))
    : 0;
  const oddPageNumber = safePageIndex % 2 === 0;

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

export function paragraphHasVisibleBorder(paragraph: ParagraphNode): boolean {
  return (
    paragraphBorderVisible(paragraph.style?.borders?.top) ||
    paragraphBorderVisible(paragraph.style?.borders?.right) ||
    paragraphBorderVisible(paragraph.style?.borders?.bottom) ||
    paragraphBorderVisible(paragraph.style?.borders?.left) ||
    paragraphBorderVisible(paragraph.style?.borders?.between) ||
    paragraphBorderVisible(paragraph.style?.borders?.bar)
  );
}

export function paragraphIsSectionBreakAnchorCarryover(
  paragraph: ParagraphNode
): boolean {
  if (!paragraph.sourceXml?.includes("<w:sectPr")) {
    return false;
  }

  // Section-break anchor paragraphs can legitimately carry a text-bearing
  // synthetic textbox duplicate used by Word's floating-shape anchoring.
  // Treat these as carryover anchors so they don't consume a standalone page.
  if (!paragraphIsAbsoluteFloatingImageAnchorOnly(paragraph)) {
    return false;
  }

  if (paragraphHasVisibleText(paragraph) || paragraphHasFormField(paragraph)) {
    return false;
  }

  return !paragraphHasVisibleBorder(paragraph);
}

function resolveSectionIndexForNodeIndex(
  sections: HeaderFooterDocumentSection[],
  nodeIndex: number,
  previousSectionIndex: number
): number {
  if (sections.length === 0) {
    return 0;
  }

  const safePrevious = Math.max(
    0,
    Math.min(previousSectionIndex, sections.length - 1)
  );
  let sectionIndex = safePrevious;

  if (nodeIndex < sections[sectionIndex].startNodeIndex) {
    sectionIndex = 0;
  }

  while (
    sectionIndex + 1 < sections.length &&
    sections[sectionIndex + 1].startNodeIndex <= nodeIndex
  ) {
    sectionIndex += 1;
  }

  return sectionIndex;
}

export function resolveSectionIndexForPageSegments(
  sections: HeaderFooterDocumentSection[],
  nodes: DocModel["nodes"],
  pageSegments: DocumentPageNodeSegment[],
  previousSectionIndex: number
): number {
  if (sections.length === 0 || pageSegments.length === 0) {
    return 0;
  }

  const firstNodeIndex = pageSegments[0]?.nodeIndex;
  let sectionIndex = Number.isFinite(firstNodeIndex)
    ? resolveSectionIndexForNodeIndex(
        sections,
        firstNodeIndex as number,
        previousSectionIndex
      )
    : Math.max(0, Math.min(previousSectionIndex, sections.length - 1));
  sectionIndex = Math.max(
    sectionIndex,
    Math.max(0, Math.min(previousSectionIndex, sections.length - 1))
  );

  let walkingSectionIndex = sectionIndex;
  for (const segment of pageSegments) {
    walkingSectionIndex = resolveSectionIndexForNodeIndex(
      sections,
      segment.nodeIndex,
      walkingSectionIndex
    );
    if (walkingSectionIndex <= sectionIndex) {
      continue;
    }

    const nextSection = sections[walkingSectionIndex];
    if (!nextSection) {
      continue;
    }

    if (
      parseSectionStartType(nextSection.sectionPropertiesXml) === "continuous"
    ) {
      continue;
    }

    const candidateNode = nodes[segment.nodeIndex];
    if (
      candidateNode?.type === "paragraph" &&
      paragraphIsSectionBreakAnchorCarryover(candidateNode)
    ) {
      continue;
    }

    sectionIndex = walkingSectionIndex;
  }

  return sectionIndex;
}

export function nodeAlreadyEndsAtExplicitPageBoundary(
  node: DocModel["nodes"][number] | undefined
): boolean {
  if (!node || node.type !== "paragraph") {
    return false;
  }

  return (
    paragraphHasExplicitPageBreak(node) ||
    paragraphHasPageBreakBefore(node) ||
    sectionBreakAfterParagraphStartsNewPage(node)
  );
}
