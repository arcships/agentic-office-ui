// Per-section page numbering: `w:pgNumType` start value and number format,
// resolved against the document's physical page sequence.
//
// Word numbers a document as one sequence of physical pages and lets every
// section override what that sequence *displays*: `w:pgNumType/@w:start`
// restarts the count, `w:pgNumType/@w:fmt` picks the format (`decimal`,
// `lowerRoman`, `upperRoman`, `lowerLetter`, `upperLetter`, …). A section break
// can also demand that the next section starts on an even or odd page
// (`w:type` = `evenPage`/`oddPage`); Word then inserts a blank page, which
// belongs to the section that ends there — it keeps that section's
// header/footer and continues its numbering, while the new section's first
// content page takes its own start number.
//
// The tracker turns "which section owns each page" plus the sections'
// `w:sectPr` into the values `PAGE`, `NUMPAGES` and `SECTIONPAGES` resolve to,
// into the label a thumbnail rail shows, and into the list of parity blank
// pages, so a page without body content keeps its label instead of reading as a
// pagination failure.
//
// This module deliberately has no module-level mutable state: like
// `sections.ts` it is reachable from the platform-neutral `core.ts` runtime
// graph, where top-level mutable state is not allowed.

import {
  resolveSectionIndexForNodeIndex,
  sectionBreakRequiredPageParity,
  type SectionStartPageParity
} from "./sections";
import type { DocumentPageRange } from "./pagination";
import {
  parseSectionPageNumberFormat,
  parseSectionPageNumberStartOverride
} from "../editor/helpers/header-footer";
import { formatNumberingCounter } from "../editor/helpers/numbering";

/**
 * Page-number formats this module models. The values are `ST_NumberFormat`
 * names with the casing `w:pgNumType/@w:fmt` uses.
 */
export type DocxPageNumberFormat =
  | "decimal"
  | "decimalZero"
  | "lowerRoman"
  | "upperRoman"
  | "lowerLetter"
  | "upperLetter"
  | "none";

const CANONICAL_PAGE_NUMBER_FORMATS: Readonly<
  Record<string, DocxPageNumberFormat>
> = Object.freeze({
  decimal: "decimal",
  decimalzero: "decimalZero",
  lowerroman: "lowerRoman",
  upperroman: "upperRoman",
  lowerletter: "lowerLetter",
  upperletter: "upperLetter",
  none: "none"
});

/**
 * Canonical spelling of a `w:pgNumType/@w:fmt` value. An absent or unrecognised
 * value resolves to `decimal`, Word's default; the raw value stays available on
 * the range as `formatXml` for formats this module does not model.
 */
export function normalizeDocxPageNumberFormat(
  formatXml?: string
): DocxPageNumberFormat {
  const normalized = formatXml?.trim().toLowerCase();
  if (!normalized) return "decimal";
  return CANONICAL_PAGE_NUMBER_FORMATS[normalized] ?? "decimal";
}

/** A section's numbering input: its `w:sectPr` fragment and start node. */
export interface DocxSectionPageNumberSection {
  /** First body node index of the section (`0` when omitted). */
  startNodeIndex?: number;
  sectionPropertiesXml?: string;
}

/**
 * One page of the document's physical page sequence.
 *
 * A page without body content — the blank page a parity break inserts — omits
 * `startNodeIndex`; its section is then inherited from the previous page, the
 * way headers, footers and numbering continue across a section break.
 */
export interface DocxSectionPageNumberPage {
  /** Section owning the page; resolved from `startNodeIndex` when omitted. */
  sectionIndex?: number;
  /** First body node rendered on the page; omitted when the page has none. */
  startNodeIndex?: number;
}

/** A blank page inserted so a section starts on its required page parity. */
export interface DocxSectionParityBlankPage {
  readonly kind: "section-parity-blank-page";
  /** Zero-based index of the blank page in the document's page sequence. */
  readonly pageIndex: number;
  /** Section whose `w:type` demanded the parity. */
  readonly sectionIndex: number;
  /** Parity that section must start on (`evenPage` → `even`). */
  readonly requiredParity: SectionStartPageParity;
  /**
   * Section the blank page belongs to: the one that ends at the break, so the
   * page keeps its header/footer and continues its numbering.
   */
  readonly ownerSectionIndex: number;
}

const PARITY_BLANK_PAGE_KIND = "section-parity-blank-page";

/**
 * Build the descriptor that identifies a parity blank page.
 *
 * Throws on input that cannot describe a page: a page index, section index and
 * parity are all required, because a blank page without them could not be
 * attributed to a header/footer or a numbering range.
 */
export function createSectionParityBlankPage(input: {
  pageIndex: number;
  sectionIndex: number;
  requiredParity: SectionStartPageParity;
  ownerSectionIndex?: number;
}): DocxSectionParityBlankPage {
  const { pageIndex, sectionIndex, requiredParity } = input;
  if (!Number.isInteger(pageIndex) || pageIndex < 0) {
    throw new Error(
      `createSectionParityBlankPage: pageIndex must be a non-negative integer, got ${String(
        pageIndex
      )}`
    );
  }
  if (!Number.isInteger(sectionIndex) || sectionIndex < 0) {
    throw new Error(
      `createSectionParityBlankPage: sectionIndex must be a non-negative integer, got ${String(
        sectionIndex
      )}`
    );
  }
  if (requiredParity !== "even" && requiredParity !== "odd") {
    throw new Error(
      `createSectionParityBlankPage: requiredParity must be "even" or "odd", got ${String(
        requiredParity
      )}`
    );
  }
  const requestedOwner = input.ownerSectionIndex;
  const ownerSectionIndex =
    typeof requestedOwner === "number" &&
    Number.isInteger(requestedOwner) &&
    requestedOwner >= 0
      ? requestedOwner
      : sectionIndex;
  return Object.freeze({
    kind: PARITY_BLANK_PAGE_KIND,
    pageIndex,
    sectionIndex,
    requiredParity,
    ownerSectionIndex
  });
}

/** True for a value produced by `createSectionParityBlankPage()`. */
export function isSectionParityBlankPage(
  value: unknown
): value is DocxSectionParityBlankPage {
  if (value === null || typeof value !== "object") return false;
  const candidate = value as Partial<DocxSectionParityBlankPage>;
  return (
    candidate.kind === PARITY_BLANK_PAGE_KIND &&
    Number.isInteger(candidate.pageIndex) &&
    (candidate.pageIndex as number) >= 0 &&
    Number.isInteger(candidate.sectionIndex) &&
    (candidate.sectionIndex as number) >= 0 &&
    (candidate.requiredParity === "even" || candidate.requiredParity === "odd")
  );
}

/**
 * One section's slice of the page sequence.
 *
 * `startPageIndex`/`endPageIndex` are document page indexes with an exclusive
 * end, `startNumber`/`endNumber` are the numbers Word displays on the first and
 * last page of the slice, and `nodeRange` is the body node span it covers.
 */
export interface DocxSectionPageRange {
  sectionIndex: number;
  startPageIndex: number;
  endPageIndex: number;
  nodeRange: DocumentPageRange;
  startNumber: number;
  endNumber: number;
  format: DocxPageNumberFormat;
  /** Raw `w:pgNumType/@w:fmt` value, when the section sets one. */
  formatXml?: string;
  /** True when the section restarts numbering with `w:pgNumType/@w:start`. */
  restarts: boolean;
  /** Parity the section must start on, when its `w:type` demands one. */
  requiredParity?: SectionStartPageParity;
  /** Blank pages at the end of this range inserted for the next section. */
  parityBlankPageIndexes: readonly number[];
}

/** The field values a header/footer resolves for one page. */
export interface DocxSectionPageFields {
  /** `PAGE`: the number Word displays on the page. */
  pageNumber: number;
  /** `w:pgNumType/@w:fmt` of the page's section, when it sets one. */
  pageNumberFormat?: string;
  /** `SECTIONPAGES`: pages in the page's section. */
  sectionPages: number;
  /** `NUMPAGES`: pages in the document. */
  totalPages: number;
  /** Label the page shows, e.g. `iv` (`w:fmt="none"` → empty string). */
  pageLabel: string;
}

export interface DocxSectionPageNumberTracker {
  /** Number of pages in the document (`NUMPAGES`). */
  readonly pageCount: number;
  /** One entry per section run, in page order. */
  readonly ranges: readonly DocxSectionPageRange[];
  /** Blank pages the parity of a section break inserted. */
  readonly parityBlankPages: readonly DocxSectionParityBlankPage[];
  sectionIndexForPage(pageIndex: number): number | undefined;
  pageRangeForPage(pageIndex: number): DocxSectionPageRange | undefined;
  /** `PAGE` value for the page. */
  pageNumberForPage(pageIndex: number): number | undefined;
  /** `w:pgNumType/@w:fmt` of the page's section, when it sets one. */
  pageNumberFormatForPage(pageIndex: number): string | undefined;
  /** Label the page shows, formatted the way the `PAGE` field is. */
  pageLabelForPage(pageIndex: number): string | undefined;
  /** `SECTIONPAGES` value for the page. */
  sectionPagesForPage(pageIndex: number): number | undefined;
  fieldsForPage(pageIndex: number): DocxSectionPageFields | undefined;
  isParityBlankPage(pageIndex: number): boolean;
}

export interface CreateSectionPageNumberTrackerInput {
  /** Section `w:sectPr` fragments in document order. */
  sections: readonly DocxSectionPageNumberSection[];
  /** Pages in document order. */
  pages: readonly DocxSectionPageNumberPage[];
  /**
   * Body node count. Closes the last range's `nodeRange`; omitted, that range
   * ends at the first node index after its last content page.
   */
  nodeCount?: number;
}

interface MutableRange {
  sectionIndex: number;
  startPageIndex: number;
  endPageIndex: number;
  startNumber: number;
  endNumber: number;
  format: DocxPageNumberFormat;
  formatXml?: string;
  restarts: boolean;
  requiredParity?: SectionStartPageParity;
  parityBlankPageIndexes: number[];
}

function normalizedStartNodeIndex(startNodeIndex: unknown): number {
  if (typeof startNodeIndex !== "number" || !Number.isFinite(startNodeIndex)) return 0;
  return Math.max(0, Math.round(startNodeIndex));
}

function normalizedPageNodeIndex(startNodeIndex: unknown): number | undefined {
  if (typeof startNodeIndex !== "number" || !Number.isFinite(startNodeIndex)) {
    return undefined;
  }
  const rounded = Math.round(startNodeIndex);
  return rounded >= 0 ? rounded : undefined;
}

/**
 * Resolve the section owning every page.
 *
 * A page without body content inherits the previous page's section, which is
 * what both pagination engines and the viewers already assume: the blank page a
 * parity break inserts still belongs to the section that ends there.
 */
function resolvePageSectionIndexes(
  pages: readonly DocxSectionPageNumberPage[],
  sections: readonly DocxSectionPageNumberSection[]
): number[] {
  const sectionStarts = sections.map((section) => ({
    startNodeIndex: normalizedStartNodeIndex(section.startNodeIndex)
  }));
  const lastSectionIndex = Math.max(0, sections.length - 1);
  const sectionIndexes: number[] = [];
  let previousSectionIndex = 0;

  for (const page of pages) {
    const explicit =
      typeof page.sectionIndex === "number" && Number.isInteger(page.sectionIndex)
        ? page.sectionIndex
        : undefined;
    if (explicit !== undefined && explicit >= 0) {
      previousSectionIndex = Math.min(explicit, lastSectionIndex);
    } else {
      const startNodeIndex = normalizedPageNodeIndex(page.startNodeIndex);
      if (startNodeIndex !== undefined && sectionStarts.length > 0) {
        previousSectionIndex = resolveSectionIndexForNodeIndex(
          sectionStarts,
          startNodeIndex,
          previousSectionIndex
        );
      }
    }
    sectionIndexes.push(previousSectionIndex);
  }

  return sectionIndexes;
}

/**
 * Build the per-section page-number tracker.
 *
 * `pages` must be in document order. Sections without explicit `w:start`
 * continue the previous section's numbering, which is what Word does when a
 * section break keeps "continue from previous section" numbering.
 */
export function createSectionPageNumberTracker(
  input: CreateSectionPageNumberTrackerInput
): DocxSectionPageNumberTracker {
  const sections = input.sections ?? [];
  const pages = input.pages ?? [];
  const sectionIndexes = resolvePageSectionIndexes(pages, sections);
  const pageNodeIndexes = pages.map((page) => normalizedPageNodeIndex(page.startNodeIndex));

  const sectionNumbering = sections.map((section) => {
    const formatXml = parseSectionPageNumberFormat(section.sectionPropertiesXml);
    return {
      formatXml,
      format: normalizeDocxPageNumberFormat(formatXml),
      startOverride: parseSectionPageNumberStartOverride(section.sectionPropertiesXml),
      requiredParity: sectionBreakRequiredPageParity(section.sectionPropertiesXml)
    };
  });

  const ranges: MutableRange[] = [];
  const numbersByPage: number[] = [];
  const rangeIndexByPage: number[] = [];
  let previousEndNumber = 0;

  for (let pageIndex = 0; pageIndex < sectionIndexes.length; pageIndex += 1) {
    const sectionIndex = sectionIndexes[pageIndex];
    let range = ranges[ranges.length - 1];
    if (!range || range.sectionIndex !== sectionIndex) {
      const numbering = sectionNumbering[sectionIndex];
      const startNumber = numbering?.startOverride ?? previousEndNumber + 1;
      range = {
        sectionIndex,
        startPageIndex: pageIndex,
        endPageIndex: pageIndex + 1,
        startNumber,
        endNumber: startNumber,
        format: numbering?.format ?? "decimal",
        formatXml: numbering?.formatXml,
        restarts: numbering?.startOverride !== undefined,
        requiredParity: numbering?.requiredParity,
        parityBlankPageIndexes: []
      };
      ranges.push(range);
    }

    const pageNumber = range.startNumber + (pageIndex - range.startPageIndex);
    range.endPageIndex = pageIndex + 1;
    range.endNumber = pageNumber;
    numbersByPage[pageIndex] = pageNumber;
    rangeIndexByPage[pageIndex] = ranges.length - 1;
    previousEndNumber = pageNumber;
  }

  // Parity blank pages: a page without body content directly before a section
  // that must start on an even/odd page, whose first content page already has
  // the required parity (the reason the blank page exists at all).
  const parityBlankPages: DocxSectionParityBlankPage[] = [];
  const parityBlankPageIndexes = new Set<number>();
  for (let pageIndex = 0; pageIndex + 1 < sectionIndexes.length; pageIndex += 1) {
    if (pageNodeIndexes[pageIndex] !== undefined) continue;
    const nextSectionIndex = sectionIndexes[pageIndex + 1];
    if (nextSectionIndex === sectionIndexes[pageIndex]) continue;
    const requiredParity = sectionNumbering[nextSectionIndex]?.requiredParity;
    if (!requiredParity) continue;
    // 1-based number of the section's first content page.
    const contentPageNumber = pageIndex + 2;
    const contentStartsOnEvenPage = contentPageNumber % 2 === 0;
    if (contentStartsOnEvenPage !== (requiredParity === "even")) continue;

    const blankPage = createSectionParityBlankPage({
      pageIndex,
      sectionIndex: nextSectionIndex,
      requiredParity,
      ownerSectionIndex: sectionIndexes[pageIndex]
    });
    parityBlankPages.push(blankPage);
    parityBlankPageIndexes.add(pageIndex);
    ranges[rangeIndexByPage[pageIndex]]?.parityBlankPageIndexes.push(pageIndex);
  }

  const nextDefinedNodeIndex = (fromPageIndex: number): number | undefined => {
    for (let index = Math.max(0, fromPageIndex); index < pageNodeIndexes.length; index += 1) {
      const nodeIndex = pageNodeIndexes[index];
      if (nodeIndex !== undefined) return nodeIndex;
    }
    return undefined;
  };
  const lastDefinedNodeIndex = (fromPageIndex: number, toPageIndex: number): number | undefined => {
    for (let index = Math.min(toPageIndex, pageNodeIndexes.length) - 1; index >= fromPageIndex; index -= 1) {
      const nodeIndex = pageNodeIndexes[index];
      if (nodeIndex !== undefined) return nodeIndex;
    }
    return undefined;
  };
  const nodeCount =
    typeof input.nodeCount === "number" && Number.isFinite(input.nodeCount)
      ? Math.max(0, Math.round(input.nodeCount))
      : undefined;

  const frozenRanges: DocxSectionPageRange[] = ranges.map((range, rangeIndex) => {
    const startNodeIndex = nextDefinedNodeIndex(range.startPageIndex) ?? 0;
    const nextRangeNodeIndex = nextDefinedNodeIndex(range.endPageIndex);
    const ownLastNodeIndex = lastDefinedNodeIndex(
      range.startPageIndex,
      range.endPageIndex
    );
    const isLastRange = rangeIndex === ranges.length - 1;
    const endNodeIndex =
      nextRangeNodeIndex ??
      (isLastRange && nodeCount !== undefined
        ? nodeCount
        : (ownLastNodeIndex ?? startNodeIndex) + (ownLastNodeIndex === undefined ? 0 : 1));

    return Object.freeze({
      sectionIndex: range.sectionIndex,
      startPageIndex: range.startPageIndex,
      endPageIndex: range.endPageIndex,
      nodeRange: Object.freeze({
        startNodeIndex,
        endNodeIndex: Math.max(startNodeIndex, endNodeIndex)
      }),
      startNumber: range.startNumber,
      endNumber: range.endNumber,
      format: range.format,
      formatXml: range.formatXml,
      restarts: range.restarts,
      requiredParity: range.requiredParity,
      parityBlankPageIndexes: Object.freeze([...range.parityBlankPageIndexes])
    });
  });

  const sectionPageCounts = new Map<number, number>();
  for (const sectionIndex of sectionIndexes) {
    sectionPageCounts.set(sectionIndex, (sectionPageCounts.get(sectionIndex) ?? 0) + 1);
  }

  const rangeAt = (pageIndex: number): DocxSectionPageRange | undefined => {
    if (!Number.isInteger(pageIndex) || pageIndex < 0) return undefined;
    if (pageIndex >= rangeIndexByPage.length) return undefined;
    return frozenRanges[rangeIndexByPage[pageIndex]];
  };

  const pageLabelAt = (pageIndex: number): string | undefined => {
    const range = rangeAt(pageIndex);
    const pageNumber = numbersByPage[pageIndex];
    if (!range || pageNumber === undefined) return undefined;
    return formatNumberingCounter(range.formatXml ?? range.format, pageNumber);
  };

  return {
    pageCount: numbersByPage.length,
    ranges: Object.freeze(frozenRanges),
    parityBlankPages: Object.freeze(parityBlankPages),
    sectionIndexForPage: (pageIndex: number) => rangeAt(pageIndex)?.sectionIndex,
    pageRangeForPage: rangeAt,
    pageNumberForPage: (pageIndex: number) =>
      Number.isInteger(pageIndex) && pageIndex >= 0
        ? numbersByPage[pageIndex]
        : undefined,
    pageNumberFormatForPage: (pageIndex: number) => rangeAt(pageIndex)?.formatXml,
    pageLabelForPage: pageLabelAt,
    sectionPagesForPage: (pageIndex: number) => {
      const sectionIndex = rangeAt(pageIndex)?.sectionIndex;
      return sectionIndex === undefined ? undefined : sectionPageCounts.get(sectionIndex);
    },
    fieldsForPage: (pageIndex: number): DocxSectionPageFields | undefined => {
      const range = rangeAt(pageIndex);
      const pageNumber = numbersByPage[pageIndex];
      if (!range || pageNumber === undefined) return undefined;
      return {
        pageNumber,
        pageNumberFormat: range.formatXml,
        sectionPages: sectionPageCounts.get(range.sectionIndex) ?? 0,
        totalPages: numbersByPage.length,
        pageLabel: pageLabelAt(pageIndex) ?? String(pageNumber)
      };
    },
    isParityBlankPage: (pageIndex: number) =>
      Number.isInteger(pageIndex) && parityBlankPageIndexes.has(pageIndex)
  };
}
