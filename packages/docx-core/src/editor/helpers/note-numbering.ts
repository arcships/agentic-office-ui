// Note numbering and page placement.
//
// Footnote/endnote display numbers must not come from w:id (ids are stable
// identifiers, not ordinals: after a few edits they read 3, 7, 8 …) and the
// number printed next to a note at the page bottom must equal the number of
// its in-text marker. Both are derived here, once, from the run model:
//
//   - document order of the *references* decides the ordinal
//   - w:footnotePr / w:endnotePr decide numFmt, numRestart, numStart, and the
//     endnote position (document end vs. section end)
//   - a page owns the notes whose first reference sits on that page; endnotes
//     are moved to the last page of their section (pos=sectEnd) or of the
//     document (pos=docEnd, the default)

import type {
  DocModel,
  DocNode,
  DocumentNoteDefinition,
  DocumentNoteNumberingProperties,
  DocumentNoteProperties,
  NoteReferenceKind,
  NumberingDefinitionSet,
} from "../../engine/types";
import type { DocumentPageNodeSegment } from "../../layout/page-segmentation-core";
import {
  resolveDocumentSectionsFromMetadata,
  resolveSectionIndexForNodeIndex,
} from "../../layout/pagination";
import {
  mergeNoteNumberingProperties,
  parseNoteNumberingProperties,
} from "../../engine/note-properties";
import { formatNumberingCounter } from "./numbering";
import { nodeNoteReferences } from "./note-references";

export const DEFAULT_FOOTNOTE_NUMBER_FORMAT = "decimal";
export const DEFAULT_ENDNOTE_NUMBER_FORMAT = "lowerRoman";

export interface DocumentNoteEntry {
  kind: NoteReferenceKind;
  note: DocumentNoteDefinition;
  /** 1-based display number inside its numbering run. */
  ordinal: number;
  /** Formatted marker label, e.g. "1", "iv", "c". */
  label: string;
}

export interface DocumentPageNotes {
  footnotes: DocumentNoteEntry[];
  endnotes: DocumentNoteEntry[];
  /** Note id → display ordinal for the text runs printed on this page. */
  footnoteNumbers: Map<number, number>;
  endnoteNumbers: Map<number, number>;
  /** Note id → formatted marker label for the text runs printed on this page. */
  footnoteLabels: Map<number, string>;
  endnoteLabels: Map<number, string>;
}

export interface DocumentNoteNumbering {
  footnote: Map<number, number>;
  endnote: Map<number, number>;
}

function emptyPageNotes(): DocumentPageNotes {
  return {
    footnotes: [],
    endnotes: [],
    footnoteNumbers: new Map<number, number>(),
    endnoteNumbers: new Map<number, number>(),
    footnoteLabels: new Map<number, string>(),
    endnoteLabels: new Map<number, string>(),
  };
}

/** Marker/note label for one ordinal, honouring w:numFmt. */
export function formatNoteMarkerLabel(
  kind: NoteReferenceKind,
  ordinal: number,
  numFmt?: string
): string {
  const format =
    numFmt ??
    (kind === "footnote"
      ? DEFAULT_FOOTNOTE_NUMBER_FORMAT
      : DEFAULT_ENDNOTE_NUMBER_FORMAT);
  return formatNumberingCounter(format, ordinal);
}

function noteKindProperty(
  properties: DocumentNoteProperties | undefined,
  kind: NoteReferenceKind
): DocumentNoteNumberingProperties | undefined {
  return kind === "footnote" ? properties?.footnote : properties?.endnote;
}

interface NoteNumberingContext {
  /** Document-wide defaults merged with the section override. */
  propertiesForSection(sectionIndex: number): DocumentNoteProperties;
}

function createNumberingContext(model: DocModel): NoteNumberingContext {
  const sections = resolveDocumentSectionsFromMetadata(model.metadata);
  const cache = new Map<number, DocumentNoteProperties>();
  const documentProperties = model.metadata.noteProperties;

  return {
    propertiesForSection(sectionIndex: number): DocumentNoteProperties {
      const cached = cache.get(sectionIndex);
      if (cached) {
        return cached;
      }
      const sectionXml = sections[sectionIndex]?.sectionPropertiesXml;
      const resolved: DocumentNoteProperties = {
        footnote: mergeNoteNumberingProperties(
          documentProperties?.footnote,
          parseNoteNumberingProperties(sectionXml, "footnote")
        ),
        endnote: mergeNoteNumberingProperties(
          documentProperties?.endnote,
          parseNoteNumberingProperties(sectionXml, "endnote")
        ),
      };
      cache.set(sectionIndex, resolved);
      return resolved;
    },
  };
}

interface NoteOccurrence {
  kind: NoteReferenceKind;
  id: number;
  pageIndex: number;
  sectionIndex: number;
  ordinal: number;
  label: string;
}

function restartScope(
  properties: DocumentNoteNumberingProperties | undefined,
  pageIndex: number,
  sectionIndex: number
): string {
  const restart = properties?.numRestart ?? "continuous";
  if (restart === "eachSect") {
    return `section:${sectionIndex}`;
  }
  if (restart === "eachPage") {
    return `page:${pageIndex}`;
  }
  return "document";
}

/** Last page carrying content for every section, in page order. */
function lastPageBySection(
  sections: Array<{ startNodeIndex: number }>,
  pages: DocumentPageNodeSegment[][]
): number[] {
  const resolved: number[] = new Array(sections.length).fill(-1);
  let walkingSectionIndex = 0;
  pages.forEach((segments, pageIndex) => {
    for (const segment of segments) {
      walkingSectionIndex = resolveSectionIndexForNodeIndex(
        sections,
        segment.nodeIndex,
        walkingSectionIndex
      );
      resolved[walkingSectionIndex] = pageIndex;
    }
  });
  return resolved;
}

/**
 * Resolve every note reference of the document to a display ordinal and to the
 * page that shows its body, then format the marker labels.
 *
 * `pages` must list the document nodes in reading order; each page is a list of
 * node segments (the same shape the page surfaces render).
 */
export function buildDocumentPageNotes(
  model: DocModel,
  pages: DocumentPageNodeSegment[][]
): DocumentPageNotes[] {
  const pageNotes = pages.map(() => emptyPageNotes());
  if (pageNotes.length === 0) {
    return pageNotes;
  }

  const numbering = createNumberingContext(model);
  const sections = resolveDocumentSectionsFromMetadata(model.metadata);
  const lastPagePerSection = lastPageBySection(sections, pages);

  const footnotesById = new Map(
    (model.metadata.footnotes ?? []).map((note) => [note.id, note])
  );
  const endnotesById = new Map(
    (model.metadata.endnotes ?? []).map((note) => [note.id, note])
  );

  const counters: Record<NoteReferenceKind, Map<string, number>> = {
    footnote: new Map<string, number>(),
    endnote: new Map<string, number>(),
  };
  const emitted = new Set<string>();
  const occurrences: NoteOccurrence[] = [];

  let walkingSectionIndex = 0;
  pages.forEach((segments, pageIndex) => {
    for (const segment of segments) {
      const node = model.nodes[segment.nodeIndex];
      if (!node) {
        continue;
      }
      walkingSectionIndex = resolveSectionIndexForNodeIndex(
        sections,
        segment.nodeIndex,
        walkingSectionIndex
      );
      const references = nodeNoteReferences(
        node,
        undefined,
        segment.tableRowRange,
        segment.paragraphLineRange
      );
      if (references.length === 0) {
        continue;
      }

      const sectionProperties = numbering.propertiesForSection(walkingSectionIndex);

      for (const reference of references) {
        const properties = noteKindProperty(sectionProperties, reference.kind);
        const scope = restartScope(properties, pageIndex, walkingSectionIndex);
        const kindCounters = counters[reference.kind];

        if (!kindCounters.has(scope)) {
          // Every numbering run (document, section, or page) starts at w:numStart.
          kindCounters.set(scope, properties?.numStart ?? 1);
        }

        const ordinal = kindCounters.get(scope)!;
        kindCounters.set(scope, ordinal + 1);
        const label = formatNoteMarkerLabel(
          reference.kind,
          ordinal,
          properties?.numFmt
        );

        const numbers =
          reference.kind === "footnote"
            ? pageNotes[pageIndex].footnoteNumbers
            : pageNotes[pageIndex].endnoteNumbers;
        const labels =
          reference.kind === "footnote"
            ? pageNotes[pageIndex].footnoteLabels
            : pageNotes[pageIndex].endnoteLabels;
        if (!numbers.has(reference.id)) {
          numbers.set(reference.id, ordinal);
        }
        if (!labels.has(reference.id)) {
          labels.set(reference.id, label);
        }

        const key = `${reference.kind}:${reference.id}`;
        if (emitted.has(key)) {
          continue;
        }
        emitted.add(key);
        occurrences.push({
          kind: reference.kind,
          id: reference.id,
          pageIndex,
          sectionIndex: walkingSectionIndex,
          ordinal,
          label,
        });
      }
    }
  });

  const lastPageIndex = pageNotes.length - 1;
  occurrences.forEach((occurrence) => {
    const definition =
      occurrence.kind === "footnote"
        ? footnotesById.get(occurrence.id)
        : endnotesById.get(occurrence.id);
    if (!definition) {
      return;
    }

    let targetPageIndex = occurrence.pageIndex;
    if (occurrence.kind === "endnote") {
      const position =
        noteKindProperty(
          numbering.propertiesForSection(occurrence.sectionIndex),
          "endnote"
        )?.position ?? "docEnd";
      if (position === "sectEnd") {
        const sectionLastPage = lastPagePerSection[occurrence.sectionIndex];
        targetPageIndex = sectionLastPage >= 0 ? sectionLastPage : lastPageIndex;
      } else {
        targetPageIndex = lastPageIndex;
      }
    }

    const entry: DocumentNoteEntry = {
      kind: occurrence.kind,
      note: definition,
      ordinal: occurrence.ordinal,
      label: occurrence.label,
    };
    if (occurrence.kind === "footnote") {
      pageNotes[targetPageIndex].footnotes.push(entry);
    } else {
      pageNotes[targetPageIndex].endnotes.push(entry);
    }

    // The page that prints the note body needs the same label the marker uses
    // (endnotes placed at the document end live on a page that has no marker).
    const bodyLabels =
      occurrence.kind === "footnote"
        ? pageNotes[targetPageIndex].footnoteLabels
        : pageNotes[targetPageIndex].endnoteLabels;
    if (!bodyLabels.has(occurrence.id)) {
      bodyLabels.set(occurrence.id, occurrence.label);
    }
  });

  pageNotes.forEach((notes) => {
    notes.footnotes.sort((left, right) => left.ordinal - right.ordinal);
    notes.endnotes.sort((left, right) => left.ordinal - right.ordinal);
  });

  return pageNotes;
}

/**
 * Document-wide display ordinal per note id, ignoring page-driven restarts.
 * Useful for consumers that render markers without pagination.
 */
export function buildDocumentNoteNumbering(model: DocModel): DocumentNoteNumbering {
  const segments = model.nodes.map((_node, nodeIndex) => ({ nodeIndex }));
  const [pageNotes] = buildDocumentPageNotes(model, [segments]);
  return {
    footnote: new Map(pageNotes?.footnoteNumbers ?? []),
    endnote: new Map(pageNotes?.endnoteNumbers ?? []),
  };
}

// ── Note-height reserve ────────────────────────────────────────────
//
// Notes are printed at the bottom of the page that owns them, so the body has
// less room there. The reserve below is wired into `buildLayoutSnapshot`
// (`paginateDocumentLayout` iterates with per-page content-height overrides)
// so the last page no longer overflows when a footnote carries a tall table.

export type NoteDocNodeHeightEstimator = (
  node: DocNode,
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
) => number;

export interface NoteReserveEstimateOptions {
  estimateDocNodeHeightPx: NoteDocNodeHeightEstimator;
  numberingDefinitions?: NumberingDefinitionSet;
  docGridLinePitchPx?: number;
}

/** Top padding + separator rule of `.docx-notes-section` (see DocxNotesSection). */
const NOTE_SECTION_CHROME_PX = 19;
/** Gap between two note bodies (`margin: 3px 0` twice). */
const NOTE_ENTRY_GAP_PX = 6;
/** Extra gap between two blocks (paragraph/table) inside one note body. */
const NOTE_BLOCK_GAP_PX = 4;
/** `.docx-notes-section--endnote { margin-top: 18px }` when both kinds share a page. */
const ENDNOTE_SECTION_GAP_PX = 18;
/** Fallback body height for a note without nodes (plain `note.text`). */
const NOTE_TEXT_FALLBACK_HEIGHT_PX = 16;

/**
 * Height the rendered note bodies of one page take away from its content box.
 *
 * Counts paragraphs *and* tables: a footnote table (legal/technical docs) used
 * to be ignored, which let the last page overflow.
 */
export function estimateDocumentNoteSectionHeightPx(
  notes: readonly DocumentNoteDefinition[],
  pageContentWidthPx: number,
  options: NoteReserveEstimateOptions
): number {
  if (notes.length === 0) {
    return 0;
  }

  const contentWidthPx = Math.max(1, Math.round(pageContentWidthPx));
  let totalHeightPx = NOTE_SECTION_CHROME_PX;

  notes.forEach((note, noteIndex) => {
    if (noteIndex > 0) {
      totalHeightPx += NOTE_ENTRY_GAP_PX;
    }

    const nodes = note.nodes ?? [];
    if (nodes.length === 0) {
      totalHeightPx += NOTE_TEXT_FALLBACK_HEIGHT_PX;
      return;
    }

    nodes.forEach((node, nodeIndex) => {
      if (nodeIndex > 0) {
        totalHeightPx += NOTE_BLOCK_GAP_PX;
      }
      const heightPx = options.estimateDocNodeHeightPx(
        node,
        contentWidthPx,
        options.numberingDefinitions,
        options.docGridLinePitchPx
      );
      totalHeightPx += Math.max(1, Math.round(heightPx));
    });
  });

  return Math.max(0, Math.round(totalHeightPx));
}

export interface PageNoteReserveOptions extends NoteReserveEstimateOptions {
  /** Content-box width of a page, used to wrap the note bodies it prints. */
  resolvePageContentWidthPx?: (pageIndex: number) => number;
  /** Fallback content width when `resolvePageContentWidthPx` is absent. */
  pageContentWidthPx?: number;
}

/**
 * Reserve (in px) each page must leave for the notes printed at its bottom.
 * Returns one entry per page; `0` means the page prints no notes.
 */
export function resolvePageNoteReservePx(
  model: DocModel,
  pages: readonly DocumentPageNodeSegment[][],
  options: PageNoteReserveOptions
): number[] {
  const pageNotes = buildDocumentPageNotes(model, pages as DocumentPageNodeSegment[][]);
  const fallbackWidthPx = Math.max(1, Math.round(options.pageContentWidthPx ?? 1));

  return pageNotes.map((notes, pageIndex) => {
    if (notes.footnotes.length === 0 && notes.endnotes.length === 0) {
      return 0;
    }
    const widthPx = Math.max(
      1,
      Math.round(
        options.resolvePageContentWidthPx?.(pageIndex) ?? fallbackWidthPx
      )
    );
    const footnoteHeightPx = estimateDocumentNoteSectionHeightPx(
      notes.footnotes.map((entry) => entry.note),
      widthPx,
      options
    );
    const endnoteHeightPx = estimateDocumentNoteSectionHeightPx(
      notes.endnotes.map((entry) => entry.note),
      widthPx,
      options
    );
    if (footnoteHeightPx > 0 && endnoteHeightPx > 0) {
      return footnoteHeightPx + ENDNOTE_SECTION_GAP_PX + endnoteHeightPx;
    }
    return Math.max(footnoteHeightPx, endnoteHeightPx);
  });
}
