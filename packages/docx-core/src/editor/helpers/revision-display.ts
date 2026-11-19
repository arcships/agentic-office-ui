// Revision display models for DOCX viewing.
//
// Word's "Display for Review" choice is a view concern: inline markup keeps
// revisions in the text flow, balloons move deleted content onto the review
// pane, and none shows the final document. The editing model must keep every
// revision so export still writes them, so this module derives a read-only
// display model instead of mutating the document.
//
// The display model only removes revision content that is part of
// `paragraph.children`. `w:delText` deletions never reach the model; the run
// renderer draws those as synthetic inline segments when inline display is on.

import type {
  DocModel,
  DocNode,
  DocumentSection,
  FooterSection,
  HeaderSection,
  ParagraphCommentRange,
  ParagraphNode,
  TableNode,
} from "../../engine/types";
import type { DocxRevisionDisplay } from "./editor-types";
import { paragraphText } from "./paragraph-inspect";
import { resolveParagraphTrackedMarkup } from "./tracked-changes";
import { extractBalancedTagRanges } from "./xml-parsing";

export interface HiddenParagraphTextRange {
  /** Inclusive offset in the paragraph's model text. */
  start: number;
  /** Exclusive offset in the paragraph's model text. */
  end: number;
}

function isDeletionLike(kind: string | undefined): boolean {
  return kind === "deletion" || kind === "move-from";
}

const REVISION_TAG_PATTERN = /<w:(?:ins|del|moveFrom|moveTo)\b/i;

/**
 * Fast reject for the common case: a document without revision markup must not
 * pay for a tracked-markup scan on every render-model pass.
 */
function paragraphHasRevisionMarkup(paragraph: ParagraphNode): boolean {
  const sourceXml = paragraph.sourceXml ?? "";
  return sourceXml.length > 0 && REVISION_TAG_PATTERN.test(sourceXml);
}

/**
 * Model children removed by the display model. Deleted text and deleted
 * images/form fields share the same revision mapping; form fields keep their
 * own cursor slot like the renderer.
 */
export function hiddenParagraphChildIndexes(
  paragraph: ParagraphNode,
  display: DocxRevisionDisplay
): Set<number> {
  if (display === "inline") {
    return new Set();
  }

  if (!paragraphHasRevisionMarkup(paragraph)) {
    return new Set();
  }

  const markup = resolveParagraphTrackedMarkup(paragraph);
  if (!markup) {
    return new Set();
  }

  const hidden = new Set<number>();
  let visibleChildIndex = 0;
  paragraph.children.forEach((child, childIndex) => {
    if (child.type === "form-field") {
      return;
    }
    const change = markup.inlineChangeByVisibleChildIndex[visibleChildIndex];
    visibleChildIndex += 1;
    if (isDeletionLike(change?.kind)) {
      hidden.add(childIndex);
    }
  });
  return hidden;
}

/**
 * Character ranges the display model removes from `paragraph.children`.
 *
 * `resolveParagraphTrackedMarkup` already maps revision XML units onto the
 * model children the renderer walks, so a deleted run that the model builder
 * materialised as a normal text child is found here without re-parsing.
 */
export function hiddenParagraphTextRanges(
  paragraph: ParagraphNode,
  display: DocxRevisionDisplay
): HiddenParagraphTextRange[] {
  const hidden = hiddenParagraphChildIndexes(paragraph, display);
  if (hidden.size === 0) {
    return [];
  }

  const ranges: HiddenParagraphTextRange[] = [];
  let textOffset = 0;
  paragraph.children.forEach((child, childIndex) => {
    if (child.type !== "text") {
      return;
    }
    const start = textOffset;
    textOffset += child.text.length;
    if (hidden.has(childIndex) && child.text.length > 0) {
      ranges.push({ start, end: textOffset });
    }
  });
  return ranges;
}

function hiddenLengthBefore(
  ranges: readonly HiddenParagraphTextRange[],
  offset: number
): number {
  let hidden = 0;
  for (const range of ranges) {
    if (offset <= range.start) {
      break;
    }
    hidden += Math.min(offset, range.end) - range.start;
  }
  return hidden;
}

/**
 * Maps a source-model text offset to the offset in the display text.
 *
 * Offsets inside hidden revision text clamp to the start of that hidden range,
 * which is the nearest caret position the display can represent.
 */
export function displayTextOffsetForSourceOffset(
  ranges: readonly HiddenParagraphTextRange[],
  offset: number
): number {
  return Math.max(0, Math.round(offset) - hiddenLengthBefore(ranges, offset));
}

/** Maps a display-text offset back to the source-model text. */
export function sourceTextOffsetForDisplayOffset(
  ranges: readonly HiddenParagraphTextRange[],
  offset: number
): number {
  const target = Math.max(0, Math.round(offset));
  let source = 0;
  let display = 0;

  for (const range of ranges) {
    const visibleLength = range.start - source;
    // A display offset exactly at the hidden boundary belongs to the text
    // after the hidden run: the DOM has no caret inside deleted content, so
    // the display position collapses to the nearest visible position after it.
    if (target < display + visibleLength) {
      return source + (target - display);
    }
    display += visibleLength;
    source = range.end;
  }

  return source + (target - display);
}

function remapCommentRanges(
  ranges: readonly ParagraphCommentRange[] | undefined,
  hiddenRanges: readonly HiddenParagraphTextRange[]
): ParagraphCommentRange[] | undefined {
  if (!ranges || ranges.length === 0 || hiddenRanges.length === 0) {
    return ranges ? [...ranges] : undefined;
  }

  return ranges.map((range) => ({
    commentId: range.commentId,
    start: displayTextOffsetForSourceOffset(hiddenRanges, range.start),
    end: displayTextOffsetForSourceOffset(hiddenRanges, range.end),
  }));
}

/**
 * Removes `w:del` / `w:moveFrom` envelopes from a display-only source XML.
 *
 * The run renderer's tracked-change cache is keyed by `sourceXml`. Without
 * this, a display paragraph that dropped deleted model children would reuse
 * the source shape's cached visible-child indexes and style the wrong run.
 */
function stripDeletedRevisionXml(sourceXml: string): string {
  const ranges = [
    ...extractBalancedTagRanges(sourceXml, "w:del"),
    ...extractBalancedTagRanges(sourceXml, "w:moveFrom"),
  ].sort((left, right) => left.start - right.start);
  if (ranges.length === 0) {
    return sourceXml;
  }

  let result = "";
  let cursor = 0;
  for (const range of ranges) {
    if (range.end <= cursor) {
      continue;
    }
    if (range.start > cursor) {
      result += sourceXml.slice(cursor, range.start);
    }
    cursor = Math.max(cursor, range.end);
  }
  return result + sourceXml.slice(cursor);
}

/**
 * Paragraph as it should be drawn for `display`. Returns the input unchanged
 * when nothing is hidden, so inline display and clean documents share nodes.
 */
export function reviewDisplayParagraph(
  paragraph: ParagraphNode,
  display: DocxRevisionDisplay
): ParagraphNode {
  if (display === "inline") {
    return paragraph;
  }

  const hiddenChildren = hiddenParagraphChildIndexes(paragraph, display);
  const hiddenRanges = hiddenParagraphTextRanges(paragraph, display);
  const hideParagraphMark = paragraph.paragraphMarkDeleted === true;
  if (hiddenChildren.size === 0 && !hideParagraphMark) {
    return paragraph;
  }

  const children: ParagraphNode["children"] = paragraph.children.filter(
    (_child, childIndex) => !hiddenChildren.has(childIndex)
  );

  const next: ParagraphNode = {
    ...paragraph,
    children,
  };
  if (hideParagraphMark) {
    next.paragraphMarkDeleted = undefined;
  }
  if (hiddenChildren.size > 0 && paragraph.sourceXml) {
    next.sourceXml = stripDeletedRevisionXml(paragraph.sourceXml);
  }
  if (paragraph.commentRanges?.length) {
    next.commentRanges = remapCommentRanges(paragraph.commentRanges, hiddenRanges);
  }
  return next;
}

function reviewDisplayNode(
  node: DocNode,
  display: DocxRevisionDisplay
): DocNode {
  if (node.type === "paragraph") {
    return reviewDisplayParagraph(node, display);
  }

  return reviewDisplayTable(node, display);
}

function reviewDisplayTableCell(
  cell: TableNode["rows"][number]["cells"][number],
  display: DocxRevisionDisplay
): TableNode["rows"][number]["cells"][number] {
  let changed = false;
  const nodes = cell.nodes.map((node) => {
    const next = reviewDisplayNode(node, display);
    if (next !== node) {
      changed = true;
    }
    return next;
  });
  return changed ? { ...cell, nodes } : cell;
}

function reviewDisplayTable(
  table: TableNode,
  display: DocxRevisionDisplay
): TableNode {
  let changed = false;
  const rows = table.rows.map((row) => {
    let rowChanged = false;
    const cells = row.cells.map((cell) => {
      const next = reviewDisplayTableCell(cell, display);
      if (next !== cell) {
        rowChanged = true;
      }
      return next;
    });
    if (!rowChanged) {
      return row;
    }
    changed = true;
    return { ...row, cells };
  });
  return changed ? { ...table, rows } : table;
}

function reviewDisplaySection<T extends HeaderSection | FooterSection>(
  section: T,
  display: DocxRevisionDisplay
): T {
  let changed = false;
  const nodes = section.nodes.map((node) => {
    const next = reviewDisplayNode(node, display);
    if (next !== node) {
      changed = true;
    }
    return next;
  });
  return changed ? { ...section, nodes } : section;
}

function reviewDisplaySections<T extends HeaderSection | FooterSection>(
  sections: readonly T[] | undefined,
  display: DocxRevisionDisplay
): T[] | undefined {
  if (!sections) {
    return undefined;
  }
  let changed = false;
  const next = sections.map((section) => {
    const review = reviewDisplaySection(section, display);
    if (review !== section) {
      changed = true;
    }
    return review;
  });
  return changed ? next : (sections as T[]);
}

function reviewDisplayDocumentSections(
  sections: readonly DocumentSection[] | undefined,
  display: DocxRevisionDisplay
): DocumentSection[] | undefined {
  if (!sections) {
    return undefined;
  }

  let changed = false;
  const next = sections.map((section) => {
    const headerSections = reviewDisplaySections(section.headerSections, display);
    const footerSections = reviewDisplaySections(section.footerSections, display);
    if (
      headerSections === section.headerSections &&
      footerSections === section.footerSections
    ) {
      return section;
    }
    changed = true;
    return {
      ...section,
      headerSections: headerSections ?? [],
      footerSections: footerSections ?? [],
    };
  });
  return changed ? next : (sections as DocumentSection[]);
}

/**
 * Read-only view of `model` for a review display mode.
 *
 * - `inline` returns the same model: nothing leaves the text flow.
 * - `balloons` and `none` drop deleted/moved-from runs (and the deleted
 *   paragraph mark) from the displayed body, header and footer stories.
 *
 * `none` and `balloons` share the same display text; only the chrome differs,
 * which the renderer selects from `DocxRevisionDisplay` itself. The input model
 * is never mutated, so export and accept/reject keep every revision.
 */
export function reviewDisplayModel(
  model: DocModel,
  display: DocxRevisionDisplay
): DocModel {
  if (display === "inline") {
    return model;
  }

  let nodesChanged = false;
  const nodes = model.nodes.map((node) => {
    const next = reviewDisplayNode(node, display);
    if (next !== node) {
      nodesChanged = true;
    }
    return next;
  });

  const headerSections = reviewDisplaySections(model.metadata.headerSections, display);
  const footerSections = reviewDisplaySections(model.metadata.footerSections, display);
  const sections = reviewDisplayDocumentSections(model.metadata.sections, display);
  const metadataChanged =
    headerSections !== model.metadata.headerSections ||
    footerSections !== model.metadata.footerSections ||
    sections !== model.metadata.sections;

  if (!nodesChanged && !metadataChanged) {
    return model;
  }

  return {
    ...model,
    nodes: nodesChanged ? nodes : model.nodes,
    metadata: metadataChanged
      ? {
          ...model.metadata,
          headerSections: headerSections ?? [],
          footerSections: footerSections ?? [],
          sections,
        }
      : model.metadata,
  };
}
