// Re-emit comment range markers when a model is serialized.
//
// The WASM serializer writes a paragraph's `sourceXml` verbatim (which already
// contains `w:commentRangeStart`/`w:commentRangeEnd`) and otherwise renders it
// from model fields, which have no slot for those markers. Comment ranges now
// live on the model as character offsets, so after an edit clears `sourceXml`
// we patch the serialized body: walk the paragraph's generated text runs, split
// them at the stored offsets, and put the markers between the halves.

import type { DocModel, DocNode, ParagraphNode, TableNode } from "./types";
import { resolveParagraphCommentRanges } from "./comment-ranges";
import {
  applyDocumentXmlEdits,
  collectRunTextAtoms,
  ensurePreserveSpace,
  escapeXmlText,
  extractRunProperties,
  locateBodyParagraphSpans,
  type DocumentXmlEdit,
  type XmlTextAtom,
} from "./document-xml-patch";

interface SerializedCommentRange {
  commentId: number;
  start: number;
  end: number;
}

interface ParagraphCommentTarget {
  /** Index among every `<w:p>` in the part, in document order (cells included). */
  paragraphIndex: number;
  ranges: SerializedCommentRange[];
}

function paragraphTextOf(paragraph: ParagraphNode): string {
  let text = "";
  for (const child of paragraph.children) {
    if (child.type === "text") {
      text += child.text;
    }
  }
  return text;
}

function collectNodeParagraphs(
  nodes: readonly DocNode[],
  targets: ParagraphCommentTarget[],
  counter: { value: number }
): void {
  for (const node of nodes) {
    if (node.type === "paragraph") {
      const paragraphIndex = counter.value;
      counter.value += 1;
      // `sourceXml` is emitted verbatim, markers included.
      if (node.sourceXml || !node.commentRanges || node.commentRanges.length === 0) {
        continue;
      }
      const text = paragraphTextOf(node);
      const ranges = resolveParagraphCommentRanges(text, node.commentRanges)
        .filter((range) => range.end > range.start)
        .map((range) => ({
          commentId: range.commentId,
          start: range.start,
          end: range.end,
        }));
      if (ranges.length > 0) {
        targets.push({ paragraphIndex, ranges });
      }
      continue;
    }

    collectTableParagraphs(node, targets, counter);
  }
}

function collectTableParagraphs(
  table: TableNode,
  targets: ParagraphCommentTarget[],
  counter: { value: number }
): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      collectNodeParagraphs(cell.nodes, targets, counter);
    }
  }
}

/**
 * Paragraphs whose comment ranges only exist in the model (their `sourceXml`
 * was cleared by an edit) and the offsets to re-emit for them.
 */
export function collectUnserializedCommentRanges(
  model: DocModel
): ParagraphCommentTarget[] {
  const targets: ParagraphCommentTarget[] = [];
  const counter = { value: 0 };
  collectNodeParagraphs(model.nodes, targets, counter);
  return targets;
}

export function modelNeedsCommentRangeSerialization(model: DocModel): boolean {
  return collectUnserializedCommentRanges(model).length > 0;
}

interface MarkerEvent {
  /** Offset in the paragraph text where the marker belongs. */
  offset: number;
  xml: string;
  /** Ordering among markers at the same offset (range starts before ends). */
  order: number;
}

/**
 * Splits one text atom's run at the marker offsets inside it.
 *
 * `w:commentRangeStart` / `w:commentRangeEnd` are paragraph-level content and
 * may not sit inside `w:r`, so the run is cut into one run per text segment and
 * the markers go between the runs (each run keeps the original `w:rPr`).
 */
function splitRunAtMarkers(
  content: string,
  atom: XmlTextAtom,
  events: readonly MarkerEvent[]
): DocumentXmlEdit {
  const runXml = content.slice(atom.runStart, atom.runEnd);
  const runOpenEnd = runXml.indexOf(">") + 1;
  const runOpenTag = runXml.slice(0, runOpenEnd);
  const runProperties = extractRunProperties(runXml);
  const textOpenTag = ensurePreserveSpace(atom.openTag);
  const runForText = (text: string): string =>
    `${runOpenTag}${runProperties}${textOpenTag}${escapeXmlText(text)}${atom.closeTag}</w:r>`;

  const ordered = [...events].sort(
    (left, right) =>
      left.offset - atom.textOffset - (right.offset - atom.textOffset) ||
      left.order - right.order
  );

  const pieces: string[] = [];
  let cursor = 0;
  for (const event of ordered) {
    const localOffset = Math.max(
      cursor,
      Math.min(atom.text.length, event.offset - atom.textOffset)
    );
    if (localOffset > cursor) {
      pieces.push(runForText(atom.text.slice(cursor, localOffset)));
      cursor = localOffset;
    }
    pieces.push(event.xml);
  }
  pieces.push(runForText(atom.text.slice(cursor)));

  return {
    start: atom.runStart,
    end: atom.runEnd,
    replacement: pieces.join(""),
  };
}

function insertCommentMarkersIntoParagraphContent(
  content: string,
  ranges: readonly SerializedCommentRange[]
): string {
  if (ranges.length === 0 || !/<w:t\b/i.test(content)) {
    return content;
  }

  const atoms = collectRunTextAtoms(content);
  if (atoms.length === 0) {
    return content;
  }

  const events: MarkerEvent[] = [];
  ranges.forEach((range, rangeIndex) => {
    events.push({
      offset: range.start,
      xml: `<w:commentRangeStart w:id="${range.commentId}"/>`,
      order: rangeIndex * 2,
    });
    events.push({
      offset: range.end,
      xml: `<w:commentRangeEnd w:id="${range.commentId}"/>`,
      order: rangeIndex * 2 + 1,
    });
  });

  const edits: DocumentXmlEdit[] = [];
  const boundaryEvents = new Map<number, MarkerEvent[]>();

  for (const event of events) {
    const atom = atoms.find(
      (candidate) =>
        event.offset > candidate.textOffset &&
        event.offset < candidate.textOffset + candidate.text.length
    );
    if (atom) {
      continue;
    }

    // Boundary marker: attach it to the nearest run edge at that offset.
    let index = 0;
    for (const candidate of atoms) {
      if (candidate.textOffset + candidate.text.length <= event.offset) {
        index = candidate.runEnd;
      } else if (candidate.textOffset >= event.offset) {
        index = candidate.runStart;
        break;
      }
    }
    const existing = boundaryEvents.get(index) ?? [];
    existing.push(event);
    boundaryEvents.set(index, existing);
  }

  for (const atom of atoms) {
    const atomEvents = events.filter(
      (event) =>
        event.offset > atom.textOffset &&
        event.offset < atom.textOffset + atom.text.length
    );
    if (atomEvents.length === 0) {
      continue;
    }
    edits.push(splitRunAtMarkers(content, atom, atomEvents));
  }

  for (const [index, group] of boundaryEvents) {
    const xml = [...group]
      .sort((left, right) => left.order - right.order)
      .map((event) => event.xml)
      .join("");
    edits.push({ start: index, end: index, replacement: xml });
  }

  // Word needs a `w:commentReference` run in the paragraph; the end of the
  // paragraph keeps it out of the highlighted text.
  const references = ranges
    .map((range) => `<w:r><w:commentReference w:id="${range.commentId}"/></w:r>`)
    .join("");
  if (references.length > 0) {
    edits.push({ start: content.length, end: content.length, replacement: references });
  }

  let patched = content;
  for (const edit of edits.sort((left, right) => right.start - left.start)) {
    patched =
      patched.slice(0, edit.start) + edit.replacement + patched.slice(edit.end);
  }
  return patched;
}

/**
 * Re-attaches comment range markers to serialized paragraphs whose model
 * comment offsets are no longer present in `sourceXml`.
 */
export function applyCommentRangesToDocumentXml(
  documentXml: string,
  model: DocModel
): string {
  const targets = collectUnserializedCommentRanges(model);
  if (targets.length === 0) {
    return documentXml;
  }

  const spans = locateBodyParagraphSpans(documentXml);
  if (!spans) {
    return documentXml;
  }

  const targetByIndex = new Map(
    targets.map((target) => [target.paragraphIndex, target])
  );
  const edits: DocumentXmlEdit[] = [];

  spans.paragraphs.forEach((span, paragraphIndex) => {
    const target = targetByIndex.get(paragraphIndex);
    if (!target) return;
    const content = documentXml.slice(span.contentStart, span.contentEnd);
    const patched = insertCommentMarkersIntoParagraphContent(
      content,
      target.ranges
    );
    if (patched === content) return;
    edits.push({
      start: span.contentStart,
      end: span.contentEnd,
      replacement: patched,
    });
  });

  return applyDocumentXmlEdits(documentXml, edits);
}
