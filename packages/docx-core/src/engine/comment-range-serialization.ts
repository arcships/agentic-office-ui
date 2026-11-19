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
  /** Ids that need a `w:commentReference` run in the paragraph. */
  referenceIds: number[];
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
      if (!node.commentRanges || node.commentRanges.length === 0) {
        continue;
      }
      const text = paragraphTextOf(node);
      const resolved = resolveParagraphCommentRanges(text, node.commentRanges);
      const ranges = resolved
        .filter((range) => range.end > range.start)
        .map((range) => ({
          commentId: range.commentId,
          start: range.start,
          end: range.end,
        }));
      // A paragraph with `sourceXml` already carries the markers of the
      // comments it had when it was parsed; `applyCommentRangesToDocumentXml`
      // only inserts the ids it does not find there. A point anchor (start ===
      // end) still needs its reference run so Word keeps showing the card.
      const referenceIds = [...new Set(resolved.map((range) => range.commentId))];
      if (ranges.length > 0 || referenceIds.length > 0) {
        targets.push({ paragraphIndex, ranges, referenceIds });
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
 * Paragraphs that store comment ranges on the model and the offsets to
 * re-emit for them. The caller filters out the markers the paragraph's own
 * `sourceXml` already contains.
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
  ranges: readonly SerializedCommentRange[],
  referenceIds: readonly number[]
): string {
  const existingIds = new Set<number>();
  for (const match of content.matchAll(
    /<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*w:id="(-?\d+)"/g
  )) {
    const id = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(id)) {
      existingIds.add(id);
    }
  }

  // `sourceXml` already carries the markers of the comments it was parsed
  // with; only the ids the model added since then are inserted here.
  const pendingRanges = ranges.filter(
    (range) => !existingIds.has(range.commentId)
  );
  const missingReferences = referenceIds.filter((id) => !existingIds.has(id));

  const atoms = pendingRanges.length > 0 ? collectRunTextAtoms(content) : [];
  const events: MarkerEvent[] = [];
  pendingRanges.forEach((range, rangeIndex) => {
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

  // A paragraph without visible text (an empty paragraph or a point anchor at
  // its edge) still needs the reference run, but has no runs to split.
  if (atoms.length === 0) {
    const references = missingReferences
      .map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`)
      .join("");
    return references.length > 0 ? content + references : content;
  }

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
  const references = missingReferences
    .map((id) => `<w:r><w:commentReference w:id="${id}"/></w:r>`)
    .join("");
  if (references.length > 0) {
    edits.push({ start: content.length, end: content.length, replacement: references });
  }

  if (edits.length === 0) {
    return content;
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
      target.ranges,
      target.referenceIds
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
