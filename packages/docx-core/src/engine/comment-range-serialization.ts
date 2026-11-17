// Re-emit comment range markers when a model is serialized.
//
// The WASM serializer writes a paragraph's `sourceXml` verbatim (which already
// contains `w:commentRangeStart`/`w:commentRangeEnd`) and otherwise renders it
// from model fields, which have no slot for those markers. Comment ranges now
// live on the model as character offsets, so after an edit clears `sourceXml`
// we patch the serialized body: walk the paragraph's generated text runs, split
// them at the stored offsets, and put the markers between the halves.

import type { DocModel, DocNode, ParagraphNode, TableNode } from "./types";
import { decodeCommentXmlText, resolveParagraphCommentRanges } from "./comment-ranges";

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

const TEXT_ATOM_PATTERN = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi;

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

interface ParagraphBounds {
  contentStart: number;
  contentEnd: number;
}

/** Locates every `<w:p>` content span in the part, in document order. */
function paragraphBounds(xml: string): ParagraphBounds[] {
  const bounds: ParagraphBounds[] = [];
  const tagPattern = /<(\/?)w:p\b[^>]*?(\/?)>/g;
  const openStack: number[] = [];
  let match: RegExpExecArray | null;

  tagPattern.lastIndex = 0;
  while ((match = tagPattern.exec(xml)) !== null) {
    const closing = match[1] === "/";
    const selfClosing = match[2] === "/";
    if (closing) {
      const contentStart = openStack.pop();
      if (contentStart === undefined) continue;
      bounds.push({ contentStart, contentEnd: match.index });
      continue;
    }
    if (selfClosing) {
      const empty = match.index + match[0].length;
      bounds.push({ contentStart: empty, contentEnd: empty });
      continue;
    }
    openStack.push(match.index + match[0].length);
  }

  // Outermost paragraphs are the ones the model describes; `w:p` nesting only
  // happens through text boxes, which the model renders as run content.
  return bounds;
}

interface TextAtom {
  /** Index of the `<w:t>` open tag inside the paragraph content. */
  tagStart: number;
  /** Index just past `</w:t>`. */
  tagEnd: number;
  /** Index of the enclosing `<w:r>` open tag. */
  runStart: number;
  /** Index just past the enclosing `</w:r>`. */
  runEnd: number;
  openTag: string;
  closeTag: string;
  /** Decoded text inside the `<w:t>` element. */
  text: string;
  /** Offset of the first character of this atom in the paragraph text. */
  textOffset: number;
}

function collectTextAtoms(content: string): TextAtom[] {
  const atoms: TextAtom[] = [];
  const runOpens: number[] = [];
  const runPairs: Array<{ start: number; end: number }> = [];

  for (const open of content.matchAll(/<w:r\b[^>]*>/g)) {
    runOpens.push(open.index ?? 0);
  }
  let openIndex = 0;
  for (const close of content.matchAll(/<\/w:r>/g)) {
    const start = runOpens[openIndex];
    if (start === undefined) break;
    runPairs.push({ start, end: (close.index ?? 0) + "</w:r>".length });
    openIndex += 1;
  }

  const runFor = (index: number): { start: number; end: number } | undefined =>
    runPairs.find((pair) => index >= pair.start && index < pair.end);

  let textOffset = 0;
  for (const match of content.matchAll(TEXT_ATOM_PATTERN)) {
    const tagStart = match.index ?? 0;
    const fullMatch = match[0] ?? "";
    const tagEnd = tagStart + fullMatch.length;
    const text = decodeCommentXmlText(match[1] ?? "");
    const run = runFor(tagStart);
    atoms.push({
      tagStart,
      tagEnd,
      runStart: run ? run.start : tagStart,
      runEnd: run ? run.end : tagEnd,
      openTag: fullMatch.slice(0, fullMatch.indexOf(">") + 1),
      closeTag: "</w:t>",
      text,
      textOffset,
    });
    textOffset += text.length;
  }

  return atoms;
}

function extractRunProperties(runXml: string): string {
  const match = /<w:rPr\b[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/.exec(runXml);
  return match ? match[0] : "";
}

function escapeXmlText(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

interface MarkerEvent {
  /** Offset in the paragraph text where the marker belongs. */
  offset: number;
  xml: string;
  /** Ordering among markers at the same offset (range starts before ends). */
  order: number;
}

interface XmlEdit {
  start: number;
  end: number;
  replacement: string;
}

function ensurePreserveSpace(openTag: string): string {
  return openTag.includes("xml:space")
    ? openTag
    : openTag.replace(/>$/, ' xml:space="preserve">');
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
  atom: TextAtom,
  events: readonly MarkerEvent[]
): XmlEdit {
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

  const atoms = collectTextAtoms(content);
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

  const edits: XmlEdit[] = [];
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

  const bodyOpen = /<w:body\b[^>]*>/i.exec(documentXml);
  if (!bodyOpen || bodyOpen.index === undefined) {
    return documentXml;
  }

  const bodyContentStart = bodyOpen.index + bodyOpen[0].length;
  const bodyCloseIndex = documentXml.toLowerCase().lastIndexOf("</w:body>");
  if (bodyCloseIndex < bodyContentStart) {
    return documentXml;
  }

  const bodyXml = documentXml.slice(bodyContentStart, bodyCloseIndex);
  const bounds = paragraphBounds(bodyXml);
  const targetByIndex = new Map(
    targets.map((target) => [target.paragraphIndex, target])
  );

  interface Edit {
    start: number;
    end: number;
    replacement: string;
  }
  const edits: Edit[] = [];

  bounds.forEach((paragraph, paragraphIndex) => {
    const target = targetByIndex.get(paragraphIndex);
    if (!target) return;
    const content = bodyXml.slice(paragraph.contentStart, paragraph.contentEnd);
    const patched = insertCommentMarkersIntoParagraphContent(
      content,
      target.ranges
    );
    if (patched === content) return;
    edits.push({
      start: paragraph.contentStart,
      end: paragraph.contentEnd,
      replacement: patched,
    });
  });

  if (edits.length === 0) {
    return documentXml;
  }

  let patchedBody = bodyXml;
  for (const edit of edits.sort((left, right) => right.start - left.start)) {
    patchedBody =
      patchedBody.slice(0, edit.start) +
      edit.replacement +
      patchedBody.slice(edit.end);
  }

  return (
    documentXml.slice(0, bodyContentStart) +
    patchedBody +
    documentXml.slice(bodyCloseIndex)
  );
}
