// Comment ranges as character offsets.
//
// Word anchors a comment with a pair of run-level markers inside a paragraph:
// `<w:commentRangeStart w:id="3"/>` ... `<w:commentRangeEnd w:id="3"/>`. The
// renderer used to re-parse that markup on every paint, which made highlights
// run-granular (a range covering half a run highlighted all of it) and made
// them disappear as soon as an edit cleared `paragraph.sourceXml`.
//
// This module turns the markers into plain character offsets into the
// paragraph's own text (`paragraphText()`, i.e. the concatenation of its text
// children), stores them on the paragraph, and keeps them correct across edits.
// Offsets are UTF-16 code units, the same space the editor's `DocxTextRange`
// and DOM `Range` use, so a highlight can be rendered exactly.

import type { ParagraphCommentRange } from "./types";

export type { ParagraphCommentRange };

const COMMENT_RANGE_START_PATTERN =
  /<w:commentRangeStart\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi;
const COMMENT_RANGE_END_PATTERN =
  /<w:commentRangeEnd\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi;
const COMMENT_REFERENCE_PATTERN =
  /<w:commentReference\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi;

/**
 * Units that advance the paragraph text cursor while scanning paragraph XML.
 * Mirrors the renderer's visible units: text runs, tabs, breaks and note
 * references. Text boxes and form-field content are skipped because the model
 * renders them as separate children (images do not contribute text).
 */
const VISIBLE_UNIT_PATTERN =
  /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/?>|<w:(?:br|cr)\b[^>]*\/?>|<w:(?:footnoteReference|endnoteReference)\b[^>]*\/?>/gi;

const NAMED_XML_ENTITIES: Record<string, string> = {
  amp: "&",
  apos: "'",
  gt: ">",
  lt: "<",
  quot: '"',
  nbsp: "\u00a0",
};

export function decodeCommentXmlText(value: string): string {
  if (!value.includes("&")) {
    return value;
  }

  return value.replace(/&(#x[\da-f]+|#\d+|[a-z]+);/gi, (match, entity: string) => {
    if (entity.startsWith("#x") || entity.startsWith("#X")) {
      const codePoint = Number.parseInt(entity.slice(2), 16);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : match;
    }
    if (entity.startsWith("#")) {
      const codePoint = Number.parseInt(entity.slice(1), 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : match;
    }
    return NAMED_XML_ENTITIES[entity.toLowerCase()] ?? match;
  });
}

/**
 * Offsets covered by content controls the model renders as a form field: their
 * inner `<w:t>` elements belong to the form-field child, not to the paragraph
 * text, so comment ranges must not count them.
 */
function formFieldContentRanges(
  sourceXml: string
): Array<{ start: number; end: number }> {
  if (!sourceXml.includes("w:sdt")) {
    return [];
  }

  const ranges: Array<{ start: number; end: number }> = [];
  const openPattern = /<w:sdt\b[^>]*>/gi;
  for (const open of sourceXml.matchAll(openPattern)) {
    const start = open.index ?? 0;
    const closeIndex = sourceXml.indexOf("</w:sdt>", start);
    if (closeIndex < 0) {
      continue;
    }
    const end = closeIndex + "</w:sdt>".length;
    const content = sourceXml.slice(start, end);
    const propertiesStart = content.indexOf("<w:sdtPr");
    const propertiesEnd =
      propertiesStart >= 0 ? content.indexOf("</w:sdtPr>", propertiesStart) : -1;
    if (propertiesStart < 0 || propertiesEnd < 0) {
      continue;
    }
    const properties = content.slice(propertiesStart, propertiesEnd);
    if (/<w:(?:text|checkbox|date|dropDownList|comboBox)\b|<w:ffData\b/i.test(properties)) {
      ranges.push({ start, end });
    }
  }
  return ranges;
}

interface XmlTextPiece {
  /** Start index in `sourceXml`. */
  start: number;
  /** End index in `sourceXml`. */
  end: number;
  /** Offset of this piece in the paragraph's plain text. */
  textOffset: number;
  text: string;
}

function collectXmlTextPieces(sourceXml: string): XmlTextPiece[] {
  const skipRanges = formFieldContentRanges(sourceXml);
  const textBoxRanges: Array<{ start: number; end: number }> = [];
  const textBoxPattern = /<w:txbxContent\b[^>]*>/gi;
  for (const open of sourceXml.matchAll(textBoxPattern)) {
    const start = open.index ?? 0;
    const closeIndex = sourceXml.indexOf("</w:txbxContent>", start);
    if (closeIndex < 0) {
      continue;
    }
    textBoxRanges.push({ start, end: closeIndex + "</w:txbxContent>".length });
  }

  const isSkipped = (index: number): boolean =>
    skipRanges.some((range) => index >= range.start && index < range.end) ||
    textBoxRanges.some((range) => index >= range.start && index < range.end);

  const pieces: XmlTextPiece[] = [];
  let textOffset = 0;
  for (const match of sourceXml.matchAll(VISIBLE_UNIT_PATTERN)) {
    const start = match.index ?? 0;
    if (isSkipped(start)) {
      continue;
    }

    const unit = match[0] ?? "";
    const end = start + unit.length;
    let text = "";
    if (match[1] !== undefined) {
      text = decodeCommentXmlText(match[1]);
    } else if (/^<w:tab\b/i.test(unit)) {
      text = "\t";
    } else if (/^<w:(?:br|cr)\b/i.test(unit)) {
      text = "\n";
    }

    if (text.length === 0) {
      // Note references and empty runs advance no characters, but they still
      // sit at a known offset, which is what a reference-only anchor needs.
      pieces.push({ start, end, textOffset, text: "" });
      continue;
    }

    pieces.push({ start, end, textOffset, text });
    textOffset += text.length;
  }

  return pieces;
}

function offsetForXmlIndex(pieces: XmlTextPiece[], xmlIndex: number): number {
  if (pieces.length === 0) {
    return 0;
  }

  let resolved = 0;
  for (const piece of pieces) {
    if (piece.start >= xmlIndex) {
      return piece.textOffset;
    }
    if (xmlIndex < piece.end) {
      const consumed = Math.max(
        0,
        Math.min(piece.text.length, xmlIndex - piece.start)
      );
      // Only a prefix that is itself text (not a tag) maps to characters.
      const prefix = piece.text.slice(0, consumed);
      return piece.textOffset + prefix.length;
    }
    resolved = piece.textOffset + piece.text.length;
  }
  return resolved;
}

/**
 * Parses `w:commentRangeStart` / `w:commentRangeEnd` into character offsets.
 *
 * A range that opens in an earlier paragraph or closes in a later one uses the
 * paragraph edge, matching Word's "whole paragraph" anchoring. A comment that
 * only has a `w:commentReference` (a point anchor) produces an empty range so
 * its card still has a position.
 */
export function parseParagraphCommentRanges(
  sourceXml: string,
  paragraphTextLength?: number
): ParagraphCommentRange[] {
  if (!sourceXml || !/commentRange|commentReference/i.test(sourceXml)) {
    return [];
  }

  const pieces = collectXmlTextPieces(sourceXml);
  const totalTextLength = pieces.reduce(
    (length, piece) => length + piece.text.length,
    0
  );
  const textLength =
    paragraphTextLength === undefined
      ? totalTextLength
      : Math.max(0, Math.min(paragraphTextLength, totalTextLength));
  const fullText = pieces.map((piece) => piece.text).join("");

  const startById = new Map<number, number>();
  const endById = new Map<number, number>();
  const referenceById = new Map<number, number>();

  for (const match of sourceXml.matchAll(COMMENT_RANGE_START_PATTERN)) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId)) {
      startById.set(commentId, offsetForXmlIndex(pieces, match.index ?? 0));
    }
  }
  for (const match of sourceXml.matchAll(COMMENT_RANGE_END_PATTERN)) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId)) {
      endById.set(commentId, offsetForXmlIndex(pieces, match.index ?? 0));
    }
  }
  for (const match of sourceXml.matchAll(COMMENT_REFERENCE_PATTERN)) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId)) {
      referenceById.set(commentId, offsetForXmlIndex(pieces, match.index ?? 0));
    }
  }

  const commentIds = new Set<number>([
    ...startById.keys(),
    ...endById.keys(),
    ...referenceById.keys(),
  ]);
  const ranges: ParagraphCommentRange[] = [];
  commentIds.forEach((commentId) => {
    const hasRange = startById.has(commentId) || endById.has(commentId);
    const start = clampOffset(
      startById.get(commentId) ?? (hasRange ? 0 : referenceById.get(commentId) ?? 0),
      textLength
    );
    const rawEnd = hasRange
      ? endById.get(commentId) ?? textLength
      : referenceById.get(commentId) ?? start;
    const end = Math.max(start, clampOffset(rawEnd, textLength));
    ranges.push({
      commentId,
      start,
      end,
      anchorText: fullText.slice(start, end) || undefined,
    });
  });

  ranges.sort((left, right) => left.start - right.start || left.commentId - right.commentId);
  return ranges;
}

function clampOffset(offset: number, textLength: number): number {
  if (!Number.isFinite(offset)) {
    return 0;
  }
  return Math.max(0, Math.min(Math.round(offset), textLength));
}

/**
 * Maps stored ranges onto the paragraph's current text.
 *
 * `commentRanges` are captured at parse time and remapped by each edit, so in
 * the common case `text.slice(start, end)` still equals `anchorText`. When an
 * edit could not be mapped (bulk replacement, undo of a structural change, a
 * document assembled by a host) the range is re-anchored by searching for its
 * `anchorText`, and finally clamped so it never points past the text.
 */
export function resolveParagraphCommentRanges(
  paragraphTextValue: string,
  ranges: readonly ParagraphCommentRange[] | undefined
): ParagraphCommentRange[] {
  if (!ranges || ranges.length === 0) {
    return [];
  }

  const textLength = paragraphTextValue.length;
  const resolved: ParagraphCommentRange[] = [];
  let searchFrom = 0;

  for (const range of ranges) {
    const anchorText = range.anchorText ?? "";
    if (anchorText.length === 0) {
      const start = clampOffset(range.start, textLength);
      resolved.push({ ...range, start, end: start });
      continue;
    }

    if (
      range.start >= 0 &&
      range.start + anchorText.length <= textLength &&
      paragraphTextValue.slice(range.start, range.start + anchorText.length) ===
        anchorText
    ) {
      resolved.push({
        ...range,
        start: range.start,
        end: range.start + anchorText.length,
      });
      searchFrom = Math.max(searchFrom, range.start + anchorText.length);
      continue;
    }

    let matchIndex = paragraphTextValue.indexOf(anchorText, searchFrom);
    if (matchIndex < 0) {
      matchIndex = paragraphTextValue.indexOf(anchorText);
    }
    if (matchIndex < 0) {
      const start = clampOffset(range.start, textLength);
      resolved.push({ ...range, start, end: start });
      continue;
    }

    resolved.push({
      ...range,
      start: matchIndex,
      end: matchIndex + anchorText.length,
    });
    searchFrom = matchIndex + anchorText.length;
  }

  return resolved;
}

interface CommonAffixes {
  /** Length of the unchanged prefix. */
  prefixLength: number;
  /** Length of the unchanged suffix (never overlapping the prefix). */
  suffixLength: number;
}

function commonAffixes(previousText: string, nextText: string): CommonAffixes {
  const maxPrefix = Math.min(previousText.length, nextText.length);
  let prefixLength = 0;
  while (
    prefixLength < maxPrefix &&
    previousText[prefixLength] === nextText[prefixLength]
  ) {
    prefixLength += 1;
  }

  const maxSuffix = Math.min(
    previousText.length - prefixLength,
    nextText.length - prefixLength
  );
  let suffixLength = 0;
  while (
    suffixLength < maxSuffix &&
    previousText[previousText.length - 1 - suffixLength] ===
      nextText[nextText.length - 1 - suffixLength]
  ) {
    suffixLength += 1;
  }

  return { prefixLength, suffixLength };
}

/**
 * Shifts comment ranges from `previousText` onto `nextText` after a text edit.
 *
 * The edit is modelled as "keep the common prefix, replace the middle, keep the
 * common suffix", which covers typing, deletion, paste and find/replace.
 * Offsets at or after the edit point move by the length delta; offsets before it
 * stay put. When a boundary fell inside the replaced middle the mapping is not
 * exact, so the original `anchorText` is kept and the read path re-anchors by
 * searching for it.
 */
export function remapParagraphCommentRanges(
  previousText: string,
  nextText: string,
  ranges: readonly ParagraphCommentRange[] | undefined
): ParagraphCommentRange[] | undefined {
  if (!ranges) {
    return undefined;
  }
  if (ranges.length === 0) {
    return [];
  }
  if (previousText === nextText) {
    return ranges as ParagraphCommentRange[];
  }

  const { prefixLength, suffixLength } = commonAffixes(previousText, nextText);
  const delta = nextText.length - previousText.length;
  const nextTextLength = nextText.length;
  const unchangedSuffixStart = previousText.length - suffixLength;

  const mapOffset = (offset: number): { value: number; exact: boolean } => {
    if (offset < prefixLength) {
      return { value: offset, exact: true };
    }
    if (offset >= unchangedSuffixStart) {
      return { value: offset + delta, exact: true };
    }
    // Inside the replaced middle: the best available approximation.
    return { value: prefixLength, exact: false };
  };

  const remapped: ParagraphCommentRange[] = [];
  for (const range of ranges) {
    const mappedStart = mapOffset(range.start);
    const mappedEnd = mapOffset(range.end);
    const start = clampOffset(
      Math.min(mappedStart.value, mappedEnd.value),
      nextTextLength
    );
    const end = clampOffset(
      Math.max(mappedStart.value, mappedEnd.value),
      nextTextLength
    );
    const exact = mappedStart.exact && mappedEnd.exact && end > start;
    const anchorText = exact ? nextText.slice(start, end) : range.anchorText;

    remapped.push({
      commentId: range.commentId,
      start,
      end,
      anchorText: anchorText && anchorText.length > 0 ? anchorText : undefined,
    });
  }

  return remapped;
}

/**
 * Splits ranges at `splitOffset` when a paragraph is split in two: the left
 * side keeps the text before the caret, the right side receives the rest.
 */
export function splitParagraphCommentRanges(
  ranges: readonly ParagraphCommentRange[] | undefined,
  splitOffset: number
): {
  before: ParagraphCommentRange[];
  after: ParagraphCommentRange[];
} {
  const before: ParagraphCommentRange[] = [];
  const after: ParagraphCommentRange[] = [];
  const offset = Math.max(0, Math.round(splitOffset));

  for (const range of ranges ?? []) {
    if (range.end <= offset) {
      before.push({ ...range });
      continue;
    }
    if (range.start >= offset) {
      after.push({
        ...range,
        start: range.start - offset,
        end: range.end - offset,
      });
      continue;
    }

    // The range straddles the split: both halves keep a piece of it.
    before.push({
      ...range,
      end: offset,
      anchorText: (range.anchorText ?? "").slice(0, offset - range.start),
    });
    after.push({
      ...range,
      start: 0,
      end: range.end - offset,
      anchorText: (range.anchorText ?? "").slice(offset - range.start),
    });
  }

  return { before, after };
}

/** Shifts `ranges` by `offset` after the paragraphs they belong to were merged. */
export function shiftParagraphCommentRanges(
  ranges: readonly ParagraphCommentRange[] | undefined,
  offset: number
): ParagraphCommentRange[] {
  const shift = Math.max(0, Math.round(offset));
  return (ranges ?? []).map((range) => ({
    ...range,
    start: range.start + shift,
    end: range.end + shift,
  }));
}

/**
 * Coalesces same-comment ranges that touch after an edit joined them.
 *
 * A range that spans a deleted region can end up as two fragments (one from
 * each side of the deletion); once the deleted text is gone the fragments are
 * adjacent and must become a single range again, because OOXML expects one
 * `commentRangeStart`/`commentRangeEnd` pair per comment per paragraph.
 */
export function mergeParagraphCommentRanges(
  ranges: readonly ParagraphCommentRange[] | undefined,
  paragraphTextValue?: string
): ParagraphCommentRange[] {
  if (!ranges || ranges.length === 0) {
    return [];
  }

  const sorted = [...ranges].sort(
    (left, right) => left.start - right.start || left.commentId - right.commentId
  );
  const merged: ParagraphCommentRange[] = [];

  for (const range of sorted) {
    const previous = merged[merged.length - 1];
    if (
      previous &&
      previous.commentId === range.commentId &&
      range.start <= previous.end
    ) {
      previous.end = Math.max(previous.end, range.end);
      previous.anchorText =
        paragraphTextValue === undefined
          ? undefined
          : paragraphTextValue.slice(previous.start, previous.end);
      continue;
    }
    merged.push({ ...range });
  }

  return merged;
}
