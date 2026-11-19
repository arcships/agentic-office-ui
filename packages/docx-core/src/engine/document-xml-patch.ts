// Shared primitives for the post-serialization `word/document.xml` patches.
//
// The WASM serializer emits a paragraph's `sourceXml` verbatim and otherwise
// renders the paragraph from model fields. Model fields have no slot for
// paragraph-level markers (comment ranges, note references, ...), so every patch
// has to locate the same `<w:p>` content spans and text runs and splice marker
// elements back in. This module owns those scans so the patches cannot drift
// apart.

import { decodeCommentXmlText } from "./comment-ranges";

/** Content span of one `<w:p>`, i.e. everything between its tags. */
export interface ParagraphContentSpan {
  contentStart: number;
  contentEnd: number;
}

export interface BodyParagraphSpans {
  /** Index just past the `<w:body ...>` open tag. */
  bodyContentStart: number;
  /** Index of `</w:body>`. */
  bodyContentEnd: number;
  /** Content spans of every `<w:p>` in the body, in document order. */
  paragraphs: ParagraphContentSpan[];
}

const PARAGRAPH_TAG_PATTERN = /<(\/?)w:p\b[^>]*?(\/?)>/g;
const BODY_CLOSE_PATTERN = /<\/w:body\s*>/gi;

/** Locates every `<w:p>` content span of the body, in document order. */
function paragraphContentSpans(bodyXml: string): ParagraphContentSpan[] {
  const spans: ParagraphContentSpan[] = [];
  const openStack: number[] = [];

  PARAGRAPH_TAG_PATTERN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = PARAGRAPH_TAG_PATTERN.exec(bodyXml)) !== null) {
    const closing = match[1] === "/";
    const selfClosing = match[2] === "/";
    if (closing) {
      const contentStart = openStack.pop();
      if (contentStart === undefined) continue;
      spans.push({ contentStart, contentEnd: match.index });
      continue;
    }
    if (selfClosing) {
      const empty = match.index + match[0].length;
      spans.push({ contentStart: empty, contentEnd: empty });
      continue;
    }
    openStack.push(match.index + match[0].length);
  }

  // Outermost paragraphs are the ones the model describes; `w:p` nesting only
  // happens through text boxes, which the model renders as run content.
  return spans;
}

/**
 * Locates the body and the content span of every paragraph in it, as absolute
 * offsets into `documentXml`.
 */
export function locateBodyParagraphSpans(
  documentXml: string
): BodyParagraphSpans | undefined {
  const bodyOpen = /<w:body\b[^>]*>/i.exec(documentXml);
  if (!bodyOpen || bodyOpen.index === undefined) {
    return undefined;
  }

  const bodyContentStart = bodyOpen.index + bodyOpen[0].length;
  let bodyContentEnd = -1;
  for (const match of documentXml.matchAll(BODY_CLOSE_PATTERN)) {
    bodyContentEnd = match.index ?? -1;
  }
  if (bodyContentEnd < bodyContentStart) {
    return undefined;
  }

  const bodyXml = documentXml.slice(bodyContentStart, bodyContentEnd);
  return {
    bodyContentStart,
    bodyContentEnd,
    paragraphs: paragraphContentSpans(bodyXml).map((span) => ({
      contentStart: bodyContentStart + span.contentStart,
      contentEnd: bodyContentStart + span.contentEnd,
    })),
  };
}

/**
 * Content spans of the *outermost* `<w:p>` elements of a part, in document
 * order. Paragraphs nested inside a text box (`w:txbxContent`) belong to the
 * drawing that carries them and are never model paragraphs.
 */
export function locateOutermostParagraphSpans(
  xml: string
): ParagraphContentSpan[] {
  const spans: ParagraphContentSpan[] = [];
  const openStack: Array<{ contentStart: number; depth: number }> = [];

  PARAGRAPH_TAG_PATTERN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = PARAGRAPH_TAG_PATTERN.exec(xml)) !== null) {
    const closing = match[1] === "/";
    const selfClosing = match[2] === "/";
    if (closing) {
      const open = openStack.pop();
      if (!open) continue;
      if (open.depth === 0) {
        spans.push({ contentStart: open.contentStart, contentEnd: match.index });
      }
      continue;
    }
    if (selfClosing) {
      if (openStack.length === 0) {
        const empty = match.index + match[0].length;
        spans.push({ contentStart: empty, contentEnd: empty });
      }
      continue;
    }
    openStack.push({
      contentStart: match.index + match[0].length,
      depth: openStack.length,
    });
  }

  return spans;
}

export interface XmlTextAtom {
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

const TEXT_ATOM_PATTERN = /<w:t\b[^>]*?\/>|<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi;

/** Collects every `<w:t>` element of a paragraph body with its run bounds. */
export function collectRunTextAtoms(content: string): XmlTextAtom[] {
  const atoms: XmlTextAtom[] = [];
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
    // The serializer writes an empty text run as a self-closing `<w:t/>`; it
    // carries no characters but it is where a marker run lives.
    const selfClosing = /\/>$/.test(fullMatch);
    const text = selfClosing ? "" : decodeCommentXmlText(match[1] ?? "");
    const run = runFor(tagStart);
    atoms.push({
      tagStart,
      tagEnd,
      runStart: run ? run.start : tagStart,
      runEnd: run ? run.end : tagEnd,
      openTag: selfClosing
        ? fullMatch.replace(/\s*\/>$/, ">")
        : fullMatch.slice(0, fullMatch.indexOf(">") + 1),
      closeTag: "</w:t>",
      text,
      textOffset,
    });
    textOffset += text.length;
  }

  return atoms;
}

/** The `<w:rPr>` element of a run, or `""` when it has none. */
export function extractRunProperties(runXml: string): string {
  const match = /<w:rPr\b[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/.exec(runXml);
  return match ? match[0] : "";
}

export function escapeXmlText(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function ensurePreserveSpace(openTag: string): string {
  return openTag.includes("xml:space")
    ? openTag
    : openTag.replace(/>$/, ' xml:space="preserve">');
}

export interface DocumentXmlEdit {
  start: number;
  end: number;
  replacement: string;
}

/**
 * Applies non-overlapping edits to an XML string. Edits are applied from the end
 * of the document backwards so earlier offsets stay valid; equal offsets keep
 * their input order.
 */
export function applyDocumentXmlEdits(
  xml: string,
  edits: readonly DocumentXmlEdit[]
): string {
  if (edits.length === 0) {
    return xml;
  }

  let patched = xml;
  for (const edit of [...edits].sort((left, right) => right.start - left.start)) {
    patched =
      patched.slice(0, edit.start) + edit.replacement + patched.slice(edit.end);
  }
  return patched;
}
