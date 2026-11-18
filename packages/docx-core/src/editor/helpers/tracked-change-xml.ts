// Revision XML helpers shared by the tracked-change decision passes.
//
// Revisions live in a paragraph's `sourceXml`; accepting or rejecting one has to
// rewrite that XML rather than drop it, otherwise the remaining revisions would
// disappear from the inline view, the gutter and the exported document. These
// helpers locate, match and splice revision ranges; the decision modules decide
// what each range becomes.

import type { DocxTrackedChange, DocxTrackedChangeKind } from "./editor-types";
import {
  decodeXmlText,
  extractBalancedTagRanges,
  trackedChangeKindFromTagName,
  type XmlBalancedTagRange
} from "./xml-parsing";
import { xmlAttribute } from "./ooxml-helpers";

export interface XmlEdit {
  start: number;
  end: number;
  text: string;
}

/** Applies non-overlapping edits; nested edits are resolved in an earlier pass. */
export function applyXmlEdits(xml: string, edits: readonly XmlEdit[]): string {
  if (edits.length === 0) {
    return xml;
  }

  const ordered = [...edits].sort((left, right) => left.start - right.start);
  let output = "";
  let cursor = 0;
  for (const edit of ordered) {
    // Nested edits are resolved in an earlier pass; see the two-phase
    // property/content rewrite in applyTrackedChangeDecision.
    if (edit.start < cursor) {
      continue;
    }
    output += xml.slice(cursor, edit.start) + edit.text;
    cursor = edit.end;
  }
  return output + xml.slice(cursor);
}

function isSelfClosingTag(tagXml: string): boolean {
  return /\/>\s*$/u.test(tagXml);
}

export function innerXmlOfRange(xml: string, range: XmlBalancedTagRange): string {
  if (isSelfClosingTag(range.openTag)) {
    return "";
  }
  const slice = xml.slice(range.start, range.end);
  const closing = slice.match(/<\/([^>]+)>\s*$/u);
  if (!closing) {
    return slice.slice(range.openTag.length);
  }
  return slice.slice(range.openTag.length, slice.length - closing[0].length);
}

/**
 * Turns rejected deletion markup back into ordinary run text so the restored
 * children and the rewritten `sourceXml` describe the same content.
 */
export function deletedContentToVisibleXml(innerXml: string): string {
  return innerXml
    .replace(/<w:delText\b[^>]*>/giu, '<w:t xml:space="preserve">')
    .replace(/<\/w:delText\s*>/giu, "</w:t>")
    .replace(/<w:delInstrText\b[^>]*>/giu, "<w:instrText>")
    .replace(/<\/w:delInstrText\s*>/giu, "</w:instrText>");
}

// ---------------------------------------------------------------------------
// Revision ranges in `sourceXml`
// ---------------------------------------------------------------------------

export type ContentRevisionKind = Exclude<
  DocxTrackedChangeKind,
  | "format-change"
  | "paragraph-format-change"
  | "table-format-change"
  | "row-format-change"
  | "cell-format-change"
>;

export interface ContentRevisionRange extends XmlBalancedTagRange {
  kind: ContentRevisionKind;
  revisionId?: string;
  author?: string;
  date?: string;
}

const CONTENT_REVISION_TAGS = [
  "w:ins",
  "w:del",
  "w:moveFrom",
  "w:moveTo"
] as const;

export function collectContentRevisionRanges(
  xml: string
): ContentRevisionRange[] {
  const ranges: ContentRevisionRange[] = [];
  for (const tagName of CONTENT_REVISION_TAGS) {
    const kind = trackedChangeKindFromTagName(tagName);
    if (!kind) {
      continue;
    }
    extractBalancedTagRanges(xml, tagName).forEach((range) => {
      ranges.push({
        ...range,
        kind,
        revisionId: xmlAttribute(range.openTag, "w:id")?.trim() || undefined,
        author:
          decodeXmlText(xmlAttribute(range.openTag, "w:author") ?? "").trim() ||
          undefined,
        date: xmlAttribute(range.openTag, "w:date")?.trim() || undefined
      });
    });
  }
  return ranges.sort((left, right) => left.start - right.start);
}

export interface PropertyChangeRange extends XmlBalancedTagRange {
  revisionId?: string;
  author?: string;
  date?: string;
}

export function collectPropertyChangeRanges(
  xml: string,
  tagName: string
): PropertyChangeRange[] {
  return extractBalancedTagRanges(xml, tagName).map((range) => ({
    ...range,
    revisionId: xmlAttribute(range.openTag, "w:id")?.trim() || undefined,
    author:
      decodeXmlText(xmlAttribute(range.openTag, "w:author") ?? "").trim() ||
      undefined,
    date: xmlAttribute(range.openTag, "w:date")?.trim() || undefined
  }));
}

/** `insertion-2` → `{ kind: "insertion", revisionId: "2" }`. */
export function parseRevisionAnchor(
  inlineAnchorId: string | undefined
): { kind?: DocxTrackedChangeKind; revisionId?: string } {
  if (!inlineAnchorId) {
    return {};
  }
  const match = inlineAnchorId.match(
    /^(insertion|deletion|move-from|move-to|format-change|paragraph-format-change)-(.+)$/u
  );
  if (!match) {
    return {};
  }
  const kind = match[1] as DocxTrackedChangeKind;
  const suffix = match[2] ?? "";
  return suffix.startsWith("inline-")
    ? { kind }
    : { kind, revisionId: suffix };
}

interface RevisionMatchCandidate {
  revisionId?: string;
  author?: string;
  date?: string;
}

/**
 * Picks the source revision a collected change refers to. Word groups tracked
 * items by `w:id`, so the id wins; anonymous revisions fall back to
 * author/date and finally to the first still-unclaimed range of that kind.
 */
export function matchRevisionRange<T extends RevisionMatchCandidate>(
  change: DocxTrackedChange,
  ranges: readonly T[],
  claimed: ReadonlySet<T>
): T | undefined {
  const available = ranges.filter((range) => !claimed.has(range));
  if (available.length === 0) {
    return undefined;
  }

  const { revisionId } = parseRevisionAnchor(change.inlineAnchorId);
  if (revisionId) {
    const byId = available.find((range) => range.revisionId === revisionId);
    if (byId) {
      return byId;
    }
  }

  const byMeta = available.find(
    (range) =>
      (!change.author || range.author === change.author) &&
      (!change.date || range.date === change.date)
  );
  return byMeta ?? available[0];
}
