// Tracked-changes & comments: read-only derivation from the document model.
// Upstream editor.tsx: lines 382-418 (markup types + caches),
// 17170-17651 (format-change summarizers + paragraph markup resolvers),
// 22675-23003 (model-level collectors).
//
// Per alignment point #32: tracked changes and comments are derived purely by
// reading the model (and its sourceXml). This module never produces mutation
// marks — it only extracts existing revision/comment markup so the view layer
// can render balloons, gutters and inline highlights.
//
// The gutter card positioning/rendering pure functions (upstream 23005-23536)
// live in tracked-changes-gutter.ts.

import type {
  DocModel,
  DocNode,
  DocumentCommentAnchor,
  DocumentCommentDefinition,
  FooterSection,
  FormFieldRunNode,
  HeaderSection,
  ParagraphNode,
  TableNode,
  TextRunNode
} from "../../engine/types";
import {
  getCachedDocumentCommentAnchor,
  setCachedDocumentCommentAnchor,
} from "../../engine/comment-anchors";
import {
  parseParagraphCommentRanges,
  resolveParagraphCommentRanges,
  type ParagraphCommentRange,
} from "../../engine/comment-ranges";
import type {
  DocxComment,
  DocxInlineTrackedChangeKind,
  DocxTextRangeLocation,
  DocxTrackedChange,
  DocxTrackedChangeKind,
  ParagraphLocation
} from "./editor-types";
import {
  extractBalancedTagRanges,
  extractTopLevelTagRanges,
  decodeXmlText,
  normalizeTrackedChangeSnippet,
  parseRunStyleFromRunXml,
  parseTrackedRunTokens,
  stripTextBoxContentFromRunXml,
  trackedChangeKindFromTagName,
  type RevisionTagRange
} from "./xml-parsing";
import { xmlAttribute } from "./ooxml-helpers";
import { setCacheEntry } from "./cache-utils";
import { paragraphLocationKey } from "./text-mutation";
import {
  paragraphText,
  tableCellParagraphEntries
} from "./paragraph-inspect";
import {
  collectParagraphVisibleUnits,
  mapVisibleUnitsToParagraphChildren,
  paragraphVisibleCursorBefore
} from "./tracked-changes-inline";
import {
  resolveTableTrackedMarkup,
  tableRevisionAnchorLocation
} from "./tracked-changes-table";

// ---------------------------------------------------------------------------
// Paragraph tracked/comment markup types (upstream 382-409)
// ---------------------------------------------------------------------------

export interface ParagraphTrackedInlineChange {
  id: string;
  kind: DocxTrackedChangeKind;
  author?: string;
  date?: string;
  text?: string;
}

export interface ParagraphTrackedDeletionSegment {
  text: string;
  change: ParagraphTrackedInlineChange;
  style?: TextRunNode["style"] | FormFieldRunNode["style"];
}

export interface ParagraphTrackedMarkup {
  inlineChangeByVisibleChildIndex: Array<
    ParagraphTrackedInlineChange | undefined
  >;
  deletedSegmentsByVisibleChildIndex: Map<
    number,
    ParagraphTrackedDeletionSegment[]
  >;
  /** Inserted/deleted paragraph mark (no visible child of its own). */
  paragraphMarkChange?: ParagraphTrackedInlineChange;
  changes: ParagraphTrackedInlineChange[];
}

/** A comment range expressed in a single rendered child's text. */
export interface ParagraphCommentTextRange {
  commentId: number;
  /** Inclusive offset inside the child's own text. */
  start: number;
  /** Exclusive offset inside the child's own text. */
  end: number;
}

export interface ParagraphCommentMarkup {
  /**
   * Comment ids touching each rendered child. A run that only partly intersects
   * a range still appears here so the gutter can match it from the DOM.
   */
  commentIdsByVisibleChildIndex: Array<number[] | undefined>;
  /**
   * Exact character windows per rendered child. Present when the paragraph
   * carries model-level comment offsets; the renderer splits the child's text
   * at these boundaries instead of highlighting the whole run.
   */
  rangesByVisibleChildIndex?: Array<ParagraphCommentTextRange[] | undefined>;
}

// Shared caches keyed by paragraph sourceXml. `null` marks a confirmed miss so
// re-entry skips re-parsing. (upstream 411-418)
export const paragraphTrackedMarkupBySourceXml = new Map<
  string,
  ParagraphTrackedMarkup | null
>();
export const paragraphCommentMarkupBySourceXml = new Map<
  string,
  ParagraphCommentMarkup | null
>();

// ---------------------------------------------------------------------------
// Format-change summarizers (upstream 17170-17288)
// ---------------------------------------------------------------------------

function summarizeChangeFeatures(
  prefix: string,
  features: string[],
  fallback: string
): string {
  if (features.length === 0) {
    return fallback;
  }

  const unique = Array.from(new Set(features));
  return `${prefix}: ${unique.join(", ")}`;
}

function summarizeRunFormattingChange(changeXml: string): string {
  const features: string[] = [];
  if (/<w:rFonts\b/i.test(changeXml)) {
    features.push("font");
  }
  if (/<w:sz\b/i.test(changeXml)) {
    features.push("size");
  }
  if (/<w:color\b/i.test(changeXml)) {
    features.push("color");
  }
  if (/<w:highlight\b/i.test(changeXml)) {
    features.push("highlight");
  }
  if (/<w:b(?:Cs)?\b/i.test(changeXml)) {
    features.push("bold");
  }
  if (/<w:i(?:Cs)?\b/i.test(changeXml)) {
    features.push("italic");
  }
  if (/<w:u\b/i.test(changeXml)) {
    features.push("underline");
  }
  if (/<w:strike\b/i.test(changeXml)) {
    features.push("strikethrough");
  }
  if (/<w:vertAlign\b/i.test(changeXml)) {
    features.push("baseline");
  }
  return summarizeChangeFeatures("Run formatting", features, "Run formatting");
}

function summarizeParagraphFormattingChange(changeXml: string): string {
  const features: string[] = [];
  if (/<w:ind\b/i.test(changeXml)) {
    features.push("margins/indent");
  }
  if (/<w:spacing\b/i.test(changeXml)) {
    features.push("line spacing");
  }
  if (/<w:jc\b/i.test(changeXml)) {
    features.push("alignment");
  }
  if (/<w:tabs\b/i.test(changeXml)) {
    features.push("tabs");
  }
  if (/<w:numPr\b/i.test(changeXml)) {
    features.push("numbering");
  }
  if (/<w:pBdr\b/i.test(changeXml)) {
    features.push("borders");
  }
  if (/<w:shd\b/i.test(changeXml)) {
    features.push("shading");
  }
  if (/<w:rPr\b/i.test(changeXml)) {
    features.push("text style");
  }
  return summarizeChangeFeatures(
    "Paragraph formatting",
    features,
    "Paragraph formatting"
  );
}

/**
 * Deleted run text that has no model child (`w:delText`) — the renderer needs
 * a synthetic struck-through segment for these.
 */
function deletedRunText(runXml: string): string | undefined {
  const texts: string[] = [];
  for (const match of runXml.matchAll(
    /<w:delText\b[^>]*>([\s\S]*?)<\/w:delText>/gi
  )) {
    texts.push(decodeXmlText(match[1] ?? ""));
  }
  return normalizeTrackedChangeSnippet(texts.join(""));
}

// ---------------------------------------------------------------------------
// resolveParagraphTrackedMarkup (upstream 17290-17553)
// ---------------------------------------------------------------------------

/**
 * Extracts inline tracked-change markup from a paragraph's sourceXml, keyed by
 * visible child index so the run renderer can attach inline-change metadata
 * and deletion-segment overlays. Results are cached by sourceXml.
 */
export function resolveParagraphTrackedMarkup(
  paragraph: ParagraphNode
): ParagraphTrackedMarkup | undefined {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml) {
    return undefined;
  }

  const cached = paragraphTrackedMarkupBySourceXml.get(sourceXml);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const revisionRanges: RevisionTagRange[] = [];
  for (const tagName of ["w:ins", "w:del", "w:moveFrom", "w:moveTo"] as const) {
    const kind = trackedChangeKindFromTagName(tagName);
    if (!kind) {
      continue;
    }

    extractBalancedTagRanges(sourceXml, tagName).forEach((range) => {
      revisionRanges.push({
        ...range,
        kind,
        revisionId: xmlAttribute(range.openTag, "w:id"),
        author: decodeXmlText(xmlAttribute(range.openTag, "w:author") ?? ""),
        date: xmlAttribute(range.openTag, "w:date"),
      });
    });
  }

  // Map XML units to the model children the renderer actually walks. This is
  // the single source of truth for the visible-child cursor used below.
  const visibleUnits = collectParagraphVisibleUnits(sourceXml);
  mapVisibleUnitsToParagraphChildren(visibleUnits, paragraph.children);

  const inlineChangeByVisibleChildIndex: Array<
    ParagraphTrackedInlineChange | undefined
  > = [];
  const deletedSegmentsByVisibleChildIndex = new Map<
    number,
    ParagraphTrackedDeletionSegment[]
  >();
  const changes: ParagraphTrackedInlineChange[] = [];
  const changeByKey = new Map<string, ParagraphTrackedInlineChange>();
  let anonymousChangeCounter = 0;

  const getOrCreateChange = (
    kind: DocxTrackedChangeKind,
    revisionId: string | undefined,
    author: string | undefined,
    date: string | undefined,
    text: string | undefined
  ): ParagraphTrackedInlineChange => {
    const normalizedAuthor = author?.trim() || undefined;
    const normalizedDate = date?.trim() || undefined;
    const normalizedText = normalizeTrackedChangeSnippet(text);
    const normalizedRevisionId = revisionId?.trim() || undefined;
    // Word groups tracked items primarily by revision id. Keep a single card per
    // revision id/kind to avoid over-fragmenting changes into many tiny entries.
    const key = normalizedRevisionId
      ? `${kind}:id:${normalizedRevisionId}`
      : `${kind}:anon:${normalizedAuthor ?? ""}:${normalizedDate ?? ""}:${
          normalizedText ?? ""
        }`;
    const existing = changeByKey.get(key);
    if (existing) {
      if (!existing.text && normalizedText) {
        existing.text = normalizedText;
      } else if (
        existing.text &&
        normalizedText &&
        normalizedText.length > existing.text.length
      ) {
        existing.text = normalizedText;
      }
      if (!existing.author && normalizedAuthor) {
        existing.author = normalizedAuthor;
      }
      if (!existing.date && normalizedDate) {
        existing.date = normalizedDate;
      }
      return existing;
    }

    const stableId = normalizedRevisionId
      ? `${kind}-${normalizedRevisionId}`
      : `${kind}-inline-${anonymousChangeCounter}`;
    anonymousChangeCounter += 1;
    const next: ParagraphTrackedInlineChange = {
      id: stableId,
      kind,
      author: normalizedAuthor,
      date: normalizedDate,
      text: normalizedText,
    };
    changeByKey.set(key, next);
    changes.push(next);
    return next;
  };

  const innermostRevisionAt = (
    offset: number
  ): RevisionTagRange | undefined =>
    revisionRanges
      .filter((range) => offset >= range.start && offset < range.end)
      .sort(
        (left, right) => left.end - left.start - (right.end - right.start)
      )[0];

  // Resolve revision snippets once, then annotate the mapped model children.
  const snippetByRevision = new Map<RevisionTagRange, string | undefined>();
  revisionRanges.forEach((revisionRange) => {
    const revisionXml = sourceXml.slice(revisionRange.start, revisionRange.end);
    const includeDeletedText =
      revisionRange.kind === "deletion" || revisionRange.kind === "move-from";
    const revisionTokens = parseTrackedRunTokens(
      stripTextBoxContentFromRunXml(revisionXml),
      includeDeletedText
    );
    const revisionText = normalizeTrackedChangeSnippet(
      revisionTokens.map((token) => token.text).join("")
    );
    const hasImage = /<w:(?:drawing|pict)\b/i.test(revisionXml);
    snippetByRevision.set(
      revisionRange,
      revisionText ?? (hasImage ? "[image]" : undefined)
    );
  });

  revisionRanges.forEach((revisionRange) => {
    const revisionSnippet = snippetByRevision.get(revisionRange);
    if (!revisionSnippet) {
      return;
    }
    const change = getOrCreateChange(
      revisionRange.kind,
      revisionRange.revisionId,
      revisionRange.author,
      revisionRange.date,
      revisionSnippet
    );
    for (const unit of visibleUnits) {
      if (unit.visibleChildIndex === undefined) {
        continue;
      }
      if (unit.start < revisionRange.start || unit.start >= revisionRange.end) {
        continue;
      }
      if (innermostRevisionAt(unit.start) !== revisionRange) {
        continue;
      }
      inlineChangeByVisibleChildIndex[unit.visibleChildIndex] = change;
    }
  });

  const runPattern = /<w:r\b[\s\S]*?<\/w:r>/gi;
  for (const runMatch of sourceXml.matchAll(runPattern)) {
    const runXml = runMatch[0] ?? "";
    if (!runXml) {
      continue;
    }

    const runStart = runMatch.index ?? 0;
    const enclosingRevision = innermostRevisionAt(runStart);

    const revisionKind = enclosingRevision?.kind;
    const revisionId = enclosingRevision?.revisionId;
    const revisionAuthor = enclosingRevision?.author;
    const revisionDate = enclosingRevision?.date;
    const isDeletionLike =
      revisionKind === "deletion" || revisionKind === "move-from";
    const contentRunXml = stripTextBoxContentFromRunXml(runXml);
    const trackedRunStyle = parseRunStyleFromRunXml(contentRunXml);

    if (isDeletionLike) {
      // Deleted text with no model child (w:delText) is rendered as a synthetic
      // struck-through segment at the mapped cursor. Deleted content that does
      // have a model child (w:t) is styled inline via the revision pass above.
      const deletedText = deletedRunText(
        stripTextBoxContentFromRunXml(runXml)
      );
      if (!deletedText) {
        continue;
      }
      const change = getOrCreateChange(
        revisionKind,
        revisionId,
        revisionAuthor,
        revisionDate,
        deletedText
      );
      const cursor = paragraphVisibleCursorBefore(visibleUnits, runStart);
      const segments = deletedSegmentsByVisibleChildIndex.get(cursor) ?? [];
      segments.push({
        text: deletedText,
        change,
        style: trackedRunStyle,
      });
      deletedSegmentsByVisibleChildIndex.set(cursor, segments);
      continue;
    }

    const rPrChangeRanges = extractBalancedTagRanges(runXml, "w:rPrChange");
    rPrChangeRanges.forEach((rPrChangeRange) => {
      const rPrChangeTag = rPrChangeRange.openTag ?? "";
      const rPrChangeXml = runXml.slice(
        rPrChangeRange.start,
        rPrChangeRange.end
      );
      const formatAuthor =
        decodeXmlText(xmlAttribute(rPrChangeTag, "w:author") ?? "") ||
        revisionAuthor;
      const formatDate = xmlAttribute(rPrChangeTag, "w:date") ?? revisionDate;
      const formatId = xmlAttribute(rPrChangeTag, "w:id") ?? revisionId;
      const formatSnippet = summarizeRunFormattingChange(rPrChangeXml);
      getOrCreateChange(
        "format-change",
        formatId,
        formatAuthor,
        formatDate,
        formatSnippet
      );
    });
  }

  const pPrChangeRanges = extractBalancedTagRanges(sourceXml, "w:pPrChange");
  pPrChangeRanges.forEach((pPrChangeRange) => {
    const pPrChangeTag = pPrChangeRange.openTag ?? "";
    const pPrChangeXml = sourceXml.slice(
      pPrChangeRange.start,
      pPrChangeRange.end
    );
    getOrCreateChange(
      "paragraph-format-change",
      xmlAttribute(pPrChangeTag, "w:id"),
      decodeXmlText(xmlAttribute(pPrChangeTag, "w:author") ?? ""),
      xmlAttribute(pPrChangeTag, "w:date"),
      summarizeParagraphFormattingChange(pPrChangeXml)
    );
  });

  // A paragraph mark revision lives in pPr/rPr and has no w:t/w:delText of its
  // own, so it never reaches the run pass above.
  let paragraphMarkChange: ParagraphTrackedInlineChange | undefined;
  for (const pPrRange of extractTopLevelTagRanges(sourceXml, "w:pPr")) {
    const pPrXml = sourceXml.slice(pPrRange.start, pPrRange.end);
    const pPrChangeInnerRanges = extractBalancedTagRanges(
      pPrXml,
      "w:pPrChange"
    );
    for (const rPrRange of extractTopLevelTagRanges(pPrXml, "w:rPr")) {
      if (
        pPrChangeInnerRanges.some(
          (changeRange) =>
            rPrRange.start >= changeRange.start &&
            rPrRange.end <= changeRange.end
        )
      ) {
        continue;
      }
      const rPrXml = pPrXml.slice(rPrRange.start, rPrRange.end);
      const markMatch = rPrXml.match(/<w:(ins|del)\b[^>]*\/?>/i);
      if (!markMatch) {
        continue;
      }
      const markTag = markMatch[0] ?? "";
      const kind: DocxInlineTrackedChangeKind =
        (markMatch[1] ?? "").toLowerCase() === "ins"
          ? "insertion"
          : "deletion";
      paragraphMarkChange = getOrCreateChange(
        kind,
        xmlAttribute(markTag, "w:id"),
        decodeXmlText(xmlAttribute(markTag, "w:author") ?? ""),
        xmlAttribute(markTag, "w:date"),
        "Paragraph mark"
      );
    }
  }
  if (!paragraphMarkChange && paragraph.paragraphMarkDeleted === true) {
    paragraphMarkChange = getOrCreateChange(
      "deletion",
      undefined,
      undefined,
      undefined,
      "Paragraph mark"
    );
  }

  const hasInlineChanges = inlineChangeByVisibleChildIndex.some(Boolean);
  const hasDeletedSegments = deletedSegmentsByVisibleChildIndex.size > 0;
  if (!hasInlineChanges && !hasDeletedSegments && changes.length === 0) {
    setCacheEntry(paragraphTrackedMarkupBySourceXml, sourceXml, null);
    return undefined;
  }

  const resolved: ParagraphTrackedMarkup = {
    inlineChangeByVisibleChildIndex,
    deletedSegmentsByVisibleChildIndex,
    paragraphMarkChange,
    changes,
  };
  setCacheEntry(paragraphTrackedMarkupBySourceXml, sourceXml, resolved);
  return resolved;
}

// ---------------------------------------------------------------------------
// resolveParagraphCommentMarkup (upstream 17555-17651)
// ---------------------------------------------------------------------------

/**
 * The paragraph's live comment ranges as character offsets.
 *
 * Reads the model field that parsing put on the paragraph (kept up to date by
 * the editing commands) and falls back to re-parsing `sourceXml` for models
 * built before the field existed. Ranges are reconciled with the current text,
 * so a comment keeps its place even when an edit could not be remapped.
 */
export function resolveParagraphCommentRangesForParagraph(
  paragraph: ParagraphNode
): ParagraphCommentRange[] {
  const text = paragraphText(paragraph);
  if (paragraph.commentRanges && paragraph.commentRanges.length > 0) {
    return resolveParagraphCommentRanges(text, paragraph.commentRanges);
  }

  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml || !/commentRange|commentReference/i.test(sourceXml)) {
    return [];
  }

  return resolveParagraphCommentRanges(
    text,
    parseParagraphCommentRanges(sourceXml, text.length)
  );
}

function mapCommentRangesToVisibleChildren(
  paragraph: ParagraphNode,
  ranges: readonly ParagraphCommentRange[]
): ParagraphCommentMarkup | undefined {
  const commentIdsByVisibleChildIndex: Array<number[] | undefined> = [];
  const rangesByVisibleChildIndex: Array<ParagraphCommentTextRange[] | undefined> = [];
  let textCursor = 0;
  let visibleChildIndex = 0;
  let matched = false;

  for (const child of paragraph.children) {
    // The run renderer skips form fields without advancing its cursor.
    if (child.type === "form-field") {
      continue;
    }
    if (child.type !== "text") {
      visibleChildIndex += 1;
      continue;
    }

    const childStart = textCursor;
    const childEnd = textCursor + child.text.length;
    textCursor = childEnd;

    const childRanges: ParagraphCommentTextRange[] = [];
    const childCommentIds: number[] = [];
    for (const range of ranges) {
      const overlapStart = Math.max(range.start, childStart);
      const overlapEnd = Math.min(range.end, childEnd);
      if (overlapEnd > overlapStart) {
        childRanges.push({
          commentId: range.commentId,
          start: overlapStart - childStart,
          end: overlapEnd - childStart,
        });
        if (!childCommentIds.includes(range.commentId)) {
          childCommentIds.push(range.commentId);
        }
        matched = true;
        continue;
      }

      // Point anchor (a comment with only `w:commentReference`) or a range that
      // owns no characters of this run: tag the run without a highlight so the
      // marker and card can still find it.
      if (
        range.start === range.end &&
        range.start >= childStart &&
        range.start <= childEnd &&
        !childCommentIds.includes(range.commentId)
      ) {
        childCommentIds.push(range.commentId);
      }
    }

    if (childRanges.length > 0) {
      rangesByVisibleChildIndex[visibleChildIndex] = childRanges;
    }
    if (childCommentIds.length > 0) {
      commentIdsByVisibleChildIndex[visibleChildIndex] =
        childCommentIds.sort((left, right) => left - right);
    }
    visibleChildIndex += 1;
  }

  if (!matched && commentIdsByVisibleChildIndex.length === 0) {
    return undefined;
  }

  return { commentIdsByVisibleChildIndex, rangesByVisibleChildIndex };
}

/**
 * Resolves comment highlighting for one paragraph.
 *
 * Model offsets are authoritative: they are captured when the document is
 * parsed and remapped by every text edit, so the highlight is exact to the
 * character and survives edits that clear `sourceXml`. Only when the model has
 * no offsets does this re-parse the paragraph markup (compatibility path).
 */
export function resolveParagraphCommentMarkup(
  paragraph: ParagraphNode
): ParagraphCommentMarkup | undefined {
  const text = paragraphText(paragraph);
  if (paragraph.commentRanges && paragraph.commentRanges.length > 0) {
    return mapCommentRangesToVisibleChildren(
      paragraph,
      resolveParagraphCommentRanges(text, paragraph.commentRanges)
    );
  }

  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml || !/commentRange|commentReference/i.test(sourceXml)) {
    return undefined;
  }

  const cached = paragraphCommentMarkupBySourceXml.get(sourceXml);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const resolved = mapCommentRangesToVisibleChildren(
    paragraph,
    resolveParagraphCommentRanges(
      text,
      parseParagraphCommentRanges(sourceXml, text.length)
    )
  );
  setCacheEntry(paragraphCommentMarkupBySourceXml, sourceXml, resolved ?? null);
  return resolved;
}

// ---------------------------------------------------------------------------
// collectTablePropertyTrackedChanges (upstream 22675-22768)
// ---------------------------------------------------------------------------

function collectTablePropertyTrackedChanges(
  table: TableNode,
  tableIndex: number
): Array<{
  stableId: string;
  kind: DocxTrackedChangeKind;
  author?: string;
  date?: string;
  text?: string;
  location: DocxTextRangeLocation;
}> {
  const markup = resolveTableTrackedMarkup(table);
  if (!markup) {
    return [];
  }

  const entries: Array<{
    stableId: string;
    kind: DocxTrackedChangeKind;
    author?: string;
    date?: string;
    text?: string;
    location: DocxTextRangeLocation;
  }> = [];

  markup.revisions.forEach((revision) => {
    const location = tableRevisionAnchorLocation(table, tableIndex, revision);
    if (!location) {
      return;
    }
    entries.push({
      stableId: revision.stableId,
      kind: revision.kind,
      author: revision.author,
      date: revision.date,
      text: revision.text,
      location: {
        kind: "table-cell",
        tableIndex,
        rowIndex: location.rowIndex,
        cellIndex: location.cellIndex,
        paragraphIndex: location.paragraphIndex,
      },
    });
  });

  return entries;
}

// ---------------------------------------------------------------------------
// collectTrackedChangesFromModel (upstream 22770-22864)
// ---------------------------------------------------------------------------

/**
 * Walks the document model and collects every tracked change (inline
 * insertions/deletions/moves, run & paragraph format changes, and table/row/
 * cell property changes) as flat `DocxTrackedChange` entries anchored to their
 * paragraph locations. Read-only — never mutates the model.
 */
export function collectTrackedChangesFromModel(
  model: DocModel
): DocxTrackedChange[] {
  const trackedChanges: DocxTrackedChange[] = [];

  const appendParagraphChanges = (
    paragraph: ParagraphNode,
    nodeIndex: number,
    location: ParagraphLocation
  ): void => {
    const trackedMarkup = resolveParagraphTrackedMarkup(paragraph);
    if (!trackedMarkup) {
      return;
    }

    trackedMarkup.changes.forEach((change, changeIndex) => {
      trackedChanges.push({
        id: `${paragraphLocationKey(location)}:${change.id}:${changeIndex}`,
        inlineAnchorId: change.id,
        kind: change.kind,
        author: change.author,
        date: change.date,
        text: change.text,
        nodeIndex,
        location:
          location.kind === "paragraph"
            ? { kind: "paragraph", nodeIndex: location.nodeIndex }
            : {
                kind: "table-cell",
                tableIndex: location.tableIndex,
                rowIndex: location.rowIndex,
                cellIndex: location.cellIndex,
                paragraphIndex: location.paragraphIndex,
              },
      });
    });
  };

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      appendParagraphChanges(node, nodeIndex, {
        kind: "paragraph",
        nodeIndex,
      });
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        tableCellParagraphEntries(cell.nodes).forEach(
          ({ paragraph, paragraphIndex }) => {
            appendParagraphChanges(paragraph, nodeIndex, {
              kind: "table-cell",
              tableIndex: nodeIndex,
              rowIndex,
              cellIndex,
              paragraphIndex,
            });
          }
        );
      });
    });

    collectTablePropertyTrackedChanges(node, nodeIndex).forEach(
      (change, changeIndex) => {
        trackedChanges.push({
          id: `${paragraphLocationKey(change.location)}:${
            change.stableId
          }:${changeIndex}`,
          kind: change.kind,
          author: change.author,
          date: change.date,
          text: change.text,
          nodeIndex,
          location: change.location,
        });
      }
    );
  });

  return trackedChanges;
}

// ---------------------------------------------------------------------------
// Comment anchor-text helpers (upstream 22866-22906)
// ---------------------------------------------------------------------------

function decodeCommentRangeText(rangeXml: string): string | undefined {
  const texts: string[] = [];
  for (const match of rangeXml.matchAll(
    /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
  )) {
    texts.push(decodeXmlText(match[1] ?? ""));
  }
  const combined = texts.join("").replace(/\s+/g, " ").trim();
  if (!combined) {
    return undefined;
  }
  return combined.length > 120 ? `${combined.slice(0, 119)}…` : combined;
}

function resolveCommentAnchorText(
  sourceXml: string,
  commentId: number
): string | undefined {
  const startMatch = sourceXml.match(
    new RegExp(`<w:commentRangeStart\\b[^>]*w:id="${commentId}"[^>]*/?>`, "i")
  );
  const endMatch = sourceXml.match(
    new RegExp(`<w:commentRangeEnd\\b[^>]*w:id="${commentId}"[^>]*/?>`, "i")
  );
  const startIndex =
    startMatch?.index !== undefined
      ? startMatch.index + startMatch[0].length
      : // Range opened in an earlier paragraph: take from the paragraph start.
        endMatch?.index !== undefined
      ? 0
      : undefined;
  if (startIndex === undefined) {
    return undefined;
  }
  const endIndex =
    endMatch?.index !== undefined ? endMatch.index : sourceXml.length;
  if (endIndex <= startIndex) {
    return undefined;
  }
  return decodeCommentRangeText(sourceXml.slice(startIndex, endIndex));
}

// ---------------------------------------------------------------------------
// Comment anchor discovery
// ---------------------------------------------------------------------------

interface ResolvedCommentAnchor {
  anchor: DocumentCommentAnchor;
  location: DocxTextRangeLocation;
}

interface CommentRegionPart {
  region: "body" | "header" | "footer";
  partName?: string;
  nodes: DocNode[];
}

function sameCommentPartName(left?: string, right?: string): boolean {
  if (left === right) {
    return true;
  }
  if (!left || !right) {
    return false;
  }
  return (left.split("/").pop() ?? left) === (right.split("/").pop() ?? right);
}

function commentAnchorLocation(
  anchor: DocumentCommentAnchor
): DocxTextRangeLocation {
  if (anchor.rowIndex !== undefined && anchor.cellIndex !== undefined) {
    return {
      kind: "table-cell",
      tableIndex: anchor.nodeIndex,
      rowIndex: anchor.rowIndex,
      cellIndex: anchor.cellIndex,
      paragraphIndex: anchor.paragraphIndex ?? 0,
    };
  }
  return { kind: "paragraph", nodeIndex: anchor.nodeIndex };
}

function decodeCommentTextPrefix(xmlPrefix: string): string {
  let text = "";
  for (const match of xmlPrefix.matchAll(
    /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
  )) {
    text += decodeXmlText(match[1] ?? "");
  }
  return text;
}

function locateParagraphTextOffset(
  paragraph: ParagraphNode,
  textOffset: number
): { runIndex: number; offset: number } {
  const paragraphLength = paragraphText(paragraph).length;
  let remaining = Math.max(0, Math.min(textOffset, paragraphLength));
  let runIndex = 0;
  let lastOffset = 0;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      continue;
    }
    if (remaining <= child.text.length) {
      return { runIndex, offset: remaining };
    }
    remaining -= child.text.length;
    runIndex += 1;
    lastOffset = child.text.length;
  }

  return { runIndex: Math.max(0, runIndex - 1), offset: lastOffset };
}

function commentAnchorForParagraph(
  paragraph: ParagraphNode,
  base: Omit<
    DocumentCommentAnchor,
    "textOffset" | "runIndex" | "offset" | "anchorText"
  >,
  commentId: number,
  anchorXmlIndex: number,
  includeAnchorText: boolean
): DocumentCommentAnchor {
  const sourceXml = paragraph.sourceXml ?? "";
  const textOffset = Math.min(
    decodeCommentTextPrefix(sourceXml.slice(0, anchorXmlIndex)).length,
    paragraphText(paragraph).length
  );
  const position = locateParagraphTextOffset(paragraph, textOffset);

  return {
    ...base,
    textOffset,
    runIndex: position.runIndex,
    offset: position.offset,
    anchorText: includeAnchorText
      ? resolveCommentAnchorText(sourceXml, commentId)
      : undefined,
  };
}

function scanCommentAnchorsInParagraph(
  paragraph: ParagraphNode,
  base: Omit<
    DocumentCommentAnchor,
    "textOffset" | "runIndex" | "offset" | "anchorText"
  >,
  anchors: Map<number, ResolvedCommentAnchor>
): void {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml || !/commentRangeStart|commentReference/i.test(sourceXml)) {
    return;
  }

  const rangeStartById = new Map<number, number>();
  for (const match of sourceXml.matchAll(
    /<w:commentRangeStart\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi
  )) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId) && match.index !== undefined) {
      rangeStartById.set(commentId, match.index + match[0].length);
    }
  }

  const referenceById = new Map<number, number>();
  for (const match of sourceXml.matchAll(
    /<w:commentReference\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi
  )) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId) && match.index !== undefined) {
      referenceById.set(commentId, match.index);
    }
  }

  for (const commentId of new Set<number>([
    ...rangeStartById.keys(),
    ...referenceById.keys(),
  ])) {
    if (anchors.has(commentId)) {
      continue;
    }
    const rangeStart = rangeStartById.get(commentId);
    const anchorXmlIndex = rangeStart ?? referenceById.get(commentId);
    if (anchorXmlIndex === undefined) {
      continue;
    }
    const anchor = commentAnchorForParagraph(
      paragraph,
      base,
      commentId,
      anchorXmlIndex,
      rangeStart !== undefined
    );
    anchors.set(commentId, {
      anchor,
      location: commentAnchorLocation(anchor),
    });
  }
}

function scanCommentAnchorsInNodes(
  nodes: DocNode[],
  region: "body" | "header" | "footer",
  partName: string | undefined,
  anchors: Map<number, ResolvedCommentAnchor>
): void {
  nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      scanCommentAnchorsInParagraph(
        node,
        { region, partName, nodeIndex },
        anchors
      );
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        for (const { paragraph, paragraphIndex } of tableCellParagraphEntries(
          cell.nodes
        )) {
          scanCommentAnchorsInParagraph(
            paragraph,
            {
              region,
              partName,
              nodeIndex,
              rowIndex,
              cellIndex,
              paragraphIndex,
            },
            anchors
          );
        }
      });
    });
  });
}

function collectCommentRegionParts(model: DocModel): CommentRegionPart[] {
  const parts = new Map<string, CommentRegionPart>();
  const addSection = (
    region: "header" | "footer",
    section: HeaderSection | FooterSection
  ): void => {
    const partName = section.partName;
    const key = `${region}:${partName ?? ""}`;
    const existing = parts.get(key);
    if (!existing || (existing.nodes.length === 0 && (section.nodes ?? []).length > 0)) {
      parts.set(key, { region, partName, nodes: section.nodes ?? [] });
    }
  };

  for (const section of model.metadata.headerSections ?? []) {
    addSection("header", section);
  }
  for (const section of model.metadata.footerSections ?? []) {
    addSection("footer", section);
  }
  for (const section of model.metadata.sections ?? []) {
    for (const header of section.headerSections ?? []) {
      addSection("header", header);
    }
    for (const footer of section.footerSections ?? []) {
      addSection("footer", footer);
    }
  }

  return Array.from(parts.values());
}

/**
 * Anchors that come from the paragraph's model-level comment offsets.
 *
 * These are the authoritative anchors: parsing stores the offsets on the
 * paragraph and every text edit remaps them, so a comment keeps its anchor (and
 * its gutter card) after `sourceXml` is cleared.
 */
function collectModelCommentAnchors(
  model: DocModel
): Map<number, ResolvedCommentAnchor> {
  const anchors = new Map<number, ResolvedCommentAnchor>();
  scanModelCommentAnchorsInNodes(model.nodes, "body", undefined, anchors);
  for (const part of collectCommentRegionParts(model)) {
    scanModelCommentAnchorsInNodes(
      part.nodes,
      part.region,
      part.partName,
      anchors
    );
  }
  return anchors;
}

function scanModelCommentAnchorsInParagraph(
  paragraph: ParagraphNode,
  base: Omit<
    DocumentCommentAnchor,
    "textOffset" | "runIndex" | "offset" | "anchorText"
  >,
  anchors: Map<number, ResolvedCommentAnchor>
): void {
  if (!paragraph.commentRanges || paragraph.commentRanges.length === 0) {
    return;
  }

  const text = paragraphText(paragraph);
  for (const range of resolveParagraphCommentRanges(
    text,
    paragraph.commentRanges
  )) {
    if (anchors.has(range.commentId)) {
      continue;
    }

    const position = locateParagraphTextOffset(paragraph, range.start);
    const anchor: DocumentCommentAnchor = {
      ...base,
      textOffset: range.start,
      runIndex: position.runIndex,
      offset: position.offset,
      anchorText: text.slice(range.start, range.end) || undefined,
    };
    anchors.set(range.commentId, {
      anchor,
      location: commentAnchorLocation(anchor),
    });
  }
}

function scanModelCommentAnchorsInNodes(
  nodes: DocNode[],
  region: "body" | "header" | "footer",
  partName: string | undefined,
  anchors: Map<number, ResolvedCommentAnchor>
): void {
  nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      scanModelCommentAnchorsInParagraph(
        node,
        { region, partName, nodeIndex },
        anchors
      );
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        for (const { paragraph, paragraphIndex } of tableCellParagraphEntries(
          cell.nodes
        )) {
          scanModelCommentAnchorsInParagraph(
            paragraph,
            {
              region,
              partName,
              nodeIndex,
              rowIndex,
              cellIndex,
              paragraphIndex,
            },
            anchors
          );
        }
      });
    });
  });
}

function collectDiscoveredCommentAnchors(
  model: DocModel
): Map<number, ResolvedCommentAnchor> {
  const anchors = new Map<number, ResolvedCommentAnchor>();
  scanCommentAnchorsInNodes(model.nodes, "body", undefined, anchors);
  for (const part of collectCommentRegionParts(model)) {
    scanCommentAnchorsInNodes(
      part.nodes,
      part.region,
      part.partName,
      anchors
    );
  }
  return anchors;
}

function commentRegionNodes(
  model: DocModel,
  anchor: DocumentCommentAnchor
): DocNode[] {
  if (anchor.region === "body") {
    return model.nodes;
  }

  const parts = collectCommentRegionParts(model).filter(
    (part) => part.region === anchor.region
  );
  const matchingParts = parts.filter((part) =>
    sameCommentPartName(part.partName, anchor.partName)
  );
  const matching =
    matchingParts.find((part) => part.nodes.length > 0) ?? matchingParts[0];
  return (matching ?? parts[0])?.nodes ?? [];
}

function paragraphAtCommentAnchor(
  model: DocModel,
  anchor: DocumentCommentAnchor
): ParagraphNode | undefined {
  const node = commentRegionNodes(model, anchor)[anchor.nodeIndex];
  if (!node) {
    return undefined;
  }
  if (anchor.rowIndex === undefined) {
    return node.type === "paragraph" ? node : undefined;
  }
  if (node.type !== "table") {
    return undefined;
  }
  const cell = node.rows[anchor.rowIndex]?.cells[anchor.cellIndex ?? -1];
  if (!cell) {
    return undefined;
  }
  return tableCellParagraphEntries(cell.nodes).find(
    (entry) => entry.paragraphIndex === anchor.paragraphIndex
  )?.paragraph;
}

function firstParagraphCommentAnchor(
  nodes: DocNode[],
  region: "body" | "header" | "footer",
  partName?: string
): DocumentCommentAnchor | undefined {
  for (let nodeIndex = 0; nodeIndex < nodes.length; nodeIndex += 1) {
    const node = nodes[nodeIndex];
    if (!node) {
      continue;
    }
    if (node.type === "paragraph") {
      return { region, partName, nodeIndex };
    }
    for (let rowIndex = 0; rowIndex < node.rows.length; rowIndex += 1) {
      const row = node.rows[rowIndex];
      for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
        const entry = tableCellParagraphEntries(
          row.cells[cellIndex]?.nodes ?? []
        )[0];
        if (entry) {
          return {
            region,
            partName,
            nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex: entry.paragraphIndex,
          };
        }
      }
    }
  }
  return undefined;
}

function fallbackCommentAnchor(model: DocModel): DocumentCommentAnchor {
  return (
    firstParagraphCommentAnchor(model.nodes, "body") ??
    collectCommentRegionParts(model)
      .map((part) =>
        firstParagraphCommentAnchor(part.nodes, part.region, part.partName)
      )
      .find((anchor): anchor is DocumentCommentAnchor => Boolean(anchor)) ??
    { region: "body", nodeIndex: 0 }
  );
}

function resolveDefinedCommentAnchors(
  model: DocModel,
  definitions: DocumentCommentDefinition[]
): Map<number, ResolvedCommentAnchor> {
  const modelAnchors = collectModelCommentAnchors(model);
  const discovered = collectDiscoveredCommentAnchors(model);
  const definitionById = new Map(
    definitions.map((definition) => [definition.id, definition])
  );
  const resolved = new Map<number, ResolvedCommentAnchor>();
  const resolving = new Set<number>();
  const fallback = fallbackCommentAnchor(model);

  const resolve = (definition: DocumentCommentDefinition): ResolvedCommentAnchor => {
    const existing = resolved.get(definition.id);
    if (existing) {
      return existing;
    }
    if (resolving.has(definition.id)) {
      return { anchor: fallback, location: commentAnchorLocation(fallback) };
    }
    resolving.add(definition.id);

    const cached = getCachedDocumentCommentAnchor(definition);
    const cachedIsLocated = cached
      ? Boolean(paragraphAtCommentAnchor(model, cached))
      : false;
    const discoveredAnchor = discovered.get(definition.id)?.anchor;
    // Model offsets win: they are the only source that survives edits. The
    // cached/discovered sourceXml anchors are the compatibility fallbacks.
    const ownAnchor =
      modelAnchors.get(definition.id)?.anchor ??
      (cachedIsLocated ? cached : undefined) ??
      discoveredAnchor;

    let anchor = ownAnchor;
    if (!anchor && definition.parentId !== undefined) {
      const parent = definitionById.get(definition.parentId);
      if (parent) {
        anchor = resolve(parent).anchor;
      }
    }
    if (!anchor) {
      anchor = fallback;
    }

    if (ownAnchor) {
      setCachedDocumentCommentAnchor(definition, ownAnchor);
    }

    const result: ResolvedCommentAnchor = {
      anchor: {
        ...anchor,
        anchorText:
          ownAnchor?.anchorText ??
          discoveredAnchor?.anchorText ??
          anchor.anchorText,
      },
      location: commentAnchorLocation(anchor),
    };
    resolved.set(definition.id, result);
    resolving.delete(definition.id);
    return result;
  };

  for (const definition of definitions) {
    resolve(definition);
  }
  return resolved;
}

/**
 * Whether a comment belongs to a resolved thread. `resolved` is normally only
 * set on the thread root, so replies must walk their `parentId` chain.
 */
export function commentThreadIsResolved(
  comment: DocxComment,
  commentsById: Map<number, DocxComment>
): boolean {
  let current: DocxComment | undefined = comment;
  const seen = new Set<number>();
  while (current) {
    if (current.resolved) {
      return true;
    }
    if (current.parentId === undefined || seen.has(current.parentId)) {
      return false;
    }
    seen.add(current.parentId);
    current = commentsById.get(current.parentId);
  }
  return false;
}

// ---------------------------------------------------------------------------
// collectCommentsFromModel (upstream 22908-23003)
// ---------------------------------------------------------------------------

const COMMENT_REGION_ORDER = { body: 0, header: 1, footer: 2 } as const;

function compareCommentAnchors(
  left?: DocumentCommentAnchor,
  right?: DocumentCommentAnchor
): number {
  if (!left || !right) {
    return left ? -1 : right ? 1 : 0;
  }
  return (
    COMMENT_REGION_ORDER[left.region] - COMMENT_REGION_ORDER[right.region] ||
    left.nodeIndex - right.nodeIndex ||
    (left.rowIndex ?? 0) - (right.rowIndex ?? 0) ||
    (left.cellIndex ?? 0) - (right.cellIndex ?? 0) ||
    (left.paragraphIndex ?? 0) - (right.paragraphIndex ?? 0)
  );
}

/**
 * Collects every comment definition in `model.metadata.comments`. Locations
 * come from `w:commentRangeStart` / `w:commentReference` markup, with replies
 * inheriting their parent's anchor. Anchors are cached outside the model so
 * they survive edits that clear `sourceXml`; header and footer parts are
 * scanned alongside the body. Read-only — never mutates the document model.
 */
export function collectCommentsFromModel(model: DocModel): DocxComment[] {
  const definitions = model.metadata.comments ?? [];
  if (definitions.length === 0) {
    return [];
  }

  const definitionById = new Map(
    definitions.map((definition) => [definition.id, definition])
  );
  const anchors = resolveDefinedCommentAnchors(model, definitions);
  const comments: DocxComment[] = [];
  const emitted = new Set<number>();

  const emit = (definition: DocumentCommentDefinition): void => {
    if (emitted.has(definition.id)) {
      return;
    }
    emitted.add(definition.id);
    if (definition.parentId !== undefined) {
      const parent = definitionById.get(definition.parentId);
      if (parent) {
        emit(parent);
      }
    }

    const resolved = anchors.get(definition.id);
    if (!resolved) {
      return;
    }
    const { anchor, location } = resolved;
    const locationKey = paragraphLocationKey(location);
    comments.push({
      id:
        anchor.region === "body"
          ? `${locationKey}:comment:${definition.id}`
          : `${anchor.region}:${anchor.partName ?? ""}:${locationKey}:comment:${
              definition.id
            }`,
      commentId: definition.id,
      author: definition.author,
      initials: definition.initials,
      date: definition.date,
      text: definition.text,
      parentId: definition.parentId,
      resolved: definition.resolved,
      anchorText: anchor.anchorText,
      nodeIndex: anchor.nodeIndex,
      location,
      region: anchor.region,
      partName: anchor.partName,
    });
  };

  const orderedDefinitions = [...definitions].sort((left, right) =>
    compareCommentAnchors(
      anchors.get(left.id)?.anchor,
      anchors.get(right.id)?.anchor
    )
  );
  for (const definition of orderedDefinitions) {
    emit(definition);
  }
  return comments;
}

/**
 * Pre-computes and caches comment anchors, typically right after a DOCX is
 * parsed. This keeps anchored comments visible even if they are hidden while
 * the user edits the document and `sourceXml` is cleared.
 */
export function primeCommentAnchors(model: DocModel): void {
  if (!model.metadata.comments?.length) {
    return;
  }
  collectCommentsFromModel(model);
}
