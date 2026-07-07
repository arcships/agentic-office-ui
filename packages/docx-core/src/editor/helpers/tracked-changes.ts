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
  FormFieldRunNode,
  ParagraphNode,
  TableNode,
  TextRunNode
} from "../../engine/types";
import type {
  DocxComment,
  DocxTextRangeLocation,
  DocxTrackedChange,
  DocxTrackedChangeKind,
  ParagraphLocation
} from "./editor-types";
import {
  extractBalancedTagRanges,
  decodeXmlText,
  normalizeTrackedChangeSnippet,
  parseRunStyleFromRunXml,
  parseTrackedRunTokens,
  stripTextBoxContentFromRunXml,
  trackedChangeKindFromTagName,
  type RevisionTagRange,
  type XmlBalancedTagRange
} from "./xml-parsing";
import { xmlAttribute } from "./ooxml-helpers";
import { setCacheEntry } from "./cache-utils";
import {
  firstTableCellAnchorLocation,
  paragraphLocationKey
} from "./text-mutation";
import { tableCellParagraphs, tableCellParagraphsRecursively } from "./paragraph-inspect";

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
  changes: ParagraphTrackedInlineChange[];
}

export interface ParagraphCommentMarkup {
  commentIdsByVisibleChildIndex: Array<number[] | undefined>;
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

function summarizeTableFormattingChange(
  scope: "table" | "row" | "cell",
  changeXml: string
): string {
  const features: string[] = [];
  if (/<w:tblW\b|<w:tcW\b|<w:gridSpan\b/i.test(changeXml)) {
    features.push("width");
  }
  if (/<w:tblLayout\b/i.test(changeXml)) {
    features.push("layout");
  }
  if (/<w:tblInd\b|<w:ind\b/i.test(changeXml)) {
    features.push("indent");
  }
  if (/<w:tblCellMar\b|<w:tcMar\b/i.test(changeXml)) {
    features.push("margins");
  }
  if (/<w:(?:tblBorders|tcBorders|trBorders|pBdr)\b/i.test(changeXml)) {
    features.push("borders");
  }
  if (/<w:trHeight\b/i.test(changeXml)) {
    features.push("row height");
  }
  if (/<w:vAlign\b/i.test(changeXml)) {
    features.push("vertical align");
  }
  if (/<w:jc\b/i.test(changeXml)) {
    features.push("alignment");
  }
  if (/<w:shd\b/i.test(changeXml)) {
    features.push("shading");
  }

  const prefix =
    scope === "table"
      ? "Table formatting"
      : scope === "row"
      ? "Row formatting"
      : "Cell formatting";
  return summarizeChangeFeatures(prefix, features, prefix);
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
  let visibleChildIndex = 0;

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

  const runPattern = /<w:r\b[\s\S]*?<\/w:r>/gi;
  for (const runMatch of sourceXml.matchAll(runPattern)) {
    const runXml = runMatch[0] ?? "";
    if (!runXml) {
      continue;
    }

    const runStart = runMatch.index ?? 0;
    const runEnd = runStart + runXml.length;
    const enclosingRevision = revisionRanges
      .filter(
        (revisionRange) =>
          runStart >= revisionRange.start && runEnd <= revisionRange.end
      )
      .sort(
        (left, right) => left.end - left.start - (right.end - right.start)
      )[0];

    const revisionKind = enclosingRevision?.kind;
    const revisionId = enclosingRevision?.revisionId;
    const revisionAuthor = enclosingRevision?.author;
    const revisionDate = enclosingRevision?.date;
    const isDeletionLike =
      revisionKind === "deletion" || revisionKind === "move-from";
    const contentRunXml = stripTextBoxContentFromRunXml(runXml);
    const trackedRunStyle = parseRunStyleFromRunXml(contentRunXml);
    const visibleTokens = parseTrackedRunTokens(contentRunXml, false);
    const deletedTokens = parseTrackedRunTokens(contentRunXml, true);
    const hasImage = /<w:(?:drawing|pict)\b/i.test(runXml);
    const visibleChildCount =
      visibleTokens.filter((token) => token.text.length > 0 || token.isNote)
        .length + (hasImage ? 1 : 0);
    const visibleText = normalizeTrackedChangeSnippet(
      visibleTokens.map((token) => token.text).join("")
    );
    const deletedText = normalizeTrackedChangeSnippet(
      deletedTokens.map((token) => token.text).join("")
    );

    if (isDeletionLike) {
      const deletionSnippet = deletedText ?? (hasImage ? "[image]" : undefined);
      if (!deletionSnippet) {
        continue;
      }
      const change = getOrCreateChange(
        revisionKind,
        revisionId,
        revisionAuthor,
        revisionDate,
        deletionSnippet
      );
      if (deletedText) {
        const segments =
          deletedSegmentsByVisibleChildIndex.get(visibleChildIndex) ?? [];
        segments.push({
          text: deletedText,
          change,
          style: trackedRunStyle,
        });
        deletedSegmentsByVisibleChildIndex.set(visibleChildIndex, segments);
      }
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

    if (revisionKind === "insertion" || revisionKind === "move-to") {
      const insertionSnippet =
        visibleText ?? (hasImage ? "[image]" : undefined);
      if (!insertionSnippet) {
        visibleChildIndex += visibleChildCount;
        continue;
      }
      const inlineChange = getOrCreateChange(
        revisionKind,
        revisionId,
        revisionAuthor,
        revisionDate,
        insertionSnippet
      );
      for (let index = 0; index < visibleChildCount; index += 1) {
        inlineChangeByVisibleChildIndex[visibleChildIndex + index] =
          inlineChange;
      }
    }

    visibleChildIndex += visibleChildCount;
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

  // Some DOCX revisions wrap non-run content. Ensure every revision range is
  // represented even if the run-based pass above does not emit it.
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
    const revisionSnippet = revisionText ?? (hasImage ? "[image]" : undefined);
    if (!revisionSnippet) {
      return;
    }
    getOrCreateChange(
      revisionRange.kind,
      revisionRange.revisionId,
      revisionRange.author,
      revisionRange.date,
      revisionSnippet
    );
  });

  const hasInlineChanges = inlineChangeByVisibleChildIndex.some(Boolean);
  const hasDeletedSegments = deletedSegmentsByVisibleChildIndex.size > 0;
  if (!hasInlineChanges && !hasDeletedSegments && changes.length === 0) {
    setCacheEntry(paragraphTrackedMarkupBySourceXml, sourceXml, null);
    return undefined;
  }

  const resolved: ParagraphTrackedMarkup = {
    inlineChangeByVisibleChildIndex,
    deletedSegmentsByVisibleChildIndex,
    changes,
  };
  setCacheEntry(paragraphTrackedMarkupBySourceXml, sourceXml, resolved);
  return resolved;
}

// ---------------------------------------------------------------------------
// resolveParagraphCommentMarkup (upstream 17555-17651)
// ---------------------------------------------------------------------------

/**
 * Maps comment ranges (`commentRangeStart`/`commentRangeEnd`) to the
 * paragraph's visible child indexes using the same run accounting as
 * `resolveParagraphTrackedMarkup`, so the run renderer can highlight
 * commented content at the matching child cursor.
 */
export function resolveParagraphCommentMarkup(
  paragraph: ParagraphNode
): ParagraphCommentMarkup | undefined {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml || !/commentRange|commentReference/i.test(sourceXml)) {
    return undefined;
  }

  const cached = paragraphCommentMarkupBySourceXml.get(sourceXml);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const rangeStartById = new Map<number, number>();
  const rangeEndById = new Map<number, number>();
  for (const match of sourceXml.matchAll(
    /<w:commentRangeStart\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi
  )) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId) && match.index !== undefined) {
      rangeStartById.set(commentId, match.index + match[0].length);
    }
  }
  for (const match of sourceXml.matchAll(
    /<w:commentRangeEnd\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi
  )) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId) && match.index !== undefined) {
      rangeEndById.set(commentId, match.index);
    }
  }
  // Ranges may open in an earlier paragraph (start missing) or close in a
  // later one (end missing); treat the missing side as the paragraph edge.
  const ranges: Array<{ commentId: number; start: number; end: number }> = [];
  const rangeIds = new Set<number>([
    ...rangeStartById.keys(),
    ...rangeEndById.keys(),
  ]);
  rangeIds.forEach((commentId) => {
    const start = rangeStartById.get(commentId) ?? 0;
    const end = rangeEndById.get(commentId) ?? sourceXml.length;
    if (end > start) {
      ranges.push({ commentId, start, end });
    }
  });
  if (ranges.length === 0) {
    setCacheEntry(paragraphCommentMarkupBySourceXml, sourceXml, null);
    return undefined;
  }

  const commentIdsByVisibleChildIndex: Array<number[] | undefined> = [];
  let visibleChildIndex = 0;
  const runPattern = /<w:r\b[\s\S]*?<\/w:r>/gi;
  for (const runMatch of sourceXml.matchAll(runPattern)) {
    const runXml = runMatch[0] ?? "";
    if (!runXml) {
      continue;
    }

    const runStart = runMatch.index ?? 0;
    const contentRunXml = stripTextBoxContentFromRunXml(runXml);
    const visibleTokens = parseTrackedRunTokens(contentRunXml, false);
    const hasImage = /<w:(?:drawing|pict)\b/i.test(runXml);
    const visibleChildCount =
      visibleTokens.filter((token) => token.text.length > 0 || token.isNote)
        .length + (hasImage ? 1 : 0);
    if (visibleChildCount === 0) {
      continue;
    }

    const activeCommentIds = ranges
      .filter((range) => runStart >= range.start && runStart < range.end)
      .map((range) => range.commentId);
    if (activeCommentIds.length > 0) {
      for (let index = 0; index < visibleChildCount; index += 1) {
        commentIdsByVisibleChildIndex[visibleChildIndex + index] =
          activeCommentIds;
      }
    }
    visibleChildIndex += visibleChildCount;
  }

  if (commentIdsByVisibleChildIndex.length === 0) {
    setCacheEntry(paragraphCommentMarkupBySourceXml, sourceXml, null);
    return undefined;
  }

  const resolved: ParagraphCommentMarkup = { commentIdsByVisibleChildIndex };
  setCacheEntry(paragraphCommentMarkupBySourceXml, sourceXml, resolved);
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
  const sourceXml = table.sourceXml ?? "";
  if (!sourceXml) {
    return [];
  }

  const anchorLocation = firstTableCellAnchorLocation(table, tableIndex);
  if (!anchorLocation) {
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
  const entryByKey = new Map<string, (typeof entries)[number]>();

  const append = (
    scope: "table" | "row" | "cell",
    changeTag: XmlBalancedTagRange
  ): void => {
    const kind: DocxTrackedChangeKind = "paragraph-format-change";
    const author =
      decodeXmlText(xmlAttribute(changeTag.openTag, "w:author") ?? "") ||
      undefined;
    const date = xmlAttribute(changeTag.openTag, "w:date")?.trim() || undefined;
    const revisionId =
      xmlAttribute(changeTag.openTag, "w:id")?.trim() || undefined;
    const changeXml = sourceXml.slice(changeTag.start, changeTag.end);
    const text = summarizeTableFormattingChange(scope, changeXml);
    const key = revisionId
      ? `${scope}:${kind}:id:${revisionId}`
      : `${scope}:${kind}:${author ?? ""}:${date ?? ""}:${text}`;
    const existing = entryByKey.get(key);
    if (existing) {
      if (!existing.text && text) {
        existing.text = text;
      }
      if (!existing.author && author) {
        existing.author = author;
      }
      if (!existing.date && date) {
        existing.date = date;
      }
      return;
    }

    const stableId = revisionId
      ? `${scope}-${kind}-${revisionId}`
      : `${scope}-${kind}-${entryByKey.size}`;
    const next = {
      stableId,
      kind,
      author,
      date,
      text,
      location: {
        kind: "table-cell" as const,
        tableIndex,
        rowIndex: anchorLocation.rowIndex,
        cellIndex: anchorLocation.cellIndex,
        paragraphIndex: anchorLocation.paragraphIndex,
      },
    };
    entries.push(next);
    entryByKey.set(key, next);
  };

  extractBalancedTagRanges(sourceXml, "w:tblPrChange").forEach((range) =>
    append("table", range)
  );
  extractBalancedTagRanges(sourceXml, "w:trPrChange").forEach((range) =>
    append("row", range)
  );
  extractBalancedTagRanges(sourceXml, "w:tcPrChange").forEach((range) =>
    append("cell", range)
  );

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
        const directParagraphs = tableCellParagraphs(cell.nodes);
        directParagraphs.forEach((paragraph, paragraphIndex) => {
          appendParagraphChanges(paragraph, nodeIndex, {
            kind: "table-cell",
            tableIndex: nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex,
          });
        });

        // Nested tables inside a cell are rendered without per-paragraph location
        // attributes, so anchor these changes to the owning cell via a negative
        // paragraph index to avoid colliding with direct paragraph indexes.
        const nestedParagraphs = tableCellParagraphsRecursively(
          cell.nodes
        ).filter((paragraph) => !directParagraphs.includes(paragraph));
        nestedParagraphs.forEach((paragraph, nestedParagraphIndex) => {
          appendParagraphChanges(paragraph, nodeIndex, {
            kind: "table-cell",
            tableIndex: nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex: -(nestedParagraphIndex + 1),
          });
        });
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
// collectCommentsFromModel (upstream 22908-23003)
// ---------------------------------------------------------------------------

/**
 * Collects every comment reference found in the document body, pairing each
 * `w:commentReference` with its definition from `model.metadata.comments` and
 * resolving an anchor-text excerpt from the surrounding comment range.
 * Read-only — never mutates the model.
 */
export function collectCommentsFromModel(model: DocModel): DocxComment[] {
  const definitions = model.metadata.comments ?? [];
  if (definitions.length === 0) {
    return [];
  }
  const definitionById = new Map(
    definitions.map((definition) => [definition.id, definition])
  );

  const comments: DocxComment[] = [];
  const appendParagraphComments = (
    paragraph: ParagraphNode,
    nodeIndex: number,
    location: ParagraphLocation
  ): void => {
    const sourceXml = paragraph.sourceXml ?? "";
    if (!sourceXml || !/commentReference/i.test(sourceXml)) {
      return;
    }

    for (const match of sourceXml.matchAll(
      /<w:commentReference\b[^>]*w:id="(-?\d+)"/gi
    )) {
      const commentId = Number.parseInt(match[1] ?? "", 10);
      const definition = Number.isFinite(commentId)
        ? definitionById.get(commentId)
        : undefined;
      if (!definition) {
        continue;
      }

      comments.push({
        id: `${paragraphLocationKey(location)}:comment:${commentId}`,
        commentId,
        author: definition.author,
        initials: definition.initials,
        date: definition.date,
        text: definition.text,
        parentId: definition.parentId,
        resolved: definition.resolved,
        anchorText: resolveCommentAnchorText(sourceXml, commentId),
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
    }
  };

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      appendParagraphComments(node, nodeIndex, {
        kind: "paragraph",
        nodeIndex,
      });
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        const directParagraphs = tableCellParagraphs(cell.nodes);
        directParagraphs.forEach((paragraph, paragraphIndex) => {
          appendParagraphComments(paragraph, nodeIndex, {
            kind: "table-cell",
            tableIndex: nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex,
          });
        });

        const nestedParagraphs = tableCellParagraphsRecursively(
          cell.nodes
        ).filter((paragraph) => !directParagraphs.includes(paragraph));
        nestedParagraphs.forEach((paragraph, nestedParagraphIndex) => {
          appendParagraphComments(paragraph, nodeIndex, {
            kind: "table-cell",
            tableIndex: nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex: -(nestedParagraphIndex + 1),
          });
        });
      });
    });
  });

  return comments;
}
