// Comment authoring operations: create, reply, resolve and delete.
//
// The read path already derives comment cards and highlights from
// `model.metadata.comments` plus the per-paragraph `commentRanges` offsets
// (see engine/comment-ranges.ts and helpers/tracked-changes.ts). This module is
// the write side: it keeps a thread's definitions and its anchors in step, so
// an edit is one immutable model revision that history can undo as a whole.

import type {
  DocModel,
  DocumentCommentDefinition,
  DocNode,
  ParagraphNode,
} from "../engine/types";
import { cloneDocModelShallow, cloneDocNode } from "../engine/clone";
import { mergeParagraphCommentRanges } from "../engine/comment-ranges";
import type {
  DocxTextRange,
  DocxTextRangeLocation,
  ParagraphLocation,
} from "./helpers/editor-types";
import { normalizeTextRange } from "./helpers/editor-types";
import {
  getParagraphAtLocation,
  paragraphRangeForMutate,
  sameParagraphLocation,
} from "./helpers/selection-helpers";
import { paragraphText } from "./helpers/paragraph-inspect";

/** Author metadata stored on a comment definition. */
export interface DocumentCommentAuthorOptions {
  author?: string;
  initials?: string;
  /** ISO-8601 timestamp; hosts pin it for deterministic exports. */
  date?: string;
}

export interface AddDocumentCommentInput extends DocumentCommentAuthorOptions {
  /** Selection the comment anchors to. */
  range: DocxTextRange;
  /** Plain-text comment body. */
  text: string;
  /** Explicit comment id; defaults to the next free id. */
  id?: number;
}

export interface DocumentCommentMutationResult {
  model: DocModel;
  commentId: number;
}

/** One paragraph covered by a comment range, in the paragraph's own offsets. */
export interface DocumentCommentAnchorRange {
  location: DocxTextRangeLocation;
  start: number;
  end: number;
  /** The commented text, used by the read path to re-anchor after edits. */
  text: string;
}

function commentDefinitions(model: DocModel): DocumentCommentDefinition[] {
  return model.metadata.comments ?? [];
}

export function findDocumentComment(
  model: DocModel,
  commentId: number
): DocumentCommentDefinition | undefined {
  return commentDefinitions(model).find((comment) => comment.id === commentId);
}

/** Next free `w:id` for `word/comments.xml`. Word accepts 0-based ids. */
export function nextDocumentCommentId(model: DocModel): number {
  let next = 0;
  for (const comment of commentDefinitions(model)) {
    if (Number.isFinite(comment.id) && comment.id >= next) {
      next = comment.id + 1;
    }
  }
  return next;
}

function commentById(model: DocModel): Map<number, DocumentCommentDefinition> {
  return new Map(commentDefinitions(model).map((comment) => [comment.id, comment]));
}

/** Thread root of `commentId` (the definition without a `parentId`). */
export function commentThreadRootId(
  model: DocModel,
  commentId: number
): number | undefined {
  const byId = commentById(model);
  let current = byId.get(commentId);
  const seen = new Set<number>();
  while (current) {
    if (current.parentId === undefined) {
      return current.id;
    }
    if (seen.has(current.parentId)) {
      return undefined;
    }
    seen.add(current.parentId);
    current = byId.get(current.parentId);
  }
  return undefined;
}

/** `commentId` plus every comment that replies to it, directly or transitively. */
export function commentSubtreeIds(model: DocModel, commentId: number): number[] {
  const definitions = commentDefinitions(model);
  const ids = new Set<number>([commentId]);
  let grew = true;
  while (grew) {
    grew = false;
    for (const comment of definitions) {
      if (
        comment.parentId !== undefined &&
        ids.has(comment.parentId) &&
        !ids.has(comment.id)
      ) {
        ids.add(comment.id);
        grew = true;
      }
    }
  }
  return [...ids];
}

/** Every definition in the thread rooted at `commentId`, root included. */
export function commentThreadMemberIds(
  model: DocModel,
  commentId: number
): number[] {
  const rootId = commentThreadRootId(model, commentId);
  return rootId === undefined ? [] : commentSubtreeIds(model, rootId);
}

function clampToParagraph(paragraph: ParagraphNode, offset: number): number {
  const length = paragraphText(paragraph).length;
  if (!Number.isFinite(offset)) {
    return 0;
  }
  return Math.max(0, Math.min(length, Math.round(offset)));
}

/**
 * Splits a selection into per-paragraph anchors.
 *
 * A range that spans paragraphs (or table cells) anchors every paragraph it
 * touches: the first and last paragraphs keep the selection offsets, the ones
 * in between are commented in full. That matches how Word stores a comment
 * spanning several paragraphs, and it keeps the offsets in the same space the
 * renderer and the gutter already use.
 */
export function resolveDocumentCommentAnchors(
  model: DocModel,
  range: DocxTextRange
): DocumentCommentAnchorRange[] {
  const normalized = normalizeTextRange(range);
  const locations = paragraphRangeForMutate(
    model,
    normalized.start.location,
    normalized.end.location
  );
  if (locations.length === 0) {
    return [];
  }

  const anchors: DocumentCommentAnchorRange[] = [];
  for (const entry of locations) {
    const location: ParagraphLocation = entry.location;
    const { paragraph } = getParagraphAtLocation(model, location);
    if (!paragraph) {
      continue;
    }

    const length = paragraphText(paragraph).length;
    const isFirst = sameParagraphLocation(location, normalized.start.location);
    const isLast = sameParagraphLocation(location, normalized.end.location);
    const firstOffset = isFirst
      ? clampToParagraph(paragraph, normalized.start.offset)
      : 0;
    const lastOffset = isLast
      ? clampToParagraph(paragraph, normalized.end.offset)
      : length;
    const start = Math.min(firstOffset, lastOffset);
    const end = Math.max(firstOffset, lastOffset);
    anchors.push({
      location,
      start,
      end,
      text: paragraphText(paragraph).slice(start, end),
    });
  }

  return anchors;
}

/**
 * Copy-on-write helper: replaces one top-level node with a deletable clone the
 * first time an anchor touches it, so several anchors in the same table all
 * land in the same revision.
 */
function cloneTopLevelNodeOnce(
  model: DocModel,
  nodeIndex: number,
  cloned: Set<number>
): void {
  if (cloned.has(nodeIndex)) {
    return;
  }
  const node = model.nodes[nodeIndex];
  if (!node) {
    return;
  }
  model.nodes[nodeIndex] = cloneDocNode(node);
  cloned.add(nodeIndex);
}

function attachAnchorRange(
  model: DocModel,
  anchor: DocumentCommentAnchorRange,
  commentId: number,
  cloned: Set<number>
): void {
  cloneTopLevelNodeOnce(
    model,
    anchor.location.kind === "table-cell"
      ? anchor.location.tableIndex
      : anchor.location.nodeIndex,
    cloned
  );

  const { paragraph } = getParagraphAtLocation(model, anchor.location);
  if (!paragraph) {
    return;
  }

  paragraph.commentRanges = mergeParagraphCommentRanges(
    [
      ...(paragraph.commentRanges ?? []),
      {
        commentId,
        start: anchor.start,
        end: anchor.end,
        anchorText: anchor.text.length > 0 ? anchor.text : undefined,
      },
    ],
    paragraphText(paragraph)
  );
}

function initialsFromAuthor(author: string): string {
  const parts = author
    .split(/[\s._-]+/)
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length === 0) {
    return "";
  }
  const initials = parts
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("");
  return initials || (author[0]?.toUpperCase() ?? "");
}

function withCommentDefinitions(
  model: DocModel,
  definitions: DocumentCommentDefinition[]
): DocModel {
  const next = cloneDocModelShallow(model);
  next.metadata = { ...model.metadata, comments: definitions };
  return next;
}

/**
 * Creates a comment anchored to `range` (one definition, one range per
 * paragraph the selection covers). Returns `undefined` for an empty body or a
 * range that resolves to no paragraph.
 */
export function addDocumentComment(
  model: DocModel,
  input: AddDocumentCommentInput
): DocumentCommentMutationResult | undefined {
  const text = input.text.trim();
  if (text.length === 0) {
    return undefined;
  }

  const anchors = resolveDocumentCommentAnchors(model, input.range);
  if (anchors.length === 0) {
    return undefined;
  }

  const commentId = input.id ?? nextDocumentCommentId(model);
  if (findDocumentComment(model, commentId)) {
    return undefined;
  }

  const author = input.author?.trim() || undefined;
  const definition: DocumentCommentDefinition = {
    id: commentId,
    text,
    author,
    initials: input.initials?.trim() || (author ? initialsFromAuthor(author) : undefined),
    date: input.date,
    parentId: undefined,
    resolved: undefined,
  };

  const next = withCommentDefinitions(model, [
    ...commentDefinitions(model),
    definition,
  ]);
  const cloned = new Set<number>();
  for (const anchor of anchors) {
    attachAnchorRange(next, anchor, commentId, cloned);
  }

  return { model: next, commentId };
}

/**
 * Appends a reply to the thread that owns `parentId`. Replies carry no anchor
 * of their own; the read path inherits the parent's location.
 */
export function replyToDocumentComment(
  model: DocModel,
  parentId: number,
  text: string,
  options: DocumentCommentAuthorOptions & { id?: number } = {}
): DocumentCommentMutationResult | undefined {
  const body = text.trim();
  if (body.length === 0 || !findDocumentComment(model, parentId)) {
    return undefined;
  }

  const commentId = options.id ?? nextDocumentCommentId(model);
  if (findDocumentComment(model, commentId)) {
    return undefined;
  }

  const author = options.author?.trim() || undefined;
  const definition: DocumentCommentDefinition = {
    id: commentId,
    text: body,
    author,
    initials: options.initials?.trim() || (author ? initialsFromAuthor(author) : undefined),
    date: options.date,
    parentId,
    resolved: undefined,
  };

  return {
    model: withCommentDefinitions(model, [
      ...commentDefinitions(model),
      definition,
    ]),
    commentId,
  };
}

/**
 * Marks the thread that owns `commentId` resolved (`w15:done="1"`) or reopens
 * it. Resolving stores the flag on the thread root, which is where Word keeps
 * it; reopening clears the flag on every member so a reply cannot keep a
 * reopened thread hidden.
 */
export function resolveDocumentComment(
  model: DocModel,
  commentId: number,
  resolved = true
): DocModel | undefined {
  const rootId = commentThreadRootId(model, commentId);
  if (rootId === undefined) {
    return undefined;
  }

  const memberIds = new Set(commentThreadMemberIds(model, rootId));
  let changed = false;
  const definitions = commentDefinitions(model).map((comment) => {
    const shouldChange = resolved
      ? comment.id === rootId && comment.resolved !== true
      : memberIds.has(comment.id) && comment.resolved !== undefined;
    if (!shouldChange) {
      return comment;
    }
    changed = true;
    return { ...comment, resolved: resolved ? true : undefined };
  });

  return changed ? withCommentDefinitions(model, definitions) : model;
}

/** Immutable strip of `commentIds` ranges from one node tree. */
function stripRangesFromNodes(
  nodes: readonly DocNode[],
  commentIds: ReadonlySet<number>
): DocNode[] | undefined {
  let changed = false;
  const next = nodes.map((node): DocNode => {
    if (node.type === "paragraph") {
      if (!node.commentRanges?.some((range) => commentIds.has(range.commentId))) {
        return node;
      }
      const ranges = node.commentRanges.filter(
        (range) => !commentIds.has(range.commentId)
      );
      changed = true;
      return {
        ...node,
        commentRanges: ranges.length > 0 ? ranges : undefined,
      };
    }

    let tableChanged = false;
    const rows = node.rows.map((row) => {
      const cells = row.cells.map((cell) => {
        const cellNodes = stripRangesFromNodes(cell.nodes, commentIds);
        if (!cellNodes) {
          return cell;
        }
        tableChanged = true;
        return { ...cell, nodes: cellNodes };
      });
      return tableChanged ? { ...row, cells } : row;
    });
    if (!tableChanged) {
      return node;
    }
    changed = true;
    return { ...node, rows };
  });

  return changed ? next : undefined;
}

/**
 * Deletes a comment. Deleting a thread root removes its replies as well;
 * deleting one reply removes that reply and anything replying to it. The
 * paragraph anchors for every removed definition go away with it.
 */
export function deleteDocumentComment(
  model: DocModel,
  commentId: number
): DocModel | undefined {
  if (!findDocumentComment(model, commentId)) {
    return undefined;
  }

  const removedIds = new Set(commentSubtreeIds(model, commentId));
  const next = cloneDocModelShallow(model);
  next.metadata = {
    ...model.metadata,
    comments: commentDefinitions(model).filter(
      (comment) => !removedIds.has(comment.id)
    ),
    headerSections: model.metadata.headerSections.map((section) => {
      const nodes = stripRangesFromNodes(section.nodes, removedIds);
      return nodes ? { ...section, nodes } : section;
    }),
    footerSections: model.metadata.footerSections.map((section) => {
      const nodes = stripRangesFromNodes(section.nodes, removedIds);
      return nodes ? { ...section, nodes } : section;
    }),
    sections: model.metadata.sections?.map((section) => ({
      ...section,
      headerSections: section.headerSections.map((header) => {
        const nodes = stripRangesFromNodes(header.nodes, removedIds);
        return nodes ? { ...header, nodes } : header;
      }),
      footerSections: section.footerSections.map((footer) => {
        const nodes = stripRangesFromNodes(footer.nodes, removedIds);
        return nodes ? { ...footer, nodes } : footer;
      }),
    })),
  };

  const bodyNodes = stripRangesFromNodes(next.nodes, removedIds);
  if (bodyNodes) {
    next.nodes = bodyNodes;
  }

  return next;
}
