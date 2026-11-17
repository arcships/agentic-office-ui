// Comment anchor cache.
//
// Comment locations are discovered from `w:commentReference` /
// `w:commentRangeStart` markup on first collection. Editing a paragraph clears
// its `sourceXml`, so the cache keeps the discovered location alive across
// `cloneDocModel()` revisions without adding parser-owned fields to the public
// document model.
//
// The anchor is stored as a non-enumerable symbol property on the comment
// definition. A `WeakMap` alone is not enough here: Vue wraps the model in a
// reactive proxy, and a proxy is a different WeakMap key than its raw target.
// Symbol property reads pass through the proxy, and `cloneDocModel` transfers
// the anchor to the cloned definition objects explicitly.

import type {
  DocumentCommentAnchor,
  DocumentCommentDefinition,
} from "./types";

const DOCUMENT_COMMENT_ANCHOR = Symbol("documentCommentAnchor");

type DocumentCommentDefinitionWithAnchor = DocumentCommentDefinition & {
  [DOCUMENT_COMMENT_ANCHOR]?: DocumentCommentAnchor;
};

export function cloneDocumentCommentAnchor(
  anchor: DocumentCommentAnchor
): DocumentCommentAnchor {
  return { ...anchor };
}

export function getCachedDocumentCommentAnchor(
  definition: DocumentCommentDefinition
): DocumentCommentAnchor | undefined {
  return (definition as DocumentCommentDefinitionWithAnchor)[
    DOCUMENT_COMMENT_ANCHOR
  ];
}

export function setCachedDocumentCommentAnchor(
  definition: DocumentCommentDefinition,
  anchor: DocumentCommentAnchor
): void {
  const target = definition as DocumentCommentDefinitionWithAnchor;
  if (target[DOCUMENT_COMMENT_ANCHOR] === anchor) {
    return;
  }
  try {
    Object.defineProperty(target, DOCUMENT_COMMENT_ANCHOR, {
      value: anchor,
      configurable: true,
      enumerable: false,
      writable: true,
    });
  } catch {
    // Frozen definitions cannot carry a hidden anchor; callers fall back to
    // re-reading sourceXml when available.
  }
}

export function clearCachedDocumentCommentAnchor(
  definition: DocumentCommentDefinition
): void {
  delete (definition as DocumentCommentDefinitionWithAnchor)[
    DOCUMENT_COMMENT_ANCHOR
  ];
}
