// Bookmark navigation index.
//
// Word stores an internal target as a `w:bookmarkStart` marker inside a body
// paragraph (or a table cell paragraph). Links to that target use
// `w:hyperlink/@w:anchor` or the `w:anchor` attribute of a HYPERLINK field,
// which the importer exposes as an `#name` href. This module turns the document
// model into the name -> node lookup navigation needs, so the renderer never
// has to re-walk the whole document while the reader clicks a table of
// contents.
//
// The names come from the same `BOOKMARK_START_XML_PATTERN` the REF/PAGEREF
// field resolvers use, keeping field references and link targets in agreement.

import type { DocModel, ParagraphNode } from "../../engine/types";
import { paragraphText } from "./paragraph-inspect";
import { paragraphBookmarkNames } from "./paragraph-tracked";

/** One resolvable `w:bookmarkStart` target in the body. */
export interface DocxBookmarkTarget {
  /** Bookmark name exactly as stored in `w:bookmarkStart/@w:name`. */
  name: string;
  /**
   * Body node index that owns the bookmark. A paragraph bookmark addresses the
   * paragraph node; a table-cell bookmark addresses the whole table node.
   */
  nodeIndex: number;
  /** Plain text of the paragraph carrying the bookmark. */
  text: string;
  /**
   * Cell coordinates when the bookmark sits in a direct table-cell paragraph.
   * Nested-table bookmarks keep the outer table index without cell details.
   */
  tableCell?: {
    rowIndex: number;
    cellIndex: number;
    paragraphIndex: number;
  };
}

/** Case-insensitive bookmark lookup over a document model. */
export interface DocxBookmarkIndex {
  /** Targets in document order, one per unique bookmark name. */
  readonly targets: readonly DocxBookmarkTarget[];
  /** Resolves `name` (optionally written as `#name`, URL-encoded). */
  get(name: string | null | undefined): DocxBookmarkTarget | undefined;
  has(name: string | null | undefined): boolean;
}

/**
 * Normalizes a bookmark reference for lookup.
 *
 * Strips the leading `#` an internal href carries, decodes percent-escapes
 * (`#My%20Bookmark`), trims, and lowercases because Word rejects duplicate
 * bookmark names that differ only by case.
 */
export function normalizeDocxBookmarkName(
  name: string | null | undefined
): string {
  let value = typeof name === "string" ? name.trim() : "";
  if (value.startsWith("#")) {
    value = value.slice(1);
  }
  try {
    value = decodeURIComponent(value);
  } catch {
    // A malformed escape is a literal name, not a reason to fail navigation.
  }
  return value.trim().toLowerCase();
}

function addBookmarkTarget(
  index: Map<string, DocxBookmarkTarget>,
  targets: DocxBookmarkTarget[],
  paragraph: ParagraphNode,
  nodeIndex: number,
  tableCell?: DocxBookmarkTarget["tableCell"]
): void {
  const names = paragraphBookmarkNames(paragraph);
  if (names.length === 0) {
    return;
  }
  const text = paragraphText(paragraph).trim();
  for (const name of names) {
    const key = normalizeDocxBookmarkName(name);
    if (!key || index.has(key)) {
      continue;
    }
    const target: DocxBookmarkTarget = { name, nodeIndex, text };
    if (tableCell) {
      target.tableCell = tableCell;
    }
    index.set(key, target);
    targets.push(target);
  }
}

/**
 * Builds the bookmark lookup for `model`.
 *
 * The first target wins when a document repeats a name (Word prevents that,
 * but a hand-built or repaired model can still contain it).
 */
export function createDocxBookmarkIndex(model: DocModel): DocxBookmarkIndex {
  const index = new Map<string, DocxBookmarkTarget>();
  const targets: DocxBookmarkTarget[] = [];

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      addBookmarkTarget(index, targets, node, nodeIndex);
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        cell.nodes.forEach((entry, paragraphIndex) => {
          if (entry.type !== "paragraph") {
            return;
          }
          addBookmarkTarget(index, targets, entry, nodeIndex, {
            rowIndex,
            cellIndex,
            paragraphIndex,
          });
        });
      });
    });
  });

  const get = (name: string | null | undefined): DocxBookmarkTarget | undefined => {
    const key = normalizeDocxBookmarkName(name);
    return key ? index.get(key) : undefined;
  };

  return {
    targets,
    get,
    has: (name: string | null | undefined): boolean => get(name) !== undefined,
  };
}
