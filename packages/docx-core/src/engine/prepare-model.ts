// Export pre-serialization.
//
// The serializer's JSON bridge renders a table from the model, but the model can
// describe two things the engine cannot write on its own:
//
// 1. A *collapsed* table grid — logical zero-width columns a merge collapsed.
//    Writing that grid verbatim emits a `w:tblGrid` that disagrees with the cell
//    spans. `projectCollapsedTableGrid` derives the physical grid first. A table
//    that still carries its stored markup is left alone: the engine writes that
//    markup verbatim, and only a rebuilt (edited) table needs projection.
// 2. Nested stories — headers, footers, sections, footnotes, endnotes and
//    comments. Each owns its own `DocNode[]`, so a table edited inside a header
//    or a note needs the same preparation as a body table.
//
// A grid that cannot be projected is an illegal model: the export fails with a
// typed `DocxSerializationError` instead of writing a package Word cannot open.
//
// Ported from upstream serializer/src/prepare-model.ts (commit 6f70b92).

import type {
  DocModel,
  DocNode,
  DocumentCommentDefinition,
  DocumentNoteDefinition,
  DocumentSection,
  ParagraphNode,
  TableNode
} from "./types";
import { projectCollapsedTableGrid } from "./table-grid-projection";

export type DocxSerializationErrorCode =
  | "COLLAPSED_TABLE_GRID_STYLE_CHANGED"
  | "COLLAPSED_TABLE_GRID_BORDER_SPACING"
  | "INVALID_COLLAPSED_TABLE_GRID";

/**
 * A model the serializer refuses to write. `code` is stable, `reason` carries
 * the projection rejection (for `INVALID_COLLAPSED_TABLE_GRID`).
 */
export class DocxSerializationError extends Error {
  readonly code: DocxSerializationErrorCode;
  readonly reason?: string;

  constructor(
    code: DocxSerializationErrorCode,
    message: string,
    reason?: string
  ) {
    super(message);
    this.name = "DocxSerializationError";
    this.code = code;
    this.reason = reason;
  }
}

function mapUnchanged<T>(values: T[], transform: (value: T) => T): T[] {
  let result: T[] | undefined;
  for (let index = 0; index < values.length; index += 1) {
    const value = transform(values[index]);
    if (value !== values[index]) {
      result ??= values.slice();
      result[index] = value;
    }
  }
  return result ?? values;
}

/**
 * Paragraph children in this model are runs; text-box content is stored as
 * `textBoxText`, so a paragraph has no nested `DocNode`s to prepare.
 */
function prepareParagraph(paragraph: ParagraphNode): ParagraphNode {
  return paragraph;
}

/**
 * A collapsed grid is only safe to project while the table still has the style
 * and border spacing it was imported with: a changed style can contribute its
 * own grid widths, and visible collapsed borders consume width from the
 * physical columns.
 */
function assertCollapsedGridInvariants(table: TableNode): void {
  const style = table.style ?? {};
  if (style.styleId !== style.sourceStyleId) {
    const from = style.sourceStyleId ? `"${style.sourceStyleId}"` : "no explicit style";
    const to = style.styleId ? `"${style.styleId}"` : "no explicit style";
    throw new DocxSerializationError(
      "COLLAPSED_TABLE_GRID_STYLE_CHANGED",
      `Collapsed table grid must be materialized before changing the table style (${from} → ${to}).`
    );
  }
  const hasVisibleTableBorder = Object.values(style.borders ?? {}).some(
    (border) => border && border.type !== "none" && border.type !== "nil"
  );
  if (hasVisibleTableBorder) {
    throw new DocxSerializationError(
      "COLLAPSED_TABLE_GRID_BORDER_SPACING",
      "Collapsed table grid requires confirmed zero border spacing before export."
    );
  }
}

function prepareTable(source: TableNode): TableNode {
  let table = source;
  // Only a rebuilt table needs the physical grid: a table that still carries
  // its stored markup is written verbatim by the engine, borders and all, so
  // the grid invariants do not apply to it.
  if (Array.isArray(source.style?.collapsedGridColumns) &&
      source.style.collapsedGridColumns.length > 0 &&
      source.sourceXml === undefined) {
    assertCollapsedGridInvariants(source);
    const projected = projectCollapsedTableGrid(source, {
      // The collapse marker is only carried by tables that keep cell spacing at
      // zero; `projectCollapsedTableGrid` re-checks the source spacing itself.
      resolvedBorderSpacingPx: 0
    });
    if (projected.kind === "rejected") {
      throw new DocxSerializationError(
        "INVALID_COLLAPSED_TABLE_GRID",
        `Invalid collapsed table grid: ${projected.reason}.`,
        projected.reason
      );
    }
    if (projected.kind === "projected") {
      table = projected.table;
    }
  }

  const rows = mapUnchanged(table.rows, (row) => {
    const cells = mapUnchanged(row.cells, (cell) => {
      const nodes = prepareNodes(cell.nodes);
      return nodes === cell.nodes ? cell : { ...cell, nodes };
    });
    return cells === row.cells ? row : { ...row, cells };
  });
  return rows === table.rows ? table : { ...table, rows };
}

function prepareNodes(nodes: DocNode[]): DocNode[] {
  return mapUnchanged(nodes, (node) =>
    node.type === "table" ? prepareTable(node) : prepareParagraph(node)
  );
}

function prepareStories<T extends { nodes: DocNode[] }>(stories: T[]): T[] {
  return mapUnchanged(stories, (story) => {
    const nodes = prepareNodes(story.nodes);
    return nodes === story.nodes ? story : { ...story, nodes };
  });
}

/**
 * Footnote/endnote definitions only carry `nodes` when the note body was
 * parsed; a text-only definition has nothing nested to prepare.
 */
function prepareOptionalStories<T extends { nodes?: DocNode[] }>(
  stories: T[] | undefined
): T[] | undefined {
  if (!stories) return stories;
  return mapUnchanged(stories, (story) => {
    if (!story.nodes) return story;
    const nodes = prepareNodes(story.nodes);
    return nodes === story.nodes ? story : { ...story, nodes };
  });
}

/**
 * Comments are anchors plus text in this model (`DocumentCommentDefinition`
 * has no story nodes); they are passed through to keep the returned model
 * reference stable when only the body changed.
 */
function prepareComments(
  comments: DocumentCommentDefinition[] | undefined
): DocumentCommentDefinition[] | undefined {
  return comments;
}

function prepareSections(
  sections: DocumentSection[] | undefined
): DocumentSection[] | undefined {
  if (!sections) return sections;
  return mapUnchanged(sections, (section) => {
    const headerSections = prepareStories(section.headerSections);
    const footerSections = prepareStories(section.footerSections);
    return headerSections === section.headerSections &&
      footerSections === section.footerSections
      ? section
      : { ...section, headerSections, footerSections };
  });
}

/**
 * Prepare a model for `modelToDocumentXml` / `serializeDocModel` /
 * `serializeDocx`: recurse into every story, then project each collapsed table
 * grid onto its physical columns.
 *
 * Returns the source model when nothing needed preparation, so callers can
 * cache the prepared model by reference.
 */
export function prepareModelForSerialization(model: DocModel): DocModel {
  const nodes = prepareNodes(model.nodes);
  const original = model.metadata;
  const headerSections = prepareStories(original.headerSections);
  const footerSections = prepareStories(original.footerSections);
  const sections = prepareSections(original.sections);
  const footnotes: DocumentNoteDefinition[] | undefined = prepareOptionalStories(
    original.footnotes
  );
  const endnotes: DocumentNoteDefinition[] | undefined = prepareOptionalStories(
    original.endnotes
  );
  const comments = prepareComments(original.comments);
  const metadata =
    headerSections === original.headerSections &&
    footerSections === original.footerSections &&
    sections === original.sections &&
    footnotes === original.footnotes &&
    endnotes === original.endnotes &&
    comments === original.comments
      ? original
      : {
          ...original,
          headerSections,
          footerSections,
          sections,
          footnotes,
          endnotes,
          comments
        } as DocModel["metadata"];
  return nodes === model.nodes && metadata === original
    ? model
    : { ...model, nodes, metadata };
}
