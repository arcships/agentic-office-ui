// Document block extraction from the local DOCX model.
//
// Upstream `extractDocumentBlocks` delegates to the Rust/WASM parser
// (`wasmExtractDocumentBlocks`). This port derives the same JSON contract from
// `DocModel` directly: paragraphs, headings, lists, tables, cells, figures,
// header/footer containers and note paragraphs, each with a deterministic source
// id that the viewer renders back as `data-docx-source-id`.

import type {
  DocModel,
  DocNode,
  DocumentNoteDefinition,
  FooterSection,
  FormFieldRunNode,
  HeaderSection,
  ImageRunNode,
  ParagraphNode,
  TableCellContentNode,
  TableNode,
} from "../engine/types";
import { paragraphPageFieldSequence } from "../editor/helpers/field-helpers";
import { paragraphText } from "../editor/helpers/paragraph-inspect";
import { buildDocumentWords } from "./segmentation";
import {
  DOCX_BODY_PART_NAME,
  DOCX_ENDNOTES_PART_NAME,
  DOCX_FOOTNOTES_PART_NAME,
  docxCellParagraphSourceId,
  docxImageSourceId,
  docxNestedTableSourceId,
  docxNodeSourceId,
  docxNoteNodeSourceId,
  docxStoryContainerSourceId,
  docxStoryNodeSourceId,
  docxTableCellSourceId,
  docxTableRowSourceId,
} from "./ids";
import type {
  DocxBlockPart,
  DocxBlockPathSegment,
  DocxBlockStory,
  DocxSourceRef,
  DocumentBlock,
  DocumentBlockDetails,
  DocumentBlockType,
  ExtractDocumentBlocksOptions,
  StructuredDocument,
} from "./types";

export const DOCX_BLOCK_PARSER_VERSION = "arcships-docx-blocks/1";

const DEFAULT_DOCUMENT_ID = "docx-document";

interface ExtractionContext {
  documentId: string;
  part: DocxBlockPart;
  partName: string;
  story: DocxBlockStory;
  includeWords: boolean;
}

interface StoryTraversal {
  nodeIdFor: (nodeIndex: number) => string;
}

function paragraphStyleDetails(paragraph: ParagraphNode): DocumentBlockDetails["style"] {
  const style = paragraph.style;
  if (!style) return undefined;
  return {
    ...(style.styleId ? { styleId: style.styleId } : {}),
    ...(style.styleName ? { styleName: style.styleName } : {}),
    ...(style.headingLevel ? { headingLevel: style.headingLevel } : {}),
  };
}

function normalizedStyleName(paragraph: ParagraphNode): string {
  return (paragraph.style?.styleName ?? "").trim().toLowerCase();
}

function normalizedStyleId(paragraph: ParagraphNode): string {
  return (paragraph.style?.styleId ?? "").trim().toLowerCase();
}

function paragraphFormField(paragraph: ParagraphNode): FormFieldRunNode | undefined {
  return paragraph.children.find(
    (child): child is FormFieldRunNode => child.type === "form-field",
  );
}

function paragraphType(
  paragraph: ParagraphNode,
  formField: FormFieldRunNode | undefined,
): DocumentBlockType {
  if (paragraph.style?.headingLevel) return "heading";
  const styleName = normalizedStyleName(paragraph);
  const styleId = normalizedStyleId(paragraph);
  if (
    styleName === "title" ||
    styleName === "subtitle" ||
    styleId === "title" ||
    styleId === "subtitle"
  ) {
    return "section_heading";
  }
  if (styleName === "caption" || styleId === "caption") return "legend";
  if (formField) return "key_value";
  if (paragraphPageFieldSequence(paragraph).length > 0) return "page_number";
  return "text";
}

function numberingDetails(
  model: DocModel,
  paragraph: ParagraphNode,
): DocumentBlockDetails["list"] {
  const numbering = paragraph.style?.numbering;
  if (!numbering) return undefined;
  const definitions = model.metadata.numberingDefinitions;
  const instance = definitions?.instances.find(
    (candidate) => candidate.numId === numbering.numId,
  );
  const abstract = instance
    ? definitions?.abstracts.find(
        (candidate) => candidate.abstractNumId === instance.abstractNumId,
      )
    : undefined;
  const level = abstract?.levels.find(
    (candidate) => candidate.ilvl === numbering.ilvl,
  );
  return {
    numId: numbering.numId,
    level: numbering.ilvl,
    ...(level?.format ? { format: level.format } : {}),
  };
}

function imageBlocks(
  paragraph: ParagraphNode,
  paragraphSourceId: string,
  paragraphPath: readonly DocxBlockPathSegment[],
  parentBlockId: string | undefined,
  context: ExtractionContext,
  nodeIndex: number,
): DocumentBlock[] {
  const blocks: DocumentBlock[] = [];
  paragraph.children.forEach((child, childIndex) => {
    if (child.type !== "image" && child.type !== "drawing" && child.type !== "math") {
      return;
    }
    const figure =
      child.type === "drawing"
        ? { alt: child.alt ?? child.name ?? child.kind }
        : child.type === "math"
          ? { alt: child.text }
          : { alt: child.alt };
    const id = docxImageSourceId(paragraphSourceId, childIndex);
    blocks.push({
      object: "block",
      id,
      type: "figure",
      content: figure.alt ?? "",
      details: {
        kind: "figure",
        nodeIndex,
        childIndex,
        modelPath: paragraphPath,
        part: context.part,
        ...(figure.alt ? { alt: figure.alt } : {}),
      },
      sourceRefs: [
        {
          documentId: context.documentId,
          nodeId: id,
          partName: context.partName,
          story: context.story,
          offsetEncoding: "utf-16",
        },
      ],
      ...(parentBlockId ? { parentBlockId } : {}),
      children: [],
    });
  });
  return blocks;
}

function paragraphBlock(
  model: DocModel,
  paragraph: ParagraphNode,
  paragraphSourceId: string,
  paragraphPath: readonly DocxBlockPathSegment[],
  parentBlockId: string | undefined,
  context: ExtractionContext,
  nodeIndex: number,
  paragraphIndex: number | undefined,
): DocumentBlock[] {
  const content = paragraphText(paragraph);
  const formField = paragraphFormField(paragraph);
  const figures = imageBlocks(
    paragraph,
    paragraphSourceId,
    paragraphPath,
    undefined,
    context,
    nodeIndex,
  );

  // Empty paragraph wrappers promote their meaningful children (figures), as in
  // the upstream parser output.
  if (content.length === 0 && !formField) {
    return figures.map((figure) => ({ ...figure, parentBlockId }));
  }

  const type = paragraphType(paragraph, formField);
  const style = paragraphStyleDetails(paragraph);
  const list = numberingDetails(model, paragraph);
  const sourceRef: DocxSourceRef = {
    documentId: context.documentId,
    nodeId: paragraphSourceId,
    partName: context.partName,
    story: context.story,
    offsetEncoding: "utf-16",
  };
  const block: DocumentBlock = {
    object: "block",
    id: paragraphSourceId,
    type,
    content,
    details: {
      kind: "paragraph",
      nodeIndex,
      ...(paragraphIndex === undefined ? {} : { paragraphIndex }),
      part: context.part,
      modelPath: paragraphPath,
      ...(style ? { style } : {}),
      ...(list ? { list } : {}),
      ...(formField
        ? {
            formField: {
              type: formField.fieldType,
              ...(formField.value ? { value: formField.value } : {}),
            },
          }
        : {}),
      ...(paragraph.style?.headingLevel
        ? { headingLevel: paragraph.style.headingLevel }
        : {}),
      ...(paragraph.style?.dropCap ? { dropCap: true } : {}),
    },
    sourceRefs: [sourceRef],
    ...(parentBlockId ? { parentBlockId } : {}),
    children: figures.map((figure) => ({ ...figure, parentBlockId: paragraphSourceId })),
  };
  if (context.includeWords) {
    block.words = buildDocumentWords(paragraphSourceId, content, sourceRef);
  }
  return [block];
}

function nodeText(node: TableCellContentNode): string {
  if (node.type === "paragraph") return paragraphText(node);
  return tableSummary(node);
}

function tableSummary(table: TableNode): string {
  return table.rows
    .map((row) =>
      row.cells
        .map((cell) =>
          cell.nodes.map((node) => nodeText(node)).filter(Boolean).join("\n"),
        )
        .filter(Boolean)
        .join("\n"),
    )
    .filter(Boolean)
    .join("\n");
}

function tableColumnCount(table: TableNode): number {
  let columns = 0;
  for (const row of table.rows) {
    const rowColumns = row.cells.reduce(
      (total, cell) => total + Math.max(1, cell.style?.gridSpan ?? 1),
      0,
    );
    columns = Math.max(columns, rowColumns);
  }
  return columns;
}

function tableBlock(
  model: DocModel,
  table: TableNode,
  tableSourceId: string,
  tablePath: readonly DocxBlockPathSegment[],
  parentBlockId: string | undefined,
  context: ExtractionContext,
  nodeIndex: number,
  tableIndex: number,
): DocumentBlock {
  const children: DocumentBlock[] = [];

  table.rows.forEach((row, rowIndex) => {
    const rowSourceId = docxTableRowSourceId(tableSourceId, rowIndex);
    const rowPath: readonly DocxBlockPathSegment[] = [
      ...tablePath,
      { kind: "row", index: rowIndex },
    ];
    const cellBlocks: DocumentBlock[] = [];

    row.cells.forEach((cell, cellIndex) => {
      const cellSourceId = docxTableCellSourceId(tableSourceId, rowIndex, cellIndex);
      const cellPath: readonly DocxBlockPathSegment[] = [
        ...tablePath,
        { kind: "row", index: rowIndex },
        { kind: "cell", index: cellIndex },
      ];
      const cellChildren: DocumentBlock[] = [];
      let paragraphIndex = 0;

      cell.nodes.forEach((child, contentIndex) => {
        if (child.type === "paragraph") {
          cellChildren.push(
            ...paragraphBlock(
              model,
              child,
              docxCellParagraphSourceId(cellSourceId, paragraphIndex),
              [...cellPath, { kind: "paragraph", index: paragraphIndex }],
              cellSourceId,
              context,
              nodeIndex,
              paragraphIndex,
            ),
          );
          paragraphIndex += 1;
          return;
        }
        cellChildren.push(
          tableBlock(
            model,
            child,
            docxNestedTableSourceId(cellSourceId, contentIndex),
            [...cellPath, { kind: "table", index: contentIndex }],
            cellSourceId,
            context,
            nodeIndex,
            tableIndex,
          ),
        );
      });

      const content = cell.nodes.map((node) => nodeText(node)).filter(Boolean).join("\n");
      const hasFigure = cell.nodes.some(
        (node) =>
          node.type === "paragraph" &&
          node.children.some((child) => child.type === "image"),
      );
      const hasNestedTable = cell.nodes.some((node) => node.type === "table");
      if (content.length === 0 && !hasFigure && !hasNestedTable) return;

      cellBlocks.push({
        object: "block",
        id: cellSourceId,
        type: "table_cell",
        content,
        details: {
          kind: "table-cell",
          nodeIndex,
          tableIndex,
          rowIndex,
          cellIndex,
          part: context.part,
          modelPath: cellPath,
        },
        sourceRefs: [
          {
            documentId: context.documentId,
            nodeId: cellSourceId,
            partName: context.partName,
            story: context.story,
            offsetEncoding: "utf-16",
          },
        ],
        parentBlockId: row.style?.isHeader ? rowSourceId : tableSourceId,
        children: cellChildren,
      });
    });

    if (row.style?.isHeader) {
      children.push({
        object: "block",
        id: rowSourceId,
        type: "table_head",
        content: cellBlocks.map((cell) => cell.content).filter(Boolean).join("\n"),
        details: {
          kind: "table-row",
          nodeIndex,
          tableIndex,
          rowIndex,
          isHeaderRow: true,
          part: context.part,
          modelPath: rowPath,
        },
        sourceRefs: [
          {
            documentId: context.documentId,
            nodeId: rowSourceId,
            partName: context.partName,
            story: context.story,
            offsetEncoding: "utf-16",
          },
        ],
        parentBlockId: tableSourceId,
        children: cellBlocks,
      });
      return;
    }
    children.push(...cellBlocks);
  });

  return {
    object: "block",
    id: tableSourceId,
    type: "table",
    content: tableSummary(table),
    details: {
      kind: "table",
      nodeIndex,
      tableIndex,
      rowCount: table.rows.length,
      columnCount: tableColumnCount(table),
      part: context.part,
      modelPath: tablePath,
    },
    sourceRefs: [
      {
        documentId: context.documentId,
        nodeId: tableSourceId,
        partName: context.partName,
        story: context.story,
        offsetEncoding: "utf-16",
      },
    ],
    ...(parentBlockId ? { parentBlockId } : {}),
    children,
  };
}

function extractNodes(
  model: DocModel,
  nodes: readonly DocNode[],
  context: ExtractionContext,
  traversal: StoryTraversal,
  parentBlockId?: string,
): DocumentBlock[] {
  const blocks: DocumentBlock[] = [];
  nodes.forEach((node, nodeIndex) => {
    const sourceId = traversal.nodeIdFor(nodeIndex);
    if (node.type === "paragraph") {
      blocks.push(
        ...paragraphBlock(
          model,
          node,
          sourceId,
          [{ kind: "paragraph", index: nodeIndex }],
          parentBlockId,
          context,
          nodeIndex,
          undefined,
        ),
      );
      return;
    }
    blocks.push(
      tableBlock(
        model,
        node,
        sourceId,
        [{ kind: "table", index: nodeIndex }],
        parentBlockId,
        context,
        nodeIndex,
        nodeIndex,
      ),
    );
  });
  return blocks;
}

function storyContext(
  documentId: string,
  story: DocxBlockStory,
  part: DocxBlockPart,
  partName: string,
  includeWords: boolean,
): ExtractionContext {
  return { documentId, story, part, partName, includeWords };
}

function headerFooterBlocks(
  model: DocModel,
  section: HeaderSection | FooterSection,
  kind: "header" | "footer",
  documentId: string,
  includeWords: boolean,
): DocumentBlock[] {
  if (section.nodes.length === 0) return [];
  const containerId = docxStoryContainerSourceId(kind, section.partName);
  const context = storyContext(
    documentId,
    kind,
    kind === "header"
      ? { kind: "header", partName: section.partName }
      : { kind: "footer", partName: section.partName },
    section.partName,
    includeWords,
  );
  const children = extractNodes(
    model,
    section.nodes,
    context,
    {
      nodeIdFor: (nodeIndex) => docxStoryNodeSourceId(containerId, nodeIndex),
    },
    containerId,
  );
  return [{
    object: "block",
    id: containerId,
    type: kind,
    content: children.map((child) => child.content).filter(Boolean).join("\n"),
    details: {
      kind,
      part: context.part,
      modelPath: [],
    },
    sourceRefs: [
      {
        documentId,
        nodeId: containerId,
        partName: section.partName,
        story: kind,
        offsetEncoding: "utf-16",
      },
    ],
    children,
  }];
}

function noteBlocks(
  model: DocModel,
  note: DocumentNoteDefinition,
  kind: "footnote" | "endnote",
  documentId: string,
  includeWords: boolean,
): DocumentBlock[] {
  const partName =
    kind === "footnote" ? DOCX_FOOTNOTES_PART_NAME : DOCX_ENDNOTES_PART_NAME;
  const context = storyContext(
    documentId,
    kind,
    { kind, noteId: String(note.id) },
    partName,
    includeWords,
  );
  if (!note.nodes || note.nodes.length === 0) {
    // Models without structured note nodes still carry the note text; keep it
    // addressable instead of dropping the note from the block tree.
    if (!note.text.trim()) return [];
    const id = docxNoteNodeSourceId(kind, note.id, 0);
    const sourceRef: DocxSourceRef = {
      documentId,
      nodeId: id,
      partName,
      story: kind,
      offsetEncoding: "utf-16",
    };
    const block: DocumentBlock = {
      object: "block",
      id,
      type: "text",
      content: note.text,
      details: {
        kind: "paragraph",
        nodeIndex: 0,
        paragraphIndex: 0,
        part: context.part,
        modelPath: [{ kind: "paragraph", index: 0 }],
      },
      sourceRefs: [sourceRef],
      children: [],
    };
    if (includeWords) {
      block.words = buildDocumentWords(id, note.text, sourceRef);
    }
    return [block];
  }
  return extractNodes(model, note.nodes, context, {
    nodeIdFor: (nodeIndex) => docxNoteNodeSourceId(kind, note.id, nodeIndex),
  });
}

/**
 * Extracts the structured block tree from a model without touching a browser.
 * `extractDocumentBlocks` is the public Promise-shaped API; this synchronous
 * variant exists so callers that already hold a model can avoid an extra tick.
 */
export function extractDocumentBlocksSync(
  model: DocModel,
  options: ExtractDocumentBlocksOptions = {},
): StructuredDocument {
  const documentId = options.documentId?.trim() || DEFAULT_DOCUMENT_ID;
  const includeWords = options.includeWords === true;
  const blocks: DocumentBlock[] = [
    ...extractNodes(
      model,
      model.nodes,
      storyContext(
        documentId,
        "body",
        { kind: "body" },
        DOCX_BODY_PART_NAME,
        includeWords,
      ),
      { nodeIdFor: docxNodeSourceId },
    ),
  ];

  if (options.includeHeadersAndFooters !== false) {
    const seenParts = new Set<string>();
    for (const section of model.metadata.headerSections ?? []) {
      const key = `header:${section.partName}`;
      if (seenParts.has(key)) continue;
      seenParts.add(key);
      blocks.push(
        ...headerFooterBlocks(model, section, "header", documentId, includeWords),
      );
    }
    for (const section of model.metadata.footerSections ?? []) {
      const key = `footer:${section.partName}`;
      if (seenParts.has(key)) continue;
      seenParts.add(key);
      blocks.push(
        ...headerFooterBlocks(model, section, "footer", documentId, includeWords),
      );
    }
  }

  if (options.includeNotes !== false) {
    for (const note of model.metadata.footnotes ?? []) {
      blocks.push(...noteBlocks(model, note, "footnote", documentId, includeWords));
    }
    for (const note of model.metadata.endnotes ?? []) {
      blocks.push(
        ...noteBlocks(model, note, "endnote", documentId, includeWords),
      );
    }
  }

  return {
    schemaVersion: 1,
    parserVersion: DOCX_BLOCK_PARSER_VERSION,
    documentId,
    blocks,
    warnings: [...(model.metadata.warnings ?? [])],
    geometryStatus: "requires_layout",
  };
}

/**
 * Extracts document blocks and (optionally) word spans.
 *
 * The Promise shape matches the upstream WASM API so callers can swap imports;
 * the local implementation resolves in the same tick.
 */
export async function extractDocumentBlocks(
  model: DocModel,
  options: ExtractDocumentBlocksOptions = {},
): Promise<StructuredDocument> {
  return extractDocumentBlocksSync(model, options);
}

/** Depth-first flattening of the block tree, parents before children. */
export function flattenDocumentBlocks(
  blocks: readonly DocumentBlock[],
): DocumentBlock[] {
  return blocks.flatMap((block) => [
    block,
    ...flattenDocumentBlocks(block.children),
  ]);
}
