// Stage 2–4 object index for DOCX.
//
// The index is the one place that turns a `DocModel` into user-facing objects:
// visible content (paragraph, heading, list item, table, image, comment), the
// table → row → cell hierarchy (stage 3) and the invisible behaviors stage 4
// exposes as reference kinds. Every entry is a descriptor with a draft,
// fingerprint and the four reliability dimensions, so a Surface only has to
// forward candidates.

import type {
  OfficeObjectDescriptor,
  OfficeObjectKind,
  OfficeObjectProvider,
  OfficeObjectRegistry,
} from "@arcships/office-interaction";
import {
  createOfficeObjectDescriptor,
  createOfficeObjectRegistry,
  officeObjectReliability,
} from "@arcships/office-interaction";
import type {
  DocxReferenceContext,
} from "./reference-adapter";
import {
  createDocxBehaviorDraftFromRecord,
  createDocxBlockReferenceDraft,
  createDocxCommentReferenceDraft,
  docxBlockReferenceKind,
} from "./reference-adapter";
import type {
  DocumentBlock,
  DocxBlockPathSegment,
  StructuredDocument,
} from "./blocks/types";
import { extractDocumentBlocksSync } from "./blocks/extract";
import type { DocModel, DocNode, ParagraphNode, TableNode } from "./engine/types";
import { paragraphText } from "./editor/helpers/paragraph-inspect";
import { DOCX_BEHAVIOR_REFERENCE_KIND, enumerateDocxBehaviors } from "./docx-behaviors";

export const DOCX_STAGE2_BLOCK_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "paragraph",
  "heading",
  "list-item",
  "table",
  "table-row",
  "table-cell",
  "image",
  "section",
]);

export const DOCX_STAGE4_BEHAVIOR_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "style",
  "field",
  "bookmark",
  "section",
  "tracked-change",
  "comment",
  "cross-reference",
]);

const BLOCK_OBJECT_PREFIX = "docx:block:";
const COMMENT_OBJECT_PREFIX = "docx:comment:";
const BEHAVIOR_OBJECT_PREFIX = "docx:behavior:";

/** Reuse one extraction per model instance; edits replace the model object. */
const structuredCache = new WeakMap<DocModel, StructuredDocument>();

export function docxStructuredDocument(model: DocModel): StructuredDocument {
  const cached = structuredCache.get(model);
  if (cached) return cached;
  const extracted = extractDocumentBlocksSync(model, {
    includeWords: false,
    includeHeadersAndFooters: true,
    includeNotes: true,
  });
  structuredCache.set(model, extracted);
  return extracted;
}

function modelPathKey(path: readonly DocxBlockPathSegment[] | undefined): string {
  return (path ?? []).map((segment) => `${segment.kind}:${segment.index}`).join("/");
}

function compactLabel(block: DocumentBlock): string {
  return block.content.replace(/\s+/gu, " ").trim();
}

function docxTextHash(text: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    hash ^= code & 0xff;
    hash = Math.imul(hash, 0x01000193);
    hash ^= code >>> 8;
    hash = Math.imul(hash, 0x01000193);
  }
  return `fnv1a32:${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

function tableRowObjectId(tablePath: readonly DocxBlockPathSegment[], rowIndex: number): string {
  return `docx:table-row:${modelPathKey(tablePath)}/${rowIndex}`;
}

function lastPathIndex(
  path: readonly DocxBlockPathSegment[] | undefined,
  kind: DocxBlockPathSegment["kind"],
): number | undefined {
  const segments = path ?? [];
  for (let index = segments.length - 1; index >= 0; index -= 1) {
    if (segments[index]!.kind === kind) return segments[index]!.index;
  }
  return undefined;
}

function storyNodesForPart(model: DocModel, part: DocumentBlock["details"]["part"]): readonly DocNode[] | undefined {
  if (!part || part.kind === "body") return model.nodes;
  if (part.kind === "header" || part.kind === "footer") {
    const flat = part.kind === "header"
      ? model.metadata.headerSections
      : model.metadata.footerSections;
    const direct = flat?.find((section) => section.partName === part.partName);
    if (direct) return direct.nodes;
    for (const section of model.metadata.sections ?? []) {
      const nested = (part.kind === "header"
        ? section.headerSections
        : section.footerSections)?.find((entry) => entry.partName === part.partName);
      if (nested) return nested.nodes;
    }
    return undefined;
  }
  if (part.kind === "footnote" || part.kind === "endnote") {
    const notes = part.kind === "footnote"
      ? model.metadata.footnotes
      : model.metadata.endnotes;
    return notes?.find((note) => String(note.id) === part.noteId)?.nodes;
  }
  return undefined;
}

/** Resolve a table node from a structural path that starts with a table segment. */
function tableNodeAtPath(
  nodes: readonly DocNode[],
  path: readonly DocxBlockPathSegment[],
  index = 0,
): TableNode | undefined {
  const tableSegment = path[index];
  if (!tableSegment || tableSegment.kind !== "table") return undefined;
  const node = nodes[tableSegment.index];
  if (node?.type !== "table") return undefined;
  if (index === path.length - 1) return node;
  const rowSegment = path[index + 1];
  const cellSegment = path[index + 2];
  const nestedSegment = path[index + 3];
  if (!rowSegment || rowSegment.kind !== "row" || !cellSegment || cellSegment.kind !== "cell"
    || !nestedSegment || nestedSegment.kind !== "table") {
    return undefined;
  }
  const cell = node.rows[rowSegment.index]?.cells[cellSegment.index];
  return cell ? tableNodeAtPath(cell.nodes, path, index + 3) : undefined;
}

function rowDescriptor(
  context: DocxReferenceContext,
  part: DocumentBlock["details"]["part"],
  tablePath: readonly DocxBlockPathSegment[],
  rowIndex: number,
  tableObjectId: string | undefined,
  rowText: string,
): OfficeObjectDescriptor {
  const path = [...tablePath, { kind: "row" as const, index: rowIndex }];
  const tableIndex = lastPathIndex(tablePath, "table") ?? 0;
  const pathParts: { kind: OfficeObjectKind; label: string }[] = [
    { kind: "document", label: "Document" },
    ...(part?.kind === "header"
      ? [{ kind: "section" as const, label: `Header ${part.partName}` }]
      : part?.kind === "footer"
        ? [{ kind: "section" as const, label: `Footer ${part.partName}` }]
        : part?.kind === "footnote" || part?.kind === "endnote"
          ? [{ kind: "note" as const, label: `${part.kind === "footnote" ? "Footnote" : "Endnote"} ${part.noteId}` }]
          : []),
  ];
  for (const segment of tablePath) {
    if (segment.kind === "table") pathParts.push({ kind: "table", label: `Table ${segment.index + 1}` });
    else if (segment.kind === "row") pathParts.push({ kind: "table-row", label: `Row ${segment.index + 1}` });
    else if (segment.kind === "cell") pathParts.push({ kind: "table-cell", label: `Cell ${segment.index + 1}` });
  }
  pathParts.push({ kind: "table-row", label: `Row ${rowIndex + 1}` });
  return createOfficeObjectDescriptor({
    objectId: tableRowObjectId(tablePath, rowIndex),
    draft: {
      schemaVersion: 1,
      document: context.revision,
      kind: "table-row",
      source: "native",
      locator: {
        type: "format",
        format: "docx",
        value: {
          kind: "structure",
          part: part ?? { kind: "body" },
          path,
        },
      },
      fingerprint: {
        ...(rowText ? { contentHash: docxTextHash(rowText) } : {}),
        ancestorKeys: [
          `docx.part:${part?.kind ?? "body"}`,
          ...tablePath.map((segment) => `docx.${segment.kind}:${segment.index}`),
          `docx.row:${rowIndex}`,
        ],
      },
      reliability: officeObjectReliability(
        { level: "exact", score: 1, reasonCodes: ["docx.native-table"] },
        { level: "exact", score: 1, reasonCodes: ["docx.block-path"] },
        { level: "exact", score: 1, reasonCodes: ["docx.model-path"] },
        { level: "likely", reasonCodes: ["docx.structural-path"] },
      ),
    },
    label: `Row ${rowIndex + 1}`,
    path: pathParts,
    ...(tableObjectId === undefined ? {} : { parentObjectId: tableObjectId }),
    childrenState: "available",
    facets: {
      tableIndex,
      rowIndex,
      part: part?.kind ?? "body",
      path: modelPathKey(path),
    },
  });
}

function blockPath(block: DocumentBlock): readonly { kind: OfficeObjectKind; label: string }[] {
  const part = block.details.part;
  const path: { kind: OfficeObjectKind; label: string }[] = [
    { kind: "document", label: "Document" },
  ];
  if (part?.kind === "header") path.push({ kind: "section", label: `Header ${part.partName}` });
  else if (part?.kind === "footer") path.push({ kind: "section", label: `Footer ${part.partName}` });
  else if (part?.kind === "footnote" || part?.kind === "endnote") {
    path.push({
      kind: "note",
      label: `${part.kind === "footnote" ? "Footnote" : "Endnote"} ${part.noteId}`,
    });
  }
  const kind = docxBlockReferenceKind(block);
  const modelPath = block.details.modelPath ?? [];
  for (const segment of modelPath) {
    if (segment.kind === "table") path.push({ kind: "table", label: `Table ${segment.index + 1}` });
    else if (segment.kind === "row") path.push({ kind: "table-row", label: `Row ${segment.index + 1}` });
    else if (segment.kind === "cell") path.push({ kind: "table-cell", label: `Cell ${segment.index + 1}` });
  }
  if (block.type === "table" || block.type === "table_head" || block.type === "table_cell") {
    return path;
  }
  path.push({ kind, label: blockLabel(block, kind, compactLabel(block)) });
  return path;
}

function blockLabel(block: DocumentBlock, kind: OfficeObjectKind, compact: string): string {
  if (kind === "table") {
    const index = lastPathIndex(block.details.modelPath, "table") ?? block.details.tableIndex;
    return index === undefined ? "Table" : `Table ${index + 1}`;
  }
  if (kind === "table-row") {
    const rowIndex = lastPathIndex(block.details.modelPath, "row") ?? block.details.rowIndex;
    return rowIndex === undefined ? "Row" : `Row ${rowIndex + 1}`;
  }
  if (kind === "table-cell") {
    const cellIndex = lastPathIndex(block.details.modelPath, "cell") ?? block.details.cellIndex;
    const rowIndex = lastPathIndex(block.details.modelPath, "row") ?? block.details.rowIndex;
    if (cellIndex === undefined) return "Cell";
    return rowIndex === undefined
      ? `Cell ${cellIndex + 1}`
      : `Cell ${rowIndex + 1}, ${cellIndex + 1}`;
  }
  if (kind === "image") {
    const alt = block.details.alt?.replace(/\s+/gu, " ").trim();
    return alt ? `Image: ${alt.slice(0, 80)}` : "Image";
  }
  if (kind === "heading") return compact ? `Heading: ${compact.slice(0, 80)}` : "Heading";
  if (kind === "list-item") return compact ? `List item: ${compact.slice(0, 80)}` : "List item";
  if (kind === "section") {
    const part = block.details.part;
    return part?.kind === "header" ? `Header ${part.partName}` : part?.kind === "footer" ? `Footer ${part.partName}` : "Section";
  }
  return compact ? `Text: ${compact.slice(0, 80)}` : "Paragraph";
}

function blockReliability(kind: OfficeObjectKind) {
  if (kind === "table" || kind === "table-row" || kind === "table-cell") {
    return officeObjectReliability(
      { level: "exact", score: 1, reasonCodes: ["docx.native-table"] },
      { level: "exact", score: 1, reasonCodes: ["docx.block-path"] },
      { level: "exact", score: 1, reasonCodes: ["docx.model-path"] },
      { level: "likely", reasonCodes: ["docx.structural-path"] },
    );
  }
  if (kind === "image") {
    return officeObjectReliability(
      { level: "exact", score: 1, reasonCodes: ["docx.native-image"] },
      { level: "exact", score: 1, reasonCodes: ["docx.model-child-index"] },
      { level: "likely", reasonCodes: ["docx.model-path"] },
      { level: "likely", reasonCodes: ["docx.structural-path"] },
    );
  }
  // Paragraph-derived objects are structural: Word records styles and
  // numbering, not a stable paragraph identity.
  return officeObjectReliability(
    { level: "likely", reasonCodes: ["docx.structural-paragraph"] },
    { level: "exact", score: 1, reasonCodes: ["docx.block-source-ref"] },
    { level: "exact", score: 1, reasonCodes: ["docx.model-path"] },
    { level: "likely", reasonCodes: ["docx.structural-path"] },
  );
}

function blockDescriptor(
  context: DocxReferenceContext,
  block: DocumentBlock,
  parentId: string | undefined,
): OfficeObjectDescriptor | undefined {
  if (!block.sourceRefs.length) return undefined;
  let draft;
  try {
    draft = createDocxBlockReferenceDraft(context, block);
  } catch {
    return undefined;
  }
  const kind = docxBlockReferenceKind(block);
  const compact = block.content.replace(/\s+/gu, " ").trim();
  return createOfficeObjectDescriptor({
    objectId: `${BLOCK_OBJECT_PREFIX}${block.id}`,
    draft: { ...draft, reliability: blockReliability(kind) },
    label: blockLabel(block, kind, compact),
    path: blockPath(block),
    ...(parentId === undefined ? {} : { parentObjectId: parentId }),
    childrenState: block.children.length > 0 ? "available" : "none",
    ...(block.content ? { content: { text: block.content } } : {}),
    facets: {
      blockType: block.type,
      blockId: block.id,
      ...(block.details.nodeIndex === undefined ? {} : { nodeIndex: block.details.nodeIndex }),
      ...(block.details.tableIndex === undefined ? {} : { tableIndex: block.details.tableIndex }),
      ...(block.details.rowIndex === undefined ? {} : { rowIndex: block.details.rowIndex }),
      ...(block.details.cellIndex === undefined ? {} : { cellIndex: block.details.cellIndex }),
      part: block.details.part?.kind ?? "body",
      path: modelPathKey(block.details.modelPath),
    },
  });
}

function commentAnchorLocation(model: DocModel, commentId: number):
  | { kind: "paragraph"; nodeIndex: number }
  | { kind: "table-cell"; tableIndex: number; rowIndex: number; cellIndex: number; paragraphIndex: number }
  | undefined {
  const matches = (paragraph: ParagraphNode): boolean =>
    paragraph.commentRanges?.some((range) => range.commentId === commentId) === true;
  for (const [nodeIndex, node] of model.nodes.entries()) {
    if (node.type === "paragraph") {
      if (matches(node)) return { kind: "paragraph", nodeIndex };
      continue;
    }
    for (const [rowIndex, row] of node.rows.entries()) {
      for (const [cellIndex, cell] of row.cells.entries()) {
        for (const [paragraphIndex, paragraph] of cell.nodes.entries()) {
          if (paragraph.type === "paragraph" && matches(paragraph)) {
            return { kind: "table-cell", tableIndex: nodeIndex, rowIndex, cellIndex, paragraphIndex };
          }
        }
      }
    }
  }
  return undefined;
}

function anchorPath(anchor: ReturnType<typeof commentAnchorLocation>): readonly DocxBlockPathSegment[] {
  if (!anchor) return [];
  if (anchor.kind === "paragraph") return [{ kind: "paragraph", index: anchor.nodeIndex }];
  return [
    { kind: "table", index: anchor.tableIndex },
    { kind: "row", index: anchor.rowIndex },
    { kind: "cell", index: anchor.cellIndex },
    { kind: "paragraph", index: anchor.paragraphIndex },
  ];
}

function commentDescriptor(
  context: DocxReferenceContext,
  model: DocModel,
  comment: NonNullable<DocModel["metadata"]["comments"]>[number],
  blockIdByPath: ReadonlyMap<string, string>,
): OfficeObjectDescriptor | undefined {
  const anchor = commentAnchorLocation(model, comment.id);
  let draft;
  try {
    draft = createDocxCommentReferenceDraft(context, {
      commentId: String(comment.id),
      ...(anchor ? { anchor } : {}),
    });
  } catch {
    return undefined;
  }
  const anchorBlock = anchor ? blockIdByPath.get(modelPathKey(anchorPath(anchor))) : undefined;
  const text = comment.text.replace(/\s+/gu, " ").trim();
  const anchorPathEntries: { kind: OfficeObjectKind; label: string }[] = anchor === undefined
    ? []
    : anchor.kind === "paragraph"
      ? [{ kind: "paragraph", label: `Paragraph ${anchor.nodeIndex + 1}` }]
      : [
          { kind: "table", label: `Table ${anchor.tableIndex + 1}` },
          { kind: "table-row", label: `Row ${anchor.rowIndex + 1}` },
          { kind: "table-cell", label: `Cell ${anchor.cellIndex + 1}` },
          { kind: "paragraph", label: `Paragraph ${anchor.paragraphIndex + 1}` },
        ];
  return createOfficeObjectDescriptor({
    objectId: `${COMMENT_OBJECT_PREFIX}${comment.id}`,
    draft,
    label: text ? `Comment: ${text.slice(0, 80)}` : `Comment ${comment.id}`,
    path: [
      { kind: "document", label: "Document" },
      ...anchorPathEntries,
      { kind: "comment", label: comment.author ? `Comment by ${comment.author}` : "Comment" },
    ],
    ...(anchorBlock === undefined ? {} : { parentObjectId: `${BLOCK_OBJECT_PREFIX}${anchorBlock}` }),
    childrenState: "none",
    content: { text: comment.text },
    facets: {
      commentId: String(comment.id),
      ...(comment.author ? { author: comment.author } : {}),
      resolved: comment.resolved === true,
      ...(comment.parentId === undefined ? {} : { parentId: String(comment.parentId) }),
    },
  });
}

function isBehaviorDescriptor(descriptor: OfficeObjectDescriptor): boolean {
  return descriptor.draft.locator.type === "format"
    && descriptor.draft.locator.value.kind === "behavior";
}

function behaviorDescriptor(
  context: DocxReferenceContext,
  record: ReturnType<typeof enumerateDocxBehaviors>[number],
): OfficeObjectDescriptor {
  const draft = createDocxBehaviorDraftFromRecord(context, record);
  return createOfficeObjectDescriptor({
    objectId: `${BEHAVIOR_OBJECT_PREFIX}${record.instanceId}`,
    draft,
    label: record.label,
    path: record.path,
    childrenState: "none",
    ...(record.content === undefined ? {} : { content: record.content }),
    facets: {
      ...record.facets,
      behavior: record.behavior,
      instanceId: record.instanceId,
      referenceKind: DOCX_BEHAVIOR_REFERENCE_KIND[record.behavior],
    },
  });
}

export interface DocxObjectIndexOptions {
  /** Include header/footer and note story blocks. Defaults to true. */
  includeStoryParts?: boolean;
  /** Include stage 4 invisible behaviors. Defaults to true. */
  includeBehaviors?: boolean;
}

export interface DocxObjectIndex {
  descriptors: readonly OfficeObjectDescriptor[]
  byObjectId: ReadonlyMap<string, OfficeObjectDescriptor>
  blockByObjectId: ReadonlyMap<string, DocumentBlock>
  blockBySourceId: ReadonlyMap<string, DocumentBlock>
  objectIdByBlockId: ReadonlyMap<string, string>
  /** Visible objects only, in document order. */
  visible: readonly OfficeObjectDescriptor[]
  /** Stage 4 behaviors (no DOM outline). */
  behaviors: readonly OfficeObjectDescriptor[]
}

function buildDocxObjectIndex(
  context: DocxReferenceContext,
  options: DocxObjectIndexOptions = {},
): DocxObjectIndex {
  const structured = docxStructuredDocument(context.model);
  const includeStoryParts = options.includeStoryParts !== false;
  const blocks: DocumentBlock[] = [];
  const visit = (block: DocumentBlock): void => {
    const part = block.details.part?.kind ?? "body";
    const storyAllowed = includeStoryParts || part === "body";
    if (storyAllowed && DOCX_STAGE2_BLOCK_KINDS.includes(docxBlockReferenceKind(block))) {
      blocks.push(block);
    }
    for (const child of block.children) visit(child);
  };
  for (const block of structured.blocks) visit(block);

  const objectIdByBlockId = new Map<string, string>();
  for (const block of blocks) objectIdByBlockId.set(block.id, `${BLOCK_OBJECT_PREFIX}${block.id}`);
  const blockIdByPath = new Map<string, string>();
  for (const block of blocks) blockIdByPath.set(modelPathKey(block.details.modelPath), block.id);

  // Rows are not separate extraction blocks except for repeating header rows.
  // Build the row map from each table block's own structural path so nested
  // tables keep `table -> row -> cell` inside the cell that owns them instead
  // of being re-parented to a top-level row with the same numeric index.
  const rowBlockByPath = new Map<string, DocumentBlock>();
  for (const block of blocks) {
    if (block.type === "table_head") {
      rowBlockByPath.set(modelPathKey(block.details.modelPath), block);
    }
  }
  const tableBlocks = blocks.filter((block) => block.type === "table");
  const tableObjectIdByPath = new Map<string, string>();
  for (const block of tableBlocks) {
    tableObjectIdByPath.set(modelPathKey(block.details.modelPath), `${BLOCK_OBJECT_PREFIX}${block.id}`);
  }
  const rowObjectIdByPath = new Map<string, string>();
  const tableRows = new Map<DocumentBlock, { node: TableNode; rows: { index: number; text: string }[] }>();
  for (const block of tableBlocks) {
    const tablePath = block.details.modelPath ?? [];
    const nodes = storyNodesForPart(context.model, block.details.part);
    const node = nodes ? tableNodeAtPath(nodes, tablePath) : undefined;
    if (!node) continue;
    const rows = node.rows.map((row, index) => ({
      index,
      text: row.cells
        .map((cell) => cell.nodes
          .map((entry) => (entry.type === "paragraph" ? paragraphText(entry) : ""))
          .filter(Boolean)
          .join(" "))
        .filter(Boolean)
        .join("\t"),
    }));
    tableRows.set(block, { node, rows });
    for (const row of rows) {
      const rowPath = [...tablePath, { kind: "row" as const, index: row.index }];
      const key = modelPathKey(rowPath);
      const rowBlock = rowBlockByPath.get(key);
      rowObjectIdByPath.set(
        key,
        rowBlock ? `${BLOCK_OBJECT_PREFIX}${rowBlock.id}` : tableRowObjectId(tablePath, row.index),
      );
    }
  }
  const rowDescriptors: OfficeObjectDescriptor[] = [];
  for (const block of tableBlocks) {
    const tablePath = block.details.modelPath ?? [];
    const tableRowsEntry = tableRows.get(block);
    if (!tableRowsEntry) continue;
    const tableObjectId = `${BLOCK_OBJECT_PREFIX}${block.id}`;
    for (const row of tableRowsEntry.rows) {
      const rowPath = [...tablePath, { kind: "row" as const, index: row.index }];
      if (rowBlockByPath.has(modelPathKey(rowPath))) continue;
      rowDescriptors.push(rowDescriptor(
        context,
        block.details.part,
        tablePath,
        row.index,
        tableObjectId,
        row.text,
      ));
    }
  }

  const descriptors: OfficeObjectDescriptor[] = [...rowDescriptors];
  const blockByObjectId = new Map<string, DocumentBlock>();
  const blockBySourceId = new Map<string, DocumentBlock>();
  for (const block of blocks) {
    const blockPathSegments = block.details.modelPath ?? [];
    let parentId: string | undefined;
    if (block.type === "table_cell") {
      parentId = rowObjectIdByPath.get(modelPathKey(blockPathSegments.slice(0, -1)));
    } else if (block.type === "table_head") {
      parentId = tableObjectIdByPath.get(modelPathKey(blockPathSegments.slice(0, -1)));
    } else if (block.parentBlockId !== undefined) {
      parentId = objectIdByBlockId.get(block.parentBlockId);
    }
    const descriptor = blockDescriptor(context, block, parentId);
    if (!descriptor) continue;
    descriptors.push(descriptor);
    blockByObjectId.set(descriptor.objectId, block);
    for (const ref of block.sourceRefs) blockBySourceId.set(ref.nodeId, block);
  }

  for (const comment of context.model.metadata.comments ?? []) {
    const descriptor = commentDescriptor(context, context.model, comment, blockIdByPath);
    if (descriptor) descriptors.push(descriptor);
  }

  if (options.includeBehaviors !== false) {
    for (const record of enumerateDocxBehaviors(context.model)) {
      descriptors.push(behaviorDescriptor(context, record));
    }
  }

  const byObjectId = new Map(descriptors.map((descriptor) => [descriptor.objectId, descriptor]));
  return {
    descriptors: Object.freeze(descriptors),
    byObjectId,
    blockByObjectId,
    blockBySourceId,
    objectIdByBlockId,
    // DOCX core has no layout geometry, so "visible" means a canvas object
    // rather than a descriptor that already carries fragments. `comment` and
    // `section` exist in both stages, so the behavior locator is the
    // discriminator, not the reference kind.
    visible: Object.freeze(descriptors.filter((descriptor) => !isBehaviorDescriptor(descriptor))),
    behaviors: Object.freeze(descriptors.filter(isBehaviorDescriptor)),
  };
}

const indexCache = new WeakMap<DocModel, Map<string, DocxObjectIndex>>();

/** Build (and memoize per model + options) the DOCX object index. */
export function createDocxObjectIndex(
  context: DocxReferenceContext,
  options: DocxObjectIndexOptions = {},
): DocxObjectIndex {
  const key = `${context.revision.documentId}:${context.revision.revision}:${options.includeStoryParts !== false}:${options.includeBehaviors !== false}`;
  let entries = indexCache.get(context.model);
  if (!entries) {
    entries = new Map();
    indexCache.set(context.model, entries);
  }
  const cached = entries.get(key);
  if (cached) return cached;
  const index = buildDocxObjectIndex(context, options);
  if (entries.size > 4) entries.clear();
  entries.set(key, index);
  return index;
}

/** Default providers; hosts can register additional families. */
export const DOCX_BLOCK_OBJECT_PROVIDER: OfficeObjectProvider<DocxReferenceContext> = {
  id: "docx.blocks",
  format: "docx",
  kinds: DOCX_STAGE2_BLOCK_KINDS,
  list(context, options) {
    const index = createDocxObjectIndex(context);
    return index.descriptors.filter((descriptor) =>
      DOCX_STAGE2_BLOCK_KINDS.includes(descriptor.draft.kind)
      && (options?.kinds === undefined || options.kinds.length === 0 || options.kinds.includes(descriptor.draft.kind)));
  },
};

export const DOCX_COMMENT_OBJECT_PROVIDER: OfficeObjectProvider<DocxReferenceContext> = {
  id: "docx.comments",
  format: "docx",
  kinds: ["comment"],
  list(context, options) {
    const index = createDocxObjectIndex(context);
    return index.descriptors.filter((descriptor) =>
      descriptor.draft.kind === "comment"
      && descriptor.draft.locator.type === "format"
      && descriptor.draft.locator.value.kind === "comment");
  },
};

export const DOCX_BEHAVIOR_OBJECT_PROVIDER: OfficeObjectProvider<DocxReferenceContext> = {
  id: "docx.behaviors",
  format: "docx",
  kinds: DOCX_STAGE4_BEHAVIOR_KINDS,
  list(context) {
    const index = createDocxObjectIndex(context);
    return index.behaviors;
  },
};

/** Registry with the built-in DOCX object families registered. */
export function createDocxObjectRegistry(): OfficeObjectRegistry<DocxReferenceContext> {
  return createOfficeObjectRegistry<DocxReferenceContext>([
    DOCX_BLOCK_OBJECT_PROVIDER,
    DOCX_COMMENT_OBJECT_PROVIDER,
    DOCX_BEHAVIOR_OBJECT_PROVIDER,
  ]);
}
