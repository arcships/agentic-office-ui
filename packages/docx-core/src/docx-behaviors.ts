// Stage 4 invisible behavior descriptors for DOCX.
//
// Everything here is derived from the document model only: no DOM, no layout
// and no screen geometry. Behaviors are *references* with an owner scope, not
// canvas outlines, so they take part in the same descriptor/reference contract
// as visible objects while staying selectable from a structure tree or panel.

import type {
  DocxBehaviorLocator,
  DocxBehaviorOwner,
  DocxPartLocator,
  DocxPathSegment,
  JsonValue,
  OfficeObjectKind,
} from "@arcships/office-interaction";
import type { DocModel, DocNode, ParagraphNode, TableNode } from "./engine/types";
import { createDocxBookmarkIndex } from "./editor/helpers/bookmarks";
import { paragraphText } from "./editor/helpers/paragraph-inspect";

/** One behavior instance the document currently contains. */
export interface DocxBehaviorRecord {
  behavior: DocxBehaviorLocator["behavior"];
  /** Deterministic identity inside the owner scope; never random. */
  instanceId: string;
  owner: DocxBehaviorOwner;
  label: string;
  path: readonly { kind: OfficeObjectKind; label: string }[];
  content?: { text?: string };
  facets: Readonly<Record<string, JsonValue>>;
}

export type DocxBehaviorKindMap = {
  style: "style";
  field: "field";
  bookmark: "bookmark";
  section: "section";
  "tracked-change-state": "tracked-change";
  "comment-state": "comment";
  "cross-reference": "cross-reference";
};

export const DOCX_BEHAVIOR_REFERENCE_KIND: DocxBehaviorKindMap = {
  style: "style",
  field: "field",
  bookmark: "bookmark",
  section: "section",
  "tracked-change-state": "tracked-change",
  "comment-state": "comment",
  "cross-reference": "cross-reference",
};

const FIELD_INSTRUCTION = /<w:(?:instrText)\b[^>]*>([\s\S]*?)<\/w:instrText>|<w:fldSimple\b[^>]*\bw:instr="([^"]*)"/giu;
const TRACKED_INSERTION = /<w:ins\b([^>]*)>/giu;
const TRACKED_DELETION = /<w:del\b([^>]*)>/giu;
const AUTHOR_ATTRIBUTE = /\bw:author="([^"]*)"/iu;

function pathKey(path: readonly DocxPathSegment[]): string {
  return path.map((segment) => `${segment.kind}:${segment.index}`).join("/");
}

function nodePath(index: number): readonly DocxPathSegment[] {
  return [{ kind: "paragraph", index }];
}

function cellPath(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number,
): readonly DocxPathSegment[] {
  return [
    { kind: "table", index: tableIndex },
    { kind: "row", index: rowIndex },
    { kind: "cell", index: cellIndex },
    { kind: "paragraph", index: paragraphIndex },
  ];
}

function paragraphTexts(nodes: readonly DocNode[]): string[] {
  return nodes.map((node) => node.type === "paragraph" ? paragraphText(node) : "");
}

/** Flatten body paragraphs with the structural path needed by a behavior owner. */
function eachBodyParagraph(
  model: DocModel,
  visit: (paragraph: ParagraphNode, path: readonly DocxPathSegment[], tableIndex?: number) => void,
): void {
  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      visit(node, nodePath(nodeIndex));
      return;
    }
    const table = node as TableNode;
    table.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        cell.nodes.forEach((entry, paragraphIndex) => {
          if (entry.type !== "paragraph") return;
          visit(entry, cellPath(nodeIndex, rowIndex, cellIndex, paragraphIndex), nodeIndex);
        });
      });
    });
  });
}

function firstToken(instruction: string): string {
  return instruction.trim().split(/\s+/u)[0] ?? "";
}

function fieldTarget(instruction: string): string | undefined {
  const [, ...rest] = instruction.trim().split(/\s+/u);
  const quote = rest.join(" ");
  const quoted = /"([^"]+)"/u.exec(quote);
  if (quoted) return quoted[1];
  return rest[0];
}

function messageHash(value: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    hash ^= code & 0xff;
    hash = Math.imul(hash, 0x01000193);
    hash ^= code >>> 8;
    hash = Math.imul(hash, 0x01000193);
  }
  return `fnv1a32:${value.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

function ownerFor(path: readonly DocxPathSegment[]): DocxBehaviorOwner {
  const part: DocxPartLocator = { kind: "body" };
  return { scope: "part", part, path };
}

function styleRecords(model: DocModel): DocxBehaviorRecord[] {
  return (model.metadata.paragraphStyles ?? []).map((style, index) => {
    const name = style.name?.trim() || style.id;
    return {
      behavior: "style",
      instanceId: `docx.style:${style.id}`,
      owner: { scope: "document" },
      label: `Style: ${name}`,
      path: [
        { kind: "document", label: "Document" },
        { kind: "style", label: `Style: ${name}` },
      ],
      facets: {
        styleId: style.id,
        styleName: name,
        order: index,
        ...(style.basedOnId ? { basedOnId: style.basedOnId } : {}),
        ...(style.headingLevel ? { headingLevel: style.headingLevel } : {}),
      },
    };
  });
}

function sectionRecords(model: DocModel): DocxBehaviorRecord[] {
  const sections = model.metadata.sections ?? [];
  if (sections.length === 0) {
    if (model.nodes.length === 0) return [];
    return [{
      behavior: "section",
      instanceId: "docx.section:0",
      owner: { scope: "document" },
      label: "Section 1",
      path: [
        { kind: "document", label: "Document" },
        { kind: "section", label: "Section 1" },
      ],
      facets: { index: 0, startNodeIndex: 0, implicit: true },
    }];
  }
  return sections.map((section, index) => ({
    behavior: "section",
    instanceId: `docx.section:${index}`,
    owner: { scope: "document" },
    label: `Section ${index + 1}`,
    path: [
      { kind: "document", label: "Document" },
      { kind: "section", label: `Section ${index + 1}` },
    ],
    facets: {
      index,
      startNodeIndex: section.startNodeIndex,
      headerParts: section.headerSections.map((header) => header.partName),
      footerParts: section.footerSections.map((footer) => footer.partName),
    },
  }));
}

function bookmarkRecords(model: DocModel): DocxBehaviorRecord[] {
  return createDocxBookmarkIndex(model).targets.map((target) => ({
    behavior: "bookmark",
    instanceId: `docx.bookmark:${target.name}`,
    owner: { scope: "document" },
    label: `Bookmark: ${target.name}`,
    path: [
      { kind: "document", label: "Document" },
      { kind: "bookmark", label: `Bookmark: ${target.name}` },
    ],
    facets: {
      name: target.name,
      nodeIndex: target.nodeIndex,
      ...(target.tableCell ? { tableCell: { ...target.tableCell } } : {}),
    },
  }));
}

function fieldRecords(model: DocModel): DocxBehaviorRecord[] {
  const records: DocxBehaviorRecord[] = [];
  eachBodyParagraph(model, (paragraph, path) => {
    const xml = paragraph.sourceXml;
    if (!xml) return;
    const key = pathKey(path);
    let ordinal = 0;
    FIELD_INSTRUCTION.lastIndex = 0;
    for (let match = FIELD_INSTRUCTION.exec(xml); match; match = FIELD_INSTRUCTION.exec(xml)) {
      const instruction = (match[1] ?? match[2] ?? "").trim();
      if (!instruction) continue;
      const token = firstToken(instruction).toUpperCase();
      const target = fieldTarget(instruction);
      const crossReference = token === "REF" || token === "PAGEREF" || token === "NOTEREF";
      const behavior = crossReference ? "cross-reference" : "field";
      const label = crossReference
        ? `Cross-reference: ${target ?? token}`
        : `Field: ${target ?? token}`;
      records.push({
        behavior,
        instanceId: `docx.${behavior}:${key}:${ordinal}:${messageHash(instruction)}`,
        owner: ownerFor(path),
        label,
        path: [
          { kind: "document", label: "Document" },
          { kind: crossReference ? "cross-reference" : "field", label },
        ],
        content: { text: instruction },
        facets: {
          instruction,
          fieldType: token,
          ...(target ? { target } : {}),
          path: key,
          ordinal,
        },
      });
      ordinal += 1;
    }
  });
  return records;
}

function trackedChangeRecords(model: DocModel): DocxBehaviorRecord[] {
  const records: DocxBehaviorRecord[] = [];
  eachBodyParagraph(model, (paragraph, path) => {
    const xml = paragraph.sourceXml;
    if (!xml) return;
    const key = pathKey(path);
    for (const [type, pattern] of [["insertion", TRACKED_INSERTION], ["deletion", TRACKED_DELETION]] as const) {
      pattern.lastIndex = 0;
      let ordinal = 0;
      for (let match = pattern.exec(xml); match; match = pattern.exec(xml)) {
        const author = AUTHOR_ATTRIBUTE.exec(match[1] ?? "")?.[1] ?? "";
        records.push({
          behavior: "tracked-change-state",
          instanceId: `docx.tracked-changes:${key}:${type}:${ordinal}:${messageHash(author)}`,
          owner: ownerFor(path),
          label: `Revision by ${author || "unknown author"}`,
          path: [
            { kind: "document", label: "Document" },
            { kind: "tracked-change", label: `Revision by ${author || "unknown author"}` },
          ],
          content: { text: paragraphText(paragraph) },
          facets: { changeType: type, author, path: key, ordinal },
        });
        ordinal += 1;
      }
    }
  });
  return records;
}

function commentStateRecords(model: DocModel): DocxBehaviorRecord[] {
  return (model.metadata.comments ?? []).map((comment) => {
    const state = comment.resolved ? "resolved" : "open";
    return {
      behavior: "comment-state",
      instanceId: `docx.comment-state:${comment.id}`,
      owner: { scope: "document" },
      label: `Comment state: ${state}`,
      path: [
        { kind: "document", label: "Document" },
        { kind: "comment", label: `Comment state: ${state}` },
      ],
      content: { text: comment.text },
      facets: {
        commentId: String(comment.id),
        state,
        resolved: comment.resolved === true,
        ...(comment.parentId === undefined ? {} : { parentId: String(comment.parentId) }),
      },
    };
  });
}

/**
 * Enumerate every invisible DOCX behavior of the current model.
 *
 * Order is stable (styles, sections, bookmarks, fields, revisions, comment
 * state) so an object index, a structure tree and a test fixture agree.
 */
export function enumerateDocxBehaviors(model: DocModel): readonly DocxBehaviorRecord[] {
  return Object.freeze([
    ...styleRecords(model),
    ...sectionRecords(model),
    ...bookmarkRecords(model),
    ...fieldRecords(model),
    ...trackedChangeRecords(model),
    ...commentStateRecords(model),
  ]);
}

/** Resolve one behavior instance by its deterministic id. */
export function findDocxBehavior(
  model: DocModel,
  instanceId: string,
): DocxBehaviorRecord | undefined {
  return enumerateDocxBehaviors(model).find((record) => record.instanceId === instanceId);
}
