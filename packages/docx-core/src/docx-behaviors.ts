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

interface DocxStoryParagraph {
  paragraph: ParagraphNode;
  path: readonly DocxPathSegment[];
  part: DocxPartLocator;
}

function storyParagraphsInNodes(
  nodes: readonly DocNode[],
  prefix: readonly DocxPathSegment[],
  part: DocxPartLocator,
  visit: (entry: DocxStoryParagraph) => void,
): void {
  nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      visit({ paragraph: node, path: [...prefix, { kind: "paragraph", index: nodeIndex }], part });
      return;
    }
    const table: TableNode = node;
    table.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        storyParagraphsInNodes(
          cell.nodes,
          [
            ...prefix,
            { kind: "table", index: nodeIndex },
            { kind: "row", index: rowIndex },
            { kind: "cell", index: cellIndex },
          ],
          part,
          visit,
        );
      });
    });
  });
}

function collectStoryParts(model: DocModel): readonly { part: DocxPartLocator; nodes: readonly DocNode[] }[] {
  const parts: { part: DocxPartLocator; nodes: readonly DocNode[] }[] = [
    { part: { kind: "body" }, nodes: model.nodes },
  ];
  const seen = new Set<string>(["body"]);
  const addStory = (kind: "header" | "footer", partName: string, nodes: readonly DocNode[]): void => {
    const key = `${kind}:${partName}`;
    if (seen.has(key) || nodes.length === 0) return;
    seen.add(key);
    parts.push({ part: { kind, partName }, nodes });
  };
  for (const section of model.metadata.headerSections ?? []) {
    addStory("header", section.partName, section.nodes);
  }
  for (const section of model.metadata.footerSections ?? []) {
    addStory("footer", section.partName, section.nodes);
  }
  for (const section of model.metadata.sections ?? []) {
    for (const header of section.headerSections) addStory("header", header.partName, header.nodes);
    for (const footer of section.footerSections) addStory("footer", footer.partName, footer.nodes);
  }
  for (const note of model.metadata.footnotes ?? []) {
    if (note.nodes?.length) parts.push({ part: { kind: "footnote", noteId: String(note.id) }, nodes: note.nodes });
  }
  for (const note of model.metadata.endnotes ?? []) {
    if (note.nodes?.length) parts.push({ part: { kind: "endnote", noteId: String(note.id) }, nodes: note.nodes });
  }
  return parts;
}

/**
 * Flatten every paragraph of every story part (body, headers, footers,
 * footnotes, endnotes) with the structural path needed by a behavior owner.
 * Fields and revisions inside header/footer and note stories are behavior
 * targets exactly like body paragraphs; omitting them made stage 4 silently
 * body-only.
 */
function eachStoryParagraph(
  model: DocModel,
  visit: (entry: DocxStoryParagraph) => void,
): void {
  for (const story of collectStoryParts(model)) {
    storyParagraphsInNodes(story.nodes, [], story.part, visit);
  }
}

function storyKey(part: DocxPartLocator): string {
  if (part.kind === "body") return "body";
  if (part.kind === "header" || part.kind === "footer") return `${part.kind}:${part.partName}`;
  if (part.kind === "footnote" || part.kind === "endnote") return `${part.kind}:${part.noteId}`;
  return "body";
}

function storyPath(part: DocxPartLocator): readonly { kind: OfficeObjectKind; label: string }[] {
  if (part.kind === "body") return [];
  if (part.kind === "header") return [{ kind: "section", label: `Header ${part.partName}` }];
  if (part.kind === "footer") return [{ kind: "section", label: `Footer ${part.partName}` }];
  if (part.kind === "footnote" || part.kind === "endnote") {
    return [{
      kind: "note",
      label: `${part.kind === "footnote" ? "Footnote" : "Endnote"} ${part.noteId}`,
    }];
  }
  return [];
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

function ownerFor(path: readonly DocxPathSegment[], part: DocxPartLocator): DocxBehaviorOwner {
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
  eachStoryParagraph(model, ({ paragraph, path, part }) => {
    const xml = paragraph.sourceXml;
    if (!xml) return;
    const key = pathKey(path);
    const story = storyKey(part);
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
        instanceId: `docx.${behavior}:${story}:${key}:${ordinal}:${messageHash(instruction)}`,
        owner: ownerFor(path, part),
        label,
        path: [
          { kind: "document", label: "Document" },
          ...storyPath(part),
          { kind: crossReference ? "cross-reference" : "field", label },
        ],
        content: { text: instruction },
        facets: {
          instruction,
          fieldType: token,
          ...(target ? { target } : {}),
          path: key,
          part: story,
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
  eachStoryParagraph(model, ({ paragraph, path, part }) => {
    const xml = paragraph.sourceXml;
    if (!xml) return;
    const key = pathKey(path);
    const story = storyKey(part);
    for (const [type, pattern] of [["insertion", TRACKED_INSERTION], ["deletion", TRACKED_DELETION]] as const) {
      pattern.lastIndex = 0;
      let ordinal = 0;
      for (let match = pattern.exec(xml); match; match = pattern.exec(xml)) {
        const author = AUTHOR_ATTRIBUTE.exec(match[1] ?? "")?.[1] ?? "";
        records.push({
          behavior: "tracked-change-state",
          instanceId: `docx.tracked-changes:${story}:${key}:${type}:${ordinal}:${messageHash(author)}`,
          owner: ownerFor(path, part),
          label: `Revision by ${author || "unknown author"}`,
          path: [
            { kind: "document", label: "Document" },
            ...storyPath(part),
            { kind: "tracked-change", label: `Revision by ${author || "unknown author"}` },
          ],
          content: { text: paragraphText(paragraph) },
          facets: { changeType: type, author, path: key, part: story, ordinal },
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
