// Write preserved drawing objects back into a serialized story.
//
// The WASM serializer writes a paragraph's `sourceXml` verbatim, so an untouched
// paragraph keeps its shapes, charts, embedded objects and equations. An edit
// clears `sourceXml` and the engine rebuilds the paragraph from model fields,
// which have no slot for those objects — they would be dropped on export.
//
// This patch puts them back: it walks the paragraphs of every story part, and
// splices each preserved object into the paragraph content at the text offset it
// occupies in the model. When the part's paragraph count and the model's
// paragraph count disagree (content the model does not describe), the part is
// left untouched rather than patched at the wrong offset.

import type {
  DocModel,
  DocNode,
  ParagraphNode,
  TableNode,
} from "./types";
import type { OoxmlPackage } from "./ooxml-core";
import { findElementSpans } from "./drawing-objects";
import {
  applyDocumentXmlEdits,
  collectRunTextAtoms,
  ensurePreserveSpace,
  escapeXmlText,
  extractRunProperties,
  locateOutermostParagraphSpans,
  type DocumentXmlEdit,
} from "./document-xml-patch";

const DOCUMENT_PART = "word/document.xml";
const HEADER_FOOTER_PART = /^word\/(?:header|footer)\d+\.xml$/i;
const NOTE_PART = /^word\/(?:footnotes|endnotes)\.xml$/i;

interface PreservedObject {
  /** Offset in the paragraph's plain text where the object sits. */
  textOffset: number;
  /** Markup to write: a `w:drawing` / `w:object` / `w:pict` / `m:oMath` block. */
  xml: string;
}

interface PreservedParagraphObjects {
  /** Index among every `<w:p>` of the story part, in document order. */
  paragraphIndex: number;
  objects: PreservedObject[];
}

// ── Text box text rewrite ───────────────────────────────────────────

/**
 * Replaces the text of every `w:txbxContent` of a text box with `text`, keeping
 * each block's first paragraph and run properties so the box keeps its look.
 * Both the `mc:Choice` and the `mc:Fallback` copy are rewritten so Word and the
 * VML fallback agree.
 */
export function applyTextBoxTextToXml(xml: string, text: string | undefined): string {
  if (text === undefined) return xml;
  const spans = findElementSpans(xml, "w:txbxContent");
  if (spans.length === 0) return xml;

  let patched = xml;
  for (const span of [...spans].sort((left, right) => right.innerStart - left.innerStart)) {
    const original = patched.slice(span.innerStart, span.innerEnd);
    const paragraphProperties =
      /<w:pPr\b[\s\S]*?<\/w:pPr>|<w:pPr\b[^>]*\/>/i.exec(original)?.[0] ?? "";
    const runProperties =
      /<w:rPr\b[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/i.exec(original)?.[0] ?? "";

    const paragraphs = text.split(/\r?\n/).map((line) => {
      if (!line) {
        return `<w:p>${paragraphProperties}</w:p>`;
      }
      return (
        `<w:p>${paragraphProperties}<w:r>${runProperties}` +
        `<w:t xml:space="preserve">${escapeXmlText(line)}</w:t></w:r></w:p>`
      );
    });

    patched =
      patched.slice(0, span.innerStart) +
      paragraphs.join("") +
      patched.slice(span.innerEnd);
  }
  return patched;
}

function textBoxObjectXml(
  image: Extract<DocNode, { type: "paragraph" }>["children"][number] & { type: "image" },
): string | undefined {
  const source = image.textBoxSourceXml ?? image.sourceXml;
  if (!source) return undefined;
  return applyTextBoxTextToXml(source, image.textBoxText);
}

// ── Collection ──────────────────────────────────────────────────────

function collectObjectsFromParagraph(
  paragraph: ParagraphNode,
): PreservedObject[] | undefined {
  // A paragraph that still carries its markup is written verbatim; its objects
  // are already in there.
  if (paragraph.sourceXml) return undefined;

  const objects: PreservedObject[] = [];
  let textOffset = 0;
  for (const child of paragraph.children) {
    if (child.type === "text") {
      textOffset += child.text.length;
      continue;
    }
    if (child.type === "drawing") {
      if (child.sourceXml) objects.push({ textOffset, xml: child.sourceXml });
      continue;
    }
    if (child.type === "math") {
      if (child.sourceXml) objects.push({ textOffset, xml: child.sourceXml });
      continue;
    }
    if (child.type === "image") {
      if (!child.syntheticTextBox) continue;
      const xml = textBoxObjectXml(child);
      if (xml) objects.push({ textOffset, xml });
    }
  }

  return objects.length > 0 ? objects : undefined;
}

function collectNodes(
  nodes: readonly DocNode[],
  targets: PreservedParagraphObjects[],
  counter: { value: number },
): void {
  for (const node of nodes) {
    if (node.type === "table") {
      collectTable(node, targets, counter);
      continue;
    }
    const paragraphIndex = counter.value;
    counter.value += 1;
    const objects = collectObjectsFromParagraph(node);
    if (objects) targets.push({ paragraphIndex, objects });
  }
}

function collectTable(
  table: TableNode,
  targets: PreservedParagraphObjects[],
  counter: { value: number },
): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      collectNodes(cell.nodes, targets, counter);
    }
  }
}

/** Paragraphs of one story that need preserved objects written back. */
export function collectPreservedDrawingObjects(
  nodes: readonly DocNode[],
): PreservedParagraphObjects[] {
  const targets: PreservedParagraphObjects[] = [];
  collectNodes(nodes, targets, { value: 0 });
  return targets;
}

/** Total paragraph count of a story, used to align the patch with the part. */
export function storyParagraphCount(nodes: readonly DocNode[]): number {
  let count = 0;
  const visitNodes = (entries: readonly DocNode[]): void => {
    for (const node of entries) {
      if (node.type === "paragraph") {
        count += 1;
        continue;
      }
      for (const row of node.rows) {
        for (const cell of row.cells) visitNodes(cell.nodes);
      }
    }
  };
  visitNodes(nodes);
  return count;
}

export function modelNeedsDrawingObjectSerialization(model: DocModel): boolean {
  if (collectPreservedDrawingObjects(model.nodes).length > 0) return true;
  const stories: DocNode[][] = [];
  for (const section of model.metadata.headerSections ?? []) stories.push(section.nodes);
  for (const section of model.metadata.footerSections ?? []) stories.push(section.nodes);
  for (const section of model.metadata.sections ?? []) {
    for (const header of section.headerSections) stories.push(header.nodes);
    for (const footer of section.footerSections) stories.push(footer.nodes);
  }
  for (const note of model.metadata.footnotes ?? []) {
    if (note.nodes) stories.push(note.nodes);
  }
  for (const note of model.metadata.endnotes ?? []) {
    if (note.nodes) stories.push(note.nodes);
  }
  return stories.some((story) => collectPreservedDrawingObjects(story).length > 0);
}

// ── Insertion ───────────────────────────────────────────────────────

function objectRunXml(xml: string): string {
  const trimmed = xml.trim();
  // Equations are paragraph content (`m:oMath` is not a run). Everything else
  // is a drawing / embedded object that Word keeps inside a `w:r`.
  if (/^<m:oMath/i.test(trimmed)) return trimmed;
  return `<w:r>${trimmed}</w:r>`;
}

function paragraphPropertiesLength(content: string): number {
  const match = /^\s*<w:pPr\b[\s\S]*?<\/w:pPr>|^\s*<w:pPr\b[^>]*\/>/.exec(content);
  return match ? match[0].length : 0;
}

function insertAtTextOffset(
  content: string,
  textOffset: number,
  objectXml: string,
): string {
  const atoms = collectRunTextAtoms(content);
  const runXml = objectRunXml(objectXml);

  if (atoms.length === 0) {
    const at = paragraphPropertiesLength(content);
    return content.slice(0, at) + runXml + content.slice(at);
  }

  const inside = atoms.find(
    (atom) => textOffset > atom.textOffset && textOffset < atom.textOffset + atom.text.length,
  );
  if (inside) {
    const localOffset = textOffset - inside.textOffset;
    const originalRun = content.slice(inside.runStart, inside.runEnd);
    const runOpenTag = originalRun.slice(0, originalRun.indexOf(">") + 1);
    const runProperties = extractRunProperties(originalRun);
    const textOpenTag = ensurePreserveSpace(inside.openTag);
    const runForText = (text: string): string =>
      `${runOpenTag}${runProperties}${textOpenTag}${escapeXmlText(text)}${inside.closeTag}</w:r>`;
    return (
      content.slice(0, inside.runStart) +
      runForText(inside.text.slice(0, localOffset)) +
      runXml +
      runForText(inside.text.slice(localOffset)) +
      content.slice(inside.runEnd)
    );
  }

  if (textOffset <= (atoms[0]?.textOffset ?? 0)) {
    const at = Math.max(paragraphPropertiesLength(content), atoms[0]?.runStart ?? 0);
    return content.slice(0, at) + runXml + content.slice(at);
  }

  const before = [...atoms]
    .reverse()
    .find((atom) => atom.textOffset + atom.text.length <= textOffset);
  if (before) {
    return content.slice(0, before.runEnd) + runXml + content.slice(before.runEnd);
  }
  return content + runXml;
}

/** Writes the preserved objects of one paragraph content into it. */
export function applyObjectsToParagraphContent(
  content: string,
  objects: readonly PreservedObject[],
): string {
  let patched = content;
  for (const object of objects) {
    patched = insertAtTextOffset(patched, object.textOffset, object.xml);
  }
  return patched;
}

/**
 * Patches one story XML with the objects its model paragraphs preserved.
 *
 * The patch is all-or-nothing: a part whose paragraph count differs from the
 * model's (text-box content, tracked-change wrappers the model flattens, ...)
 * is returned unchanged instead of being patched at a shifted offset.
 */
export function applyPreservedDrawingObjectsToStoryXml(
  xml: string,
  nodes: readonly DocNode[],
): string {
  const targets = collectPreservedDrawingObjects(nodes);
  if (targets.length === 0) return xml;

  const spans = locateOutermostParagraphSpans(xml);
  if (spans.length !== storyParagraphCount(nodes)) return xml;

  const targetByIndex = new Map(targets.map((target) => [target.paragraphIndex, target]));
  const edits: DocumentXmlEdit[] = [];
  spans.forEach((span, paragraphIndex) => {
    const target = targetByIndex.get(paragraphIndex);
    if (!target) return;
    const content = xml.slice(span.contentStart, span.contentEnd);
    const patched = applyObjectsToParagraphContent(content, target.objects);
    if (patched === content) return;
    edits.push({
      start: span.contentStart,
      end: span.contentEnd,
      replacement: patched,
    });
  });
  return applyDocumentXmlEdits(xml, edits);
}

/**
 * Body-only patch, for `modelToDocumentXml` (which returns `word/document.xml`
 * without a package to walk).
 */
export function applyPreservedDrawingObjectsToDocumentXml(
  documentXml: string,
  model: DocModel,
): string {
  return applyPreservedDrawingObjectsToStoryXml(documentXml, model.nodes);
}

function storyNodesForPart(
  model: DocModel,
  partName: string,
): readonly DocNode[] | undefined {
  const normalized = partName.toLowerCase();
  const fromModel = (sections: { partName: string; nodes: DocNode[] }[] | undefined) =>
    sections?.find((section) => section.partName.toLowerCase() === normalized)?.nodes;

  if (normalized === DOCUMENT_PART) return model.nodes;
  if (HEADER_FOOTER_PART.test(normalized)) {
    return (
      fromModel(model.metadata.headerSections) ??
      fromModel(model.metadata.footerSections) ??
      model.metadata.sections
        ?.flatMap((section) => [...section.headerSections, ...section.footerSections])
        .find((section) => section.partName.toLowerCase() === normalized)?.nodes
    );
  }
  if (NOTE_PART.test(normalized)) {
    const note = normalized.includes("endnotes")
      ? model.metadata.endnotes
      : model.metadata.footnotes;
    const nodes: DocNode[] = [];
    for (const definition of note ?? []) {
      if (definition.nodes) nodes.push(...definition.nodes);
    }
    return nodes.length > 0 ? nodes : undefined;
  }
  return undefined;
}

/** Patches one package part; parts without a story are returned unchanged. */
export function applyPreservedDrawingObjectsToPartContent(
  partName: string,
  content: string,
  model: DocModel,
): string {
  const nodes = storyNodesForPart(model, partName);
  if (!nodes) return content;
  return applyPreservedDrawingObjectsToStoryXml(content, nodes);
}

/** Patches every story part of a parsed package (`serializeDocModel`).
 *
 * The body part is patched by `patchPackageDocumentXml` before this runs, so it
 * is skipped here to keep the preserved objects from being written twice.
 */
export function applyPreservedDrawingObjectsToPackage(
  pkg: OoxmlPackage,
  model: DocModel,
): OoxmlPackage {
  if (!modelNeedsDrawingObjectSerialization(model)) return pkg;
  let parts: Map<string, { name: string; content: string }> | undefined;
  for (const [partName, part] of pkg.parts) {
    if (partName.toLowerCase() === DOCUMENT_PART) continue;
    const content = applyPreservedDrawingObjectsToPartContent(partName, part.content, model);
    if (content === part.content) continue;
    parts ??= new Map(pkg.parts);
    parts.set(partName, { name: part.name, content });
  }
  return parts ? { parts, binaryAssets: pkg.binaryAssets } : pkg;
}
