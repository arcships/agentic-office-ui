// Comment part serialization: `word/comments.xml` + `word/commentsExtended.xml`.
//
// The WASM serializer writes `word/document.xml` from the model and carries
// every other part over from the base package, which is exactly right for a
// document that only changed its text. Authoring comments does not go through
// WASM, so the model's comment definitions are written here: the comment bodies,
// the reply/`done` threading in commentsExtended, the package relationship and
// content-type entries, and the removal of both parts when the last comment is
// deleted.

import { strFromU8, strToU8 } from "fflate";
import type { DocModel, DocumentCommentDefinition } from "./types";
import type { OoxmlPackage } from "./ooxml-core";
import { escapeXmlText } from "./document-xml-patch";

export const DOCX_COMMENTS_PART = "word/comments.xml";
export const DOCX_COMMENTS_EXTENDED_PART = "word/commentsExtended.xml";

const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_EXTENDED_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
const COMMENTS_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_EXTENDED_REL_TYPE =
  "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const WORD_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_2010_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const WORD_2012_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";

/** Stable `w14:paraId` for a comment, derived from its `w:id`. */
export function commentParagraphId(commentId: number): string {
  const value = (0x5eed0000 + (Number.isFinite(commentId) ? commentId : 0)) & 0x7fffffff;
  return value.toString(16).toUpperCase().padStart(8, "0");
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/g, "&quot;");
}

/** Comment bodies keep their line breaks as separate comment paragraphs. */
function commentParagraphsXml(comment: DocumentCommentDefinition): string {
  const lines = comment.text.split(/\r?\n/);
  const paragraphId = commentParagraphId(comment.id);
  return lines
    .map((line, index) => {
      const attributes =
        index === lines.length - 1 ? ` w14:paraId="${paragraphId}"` : "";
      return (
        `<w:p${attributes}><w:r>` +
        `<w:t xml:space="preserve">${escapeXmlText(line)}</w:t>` +
        `</w:r></w:p>`
      );
    })
    .join("");
}

export function serializeCommentsXml(
  definitions: readonly DocumentCommentDefinition[]
): string {
  const body = definitions
    .map((comment) => {
      const attributes = [`w:id="${comment.id}"`];
      if (comment.author) {
        attributes.push(`w:author="${escapeXmlAttribute(comment.author)}"`);
      }
      if (comment.initials) {
        attributes.push(`w:initials="${escapeXmlAttribute(comment.initials)}"`);
      }
      if (comment.date) {
        attributes.push(`w:date="${escapeXmlAttribute(comment.date)}"`);
      }
      return `<w:comment ${attributes.join(" ")}>${commentParagraphsXml(
        comment
      )}</w:comment>`;
    })
    .join("");

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:comments xmlns:w="${WORD_NAMESPACE}" xmlns:w14="${WORD_2010_NAMESPACE}">` +
    `${body}</w:comments>`
  );
}

function commentNeedsExtendedEntry(comment: DocumentCommentDefinition): boolean {
  return comment.parentId !== undefined || comment.resolved === true;
}

export function commentsNeedExtendedPart(
  definitions: readonly DocumentCommentDefinition[]
): boolean {
  return definitions.some(commentNeedsExtendedEntry);
}

export function serializeCommentsExtendedXml(
  definitions: readonly DocumentCommentDefinition[]
): string {
  const entries = definitions
    .map((comment) => {
      const attributes = [`w15:paraId="${commentParagraphId(comment.id)}"`];
      if (comment.parentId !== undefined) {
        attributes.push(
          `w15:paraIdParent="${commentParagraphId(comment.parentId)}"`
        );
      }
      if (comment.resolved === true) {
        attributes.push(`w15:done="1"`);
      }
      return `<w15:commentEx ${attributes.join(" ")}/>`;
    })
    .join("");

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w15:commentsEx xmlns:w15="${WORD_2012_NAMESPACE}">${entries}</w15:commentsEx>`
  );
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_match, code: string) =>
      String.fromCodePoint(Number.parseInt(code, 16))
    )
    .replace(/&#(\d+);/g, (_match, code: string) =>
      String.fromCodePoint(Number.parseInt(code, 10))
    )
    .replace(/&amp;/g, "&");
}

function attributeValue(attributes: string, name: string): string | undefined {
  const match = new RegExp(`\\b${name}="([^"]*)"`).exec(attributes);
  return match ? decodeXmlText(match[1] ?? "") : undefined;
}

interface ParsedCommentParagraphIds {
  /** `w:id` → `w14:paraId` of the comment's last paragraph. */
  paraIdByCommentId: Map<number, string>;
}

function parseCommentParagraphIds(commentsXml: string): ParsedCommentParagraphIds {
  const paraIdByCommentId = new Map<number, string>();
  for (const match of commentsXml.matchAll(
    /<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g
  )) {
    const id = Number.parseInt(attributeValue(match[1] ?? "", "w:id") ?? "", 10);
    const paragraphs = [...(match[2] ?? "").matchAll(/<w:p\b([^>]*)>/g)];
    const last = paragraphs[paragraphs.length - 1];
    const paraId = last ? attributeValue(last[1] ?? "", "w14:paraId") : undefined;
    if (Number.isFinite(id) && paraId) {
      paraIdByCommentId.set(id, paraId);
    }
  }
  return { paraIdByCommentId };
}

function parseCommentsXml(commentsXml: string): DocumentCommentDefinition[] | undefined {
  const definitions: DocumentCommentDefinition[] = [];
  for (const match of commentsXml.matchAll(
    /<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g
  )) {
    const attributes = match[1] ?? "";
    const id = Number.parseInt(attributeValue(attributes, "w:id") ?? "", 10);
    if (!Number.isFinite(id)) {
      return undefined;
    }
    const text = [...(match[2] ?? "").matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
      .map((piece) => decodeXmlText(piece[1] ?? ""))
      .join("\n");
    definitions.push({
      id,
      author: attributeValue(attributes, "w:author"),
      initials: attributeValue(attributes, "w:initials"),
      date: attributeValue(attributes, "w:date"),
      text,
      parentId: undefined,
      resolved: false,
    });
  }
  return definitions;
}

function applyExtendedThreading(
  definitions: DocumentCommentDefinition[],
  commentsXml: string,
  extendedXml: string | undefined
): DocumentCommentDefinition[] {
  if (!extendedXml) {
    return definitions;
  }

  const { paraIdByCommentId } = parseCommentParagraphIds(commentsXml);
  const commentIdByParaId = new Map<string, number>();
  paraIdByCommentId.forEach((paraId, commentId) => {
    commentIdByParaId.set(paraId.toUpperCase(), commentId);
  });

  const byId = new Map(definitions.map((definition) => [definition.id, definition]));
  for (const match of extendedXml.matchAll(/<w15:commentEx\b([^>]*)\/?>/g)) {
    const attributes = match[1] ?? "";
    const paraId = attributeValue(attributes, "w15:paraId");
    const commentId = paraId
      ? commentIdByParaId.get(paraId.toUpperCase())
      : undefined;
    const definition = commentId === undefined ? undefined : byId.get(commentId);
    if (!definition) {
      continue;
    }
    const parentParaId = attributeValue(attributes, "w15:paraIdParent");
    const parentId = parentParaId
      ? commentIdByParaId.get(parentParaId.toUpperCase())
      : undefined;
    definition.parentId = parentId;
    definition.resolved = attributeValue(attributes, "w15:done") === "1";
  }
  return definitions;
}

/**
 * Signature of the fields the public comment commands can change: which
 * comments exist, how they thread and whether a thread is resolved.
 *
 * Author, date and body text are deliberately excluded: no command edits them,
 * and comparing them would rewrite (and so re-render) the comment parts of an
 * imported document whose bodies carry rich runs the parser drops.
 */
function commentSignature(
  definitions: readonly DocumentCommentDefinition[]
): string {
  return [...definitions]
    .sort((left, right) => left.id - right.id)
    .map((comment) =>
      [
        comment.id,
        comment.parentId ?? "",
        comment.resolved === true ? "1" : "",
      ].join("\u001f")
    )
    .join("\u001e");
}

/** True when `commentsXml` (+ extended) already stores exactly these definitions. */
export function commentsXmlMatchesModel(
  commentsXml: string | undefined,
  extendedXml: string | undefined,
  definitions: readonly DocumentCommentDefinition[]
): boolean {
  if (!commentsXml) {
    return false;
  }
  const parsed = parseCommentsXml(commentsXml);
  if (!parsed) {
    return false;
  }
  const threaded = applyExtendedThreading(parsed, commentsXml, extendedXml);
  return commentSignature(threaded) === commentSignature(definitions);
}

/** Ids referenced by `w:commentRangeStart`/`w:commentRangeEnd`/`w:commentReference`. */
function collectReferencedCommentIds(xml: string): Set<number> {
  const ids = new Set<number>();
  for (const match of xml.matchAll(
    /<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*w:id="(-?\d+)"[^>]*\/?>/g
  )) {
    const id = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(id)) {
      ids.add(id);
    }
  }
  return ids;
}

/**
 * Removes anchors of comments that no longer exist.
 *
 * A deleted comment leaves its `w:commentRangeStart` / `w:commentRangeEnd` and
 * its `w:commentReference` behind in `sourceXml`, which the serializer writes
 * verbatim. Word ignores dangling anchors, but they point at an id another
 * comment can later reuse, so they must go when the definition does.
 */
export function stripOrphanCommentMarkers(
  xml: string,
  knownIds: ReadonlySet<number>
): string {
  const referenced = collectReferencedCommentIds(xml);
  let orphaned = false;
  referenced.forEach((id) => {
    if (!knownIds.has(id)) {
      orphaned = true;
    }
  });
  if (!orphaned) {
    return xml;
  }

  return xml
    .replace(
      /<w:commentReference\b[^>]*w:id="(-?\d+)"[^>]*\/?>/g,
      (match, rawId: string) =>
        knownIds.has(Number.parseInt(rawId, 10)) ? match : ""
    )
    .replace(
      /<w:comment(?:RangeStart|RangeEnd)\b[^>]*w:id="(-?\d+)"[^>]*\/?>/g,
      (match, rawId: string) =>
        knownIds.has(Number.parseInt(rawId, 10)) ? match : ""
    )
    // Drop the runs the removed references left behind (runs that only carry
    // run properties have no visible content).
    .replace(
      /<w:r\b[^>]*>\s*(?:<w:rPr\b[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>)?\s*<\/w:r>/g,
      ""
    );
}

interface MutablePartSet {
  get(partName: string): string | undefined;
  set(partName: string, content: string): void;
  delete(partName: string): void;
}

function hasRelationship(relsXml: string, relationType: string): boolean {
  return new RegExp(`<Relationship\\b[^>]*Type="${relationType}"`).test(relsXml);
}

function removeRelationship(relsXml: string, relationType: string): string {
  return relsXml.replace(
    new RegExp(`<Relationship\\b[^>]*Type="${relationType}"[^>]*/>`, "g"),
    ""
  );
}

function nextRelationshipId(relsXml: string): string {
  let max = 0;
  for (const match of relsXml.matchAll(/\bId="rId(\d+)"/g)) {
    const value = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(value) && value > max) {
      max = value;
    }
  }
  return `rId${max + 1}`;
}

function addRelationship(
  relsXml: string,
  relationType: string,
  target: string
): string {
  if (hasRelationship(relsXml, relationType)) {
    return relsXml;
  }
  const entry =
    `<Relationship Id="${nextRelationshipId(relsXml)}" ` +
    `Type="${relationType}" Target="${target}"/>`;
  return relsXml.includes("</Relationships>")
    ? relsXml.replace("</Relationships>", `${entry}</Relationships>`)
    : relsXml;
}

function hasContentTypeOverride(contentTypesXml: string, partName: string): boolean {
  return new RegExp(
    `<Override\\b[^>]*PartName="${partName.replace(/[/.]/g, "\\$&")}"`
  ).test(contentTypesXml);
}

function addContentTypeOverride(
  contentTypesXml: string,
  partName: string,
  contentType: string
): string {
  if (hasContentTypeOverride(contentTypesXml, partName)) {
    return contentTypesXml;
  }
  const entry = `<Override PartName="${partName}" ContentType="${contentType}"/>`;
  return contentTypesXml.includes("</Types>")
    ? contentTypesXml.replace("</Types>", `${entry}</Types>`)
    : contentTypesXml;
}

function removeContentTypeOverride(
  contentTypesXml: string,
  partName: string
): string {
  return contentTypesXml.replace(
    new RegExp(
      `<Override\\b[^>]*PartName="${partName.replace(/[/.]/g, "\\$&")}"[^>]*/>`,
      "g"
    ),
    ""
  );
}

/**
 * Writes (or removes) the comment parts of an OOXML package and returns the
 * names it touched, so callers can skip re-zipping when nothing changed.
 */
export function applyCommentsToParts(
  parts: MutablePartSet,
  definitions: readonly DocumentCommentDefinition[]
): boolean {
  let changed = false;

  const commentsXml = parts.get(DOCX_COMMENTS_PART);
  const extendedXml = parts.get(DOCX_COMMENTS_EXTENDED_PART);
  const needsComments = definitions.length > 0;
  const needsExtended = commentsNeedExtendedPart(definitions);

  const rewriteComments =
    needsComments && !commentsXmlMatchesModel(commentsXml, extendedXml, definitions);
  if (rewriteComments) {
    parts.set(DOCX_COMMENTS_PART, serializeCommentsXml(definitions));
    changed = true;
  }
  if (!needsComments && commentsXml !== undefined) {
    parts.delete(DOCX_COMMENTS_PART);
    changed = true;
  }

  if (needsExtended && (rewriteComments || extendedXml === undefined)) {
    parts.set(DOCX_COMMENTS_EXTENDED_PART, serializeCommentsExtendedXml(definitions));
    changed = true;
  }
  if (!needsExtended && extendedXml !== undefined) {
    parts.delete(DOCX_COMMENTS_EXTENDED_PART);
    changed = true;
  }

  const relsPartName = "word/_rels/document.xml.rels";
  const relsXml = parts.get(relsPartName);
  if (relsXml !== undefined) {
    let nextRels = relsXml;
    if (needsComments) {
      nextRels = addRelationship(nextRels, COMMENTS_REL_TYPE, "comments.xml");
      if (needsExtended) {
        nextRels = addRelationship(
          nextRels,
          COMMENTS_EXTENDED_REL_TYPE,
          "commentsExtended.xml"
        );
      } else {
        nextRels = removeRelationship(nextRels, COMMENTS_EXTENDED_REL_TYPE);
      }
    } else {
      nextRels = removeRelationship(nextRels, COMMENTS_REL_TYPE);
      nextRels = removeRelationship(nextRels, COMMENTS_EXTENDED_REL_TYPE);
    }
    if (nextRels !== relsXml) {
      parts.set(relsPartName, nextRels);
      changed = true;
    }
  }

  const contentTypesXml = parts.get("[Content_Types].xml");
  if (contentTypesXml !== undefined) {
    let nextContentTypes = contentTypesXml;
    if (needsComments) {
      nextContentTypes = addContentTypeOverride(
        nextContentTypes,
        "/word/comments.xml",
        COMMENTS_CONTENT_TYPE
      );
      if (needsExtended) {
        nextContentTypes = addContentTypeOverride(
          nextContentTypes,
          "/word/commentsExtended.xml",
          COMMENTS_EXTENDED_CONTENT_TYPE
        );
      } else {
        nextContentTypes = removeContentTypeOverride(
          nextContentTypes,
          "/word/commentsExtended.xml"
        );
      }
    } else {
      nextContentTypes = removeContentTypeOverride(
        nextContentTypes,
        "/word/comments.xml"
      );
      nextContentTypes = removeContentTypeOverride(
        nextContentTypes,
        "/word/commentsExtended.xml"
      );
    }
    if (nextContentTypes !== contentTypesXml) {
      parts.set("[Content_Types].xml", nextContentTypes);
      changed = true;
    }
  }

  return changed;
}

function modelCommentIds(model: DocModel): Set<number> {
  return new Set((model.metadata.comments ?? []).map((comment) => comment.id));
}

function stripOrphansFromPart(
  parts: MutablePartSet,
  partName: string,
  knownIds: ReadonlySet<number>
): void {
  const xml = parts.get(partName);
  if (xml === undefined) {
    return;
  }
  const stripped = stripOrphanCommentMarkers(xml, knownIds);
  if (stripped !== xml) {
    parts.set(partName, stripped);
  }
}

/**
 * Applies the model's comments to a package: strips stale anchors from the
 * document and header/footer parts, then writes or removes the comment parts.
 * Only called when the model knows about comments (`metadata.comments` is
 * defined), so a document whose comments the parser never reported is left
 * untouched.
 */
export function applyCommentsToPackage(
  pkg: OoxmlPackage,
  model: DocModel
): OoxmlPackage {
  const definitions = model.metadata.comments;
  if (definitions === undefined) {
    return pkg;
  }

  const parts = new Map(pkg.parts);
  const mutable: MutablePartSet = {
    get: (name) => parts.get(name)?.content,
    set: (name, content) => {
      parts.set(name, { name, content });
    },
    delete: (name) => {
      parts.delete(name);
    },
  };

  const knownIds = modelCommentIds(model);
  stripOrphansFromPart(mutable, "word/document.xml", knownIds);
  for (const partName of [...parts.keys()]) {
    if (/^word\/(?:header|footer)\d*\.xml$/.test(partName)) {
      stripOrphansFromPart(mutable, partName, knownIds);
    }
  }
  applyCommentsToParts(mutable, definitions);

  const unchanged =
    parts.size === pkg.parts.size &&
    [...parts.entries()].every(
      ([name, part]) => pkg.parts.get(name)?.content === part.content
    );
  if (unchanged) {
    return pkg;
  }

  return { parts, binaryAssets: new Map(pkg.binaryAssets) };
}

/** True when the model knows about comments and they must be written out. */
export function modelNeedsCommentSerialization(model: DocModel): boolean {
  return model.metadata.comments !== undefined;
}

/**
 * Applies the model's comments to an already-unzipped DOCX archive. The caller
 * owns `word/document.xml` (it has already stripped stale anchors there).
 */
export function applyCommentsToDocxArchive(
  archive: Record<string, Uint8Array>,
  model: DocModel
): void {
  const definitions = model.metadata.comments;
  if (definitions === undefined) {
    return;
  }

  const mutable: MutablePartSet = {
    get: (name) => (archive[name] ? strFromU8(archive[name]) : undefined),
    set: (name, content) => {
      archive[name] = strToU8(content);
    },
    delete: (name) => {
      delete archive[name];
    },
  };

  const knownIds = modelCommentIds(model);
  for (const partName of Object.keys(archive)) {
    if (/^word\/(?:header|footer)\d*\.xml$/.test(partName)) {
      stripOrphansFromPart(mutable, partName, knownIds);
    }
  }
  applyCommentsToParts(mutable, definitions);
}
