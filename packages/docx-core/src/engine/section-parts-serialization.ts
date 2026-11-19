// Package-level serialization of header/footer parts.
//
// The DOCX serializer writes a header/footer part for every entry in
// `metadata.headerSections` / `metadata.footerSections` and emits the section's
// `<w:headerReference r:id="…"/>` / `<w:footerReference r:id="…"/>` unchanged.
// The package wrapper is copied from the source document, so a part created by
// `addHeaderFooter()` has no relationship in `word/_rels/document.xml.rels` and
// no `[Content_Types].xml` override yet. Without both, Word reports unreadable
// content even though `word/header2.xml` exists.
//
// These helpers add exactly the missing entries; a document that already
// declares its parts is returned untouched, so an unmodified export keeps the
// wrapper it came with.

import type { DocModel, FooterSection, HeaderSection } from "./types";
import { paragraphSectionPropertiesXml } from "./section-break";

const HEADER_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const HEADER_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

export const DOCX_DOCUMENT_RELATIONSHIPS_PART = "word/_rels/document.xml.rels";
export const DOCX_CONTENT_TYPES_PART = "[Content_Types].xml";

export type DocxSectionPartKind = "header" | "footer";

export interface DocxSectionPartReference {
  kind: DocxSectionPartKind;
  partName: string;
  /** `r:id` the section properties use for this part, when referenced. */
  relationshipId?: string;
  referenceType?: string;
}

function normalizeReferenceType(value?: string): string {
  const normalized = value?.trim().toLowerCase();
  return normalized && normalized.length > 0 ? normalized : "default";
}

/**
 * Closest top-level paragraph before an exclusive node index — the paragraph
 * whose `<w:sectPr>` ends the section that stops there.
 */
function sectionBreakXmlBefore(model: DocModel, endExclusive: number): string | undefined {
  for (let index = endExclusive - 1; index >= 0; index -= 1) {
    const node = model.nodes[index];
    if (node?.type === "paragraph") {
      return paragraphSectionPropertiesXml(node);
    }
  }
  return undefined;
}

interface SectionStoryContext {
  sectionPropertiesXml?: string;
  headerSections: HeaderSection[];
  footerSections: FooterSection[];
}

/**
 * Pair every `w:sectPr` fragment with the header/footer parts that belong to
 * its own section, so two sections that both use a "default" header resolve to
 * their own parts.
 */
function collectSectionStoryContexts(model: DocModel): SectionStoryContext[] {
  const sections = model.metadata.sections ?? [];
  if (sections.length === 0) {
    return [
      {
        sectionPropertiesXml: model.metadata.sectionPropertiesXml,
        headerSections: model.metadata.headerSections ?? [],
        footerSections: model.metadata.footerSections ?? [],
      },
    ];
  }

  const contexts = sections.map((section, index) => {
    const following = sections[index + 1];
    return {
      sectionPropertiesXml:
        section.sectionPropertiesXml ??
        (following ? sectionBreakXmlBefore(model, following.startNodeIndex) : undefined),
      headerSections: section.headerSections ?? [],
      footerSections: section.footerSections ?? [],
    };
  });

  const last = contexts[contexts.length - 1];
  if (
    model.metadata.sectionPropertiesXml &&
    model.metadata.sectionPropertiesXml !== last.sectionPropertiesXml
  ) {
    // The body-level `w:sectPr` is what the serializer writes for the final
    // section; it can differ from the model entry after a hand-written edit.
    contexts.push({
      sectionPropertiesXml: model.metadata.sectionPropertiesXml,
      headerSections: last.headerSections,
      footerSections: last.footerSections,
    });
  }

  return contexts;
}

/**
 * Header/footer parts the model references, paired with the `r:id` each
 * section properties fragment uses. Parts without a reference are reported too
 * (they still need a content-type override): their `relationshipId` is absent.
 */
export function collectSectionPartReferences(model: DocModel): DocxSectionPartReference[] {
  const references = new Map<string, DocxSectionPartReference>();
  const add = (reference: DocxSectionPartReference) => {
    const key = `${reference.kind}:${reference.partName.toLowerCase()}`;
    const existing = references.get(key);
    if (!existing) {
      references.set(key, reference);
      return;
    }
    if (!existing.relationshipId && reference.relationshipId) {
      references.set(key, { ...existing, relationshipId: reference.relationshipId });
    }
  };

  // Every part the model knows about (referenced or not) needs a content type.
  const allHeaderParts = new Map<string, HeaderSection>();
  const allFooterParts = new Map<string, FooterSection>();
  const registerParts = (headers: HeaderSection[] = [], footers: FooterSection[] = []) => {
    headers.forEach((entry) => allHeaderParts.set(entry.partName.toLowerCase(), entry));
    footers.forEach((entry) => allFooterParts.set(entry.partName.toLowerCase(), entry));
  };
  registerParts(model.metadata.headerSections, model.metadata.footerSections);
  (model.metadata.sections ?? []).forEach((section) =>
    registerParts(section.headerSections, section.footerSections),
  );
  allHeaderParts.forEach((entry) =>
    add({ kind: "header", partName: entry.partName, referenceType: entry.referenceType }),
  );
  allFooterParts.forEach((entry) =>
    add({ kind: "footer", partName: entry.partName, referenceType: entry.referenceType }),
  );

  const partNameFor = (
    entries: Array<HeaderSection | FooterSection>,
    referenceType: string,
  ): string | undefined =>
    entries.find(
      (entry) => normalizeReferenceType(entry.referenceType) === referenceType,
    )?.partName;

  for (const context of collectSectionStoryContexts(model)) {
    const xml = context.sectionPropertiesXml;
    if (!xml) continue;

    for (const match of xml.matchAll(/<w:(header|footer)Reference\b[^>]*>/gi)) {
      const referenceTag = match[0];
      const kind = match[1].toLowerCase() === "header" ? "header" : "footer";
      const relationshipId = referenceTag.match(/\br:id\s*=\s*"([^"]+)"/i)?.[1];
      if (!relationshipId) continue;

      const referenceType = normalizeReferenceType(
        referenceTag.match(/\bw:type\s*=\s*"([^"]*)"/i)?.[1],
      );
      const entries: Array<HeaderSection | FooterSection> =
        kind === "header"
          ? [...context.headerSections, ...allHeaderParts.values()]
          : [...context.footerSections, ...allFooterParts.values()];
      const partName = partNameFor(entries, referenceType);
      if (!partName) continue;

      add({ kind, partName, relationshipId, referenceType });
    }
  }

  return [...references.values()];
}

function escapeXmlAttribute(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function decodeXmlAttribute(value: string): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

/** Relationship targets are relative to `word/document.xml`. */
function relationshipTarget(partName: string): string {
  const prefix = "word/";
  return partName.startsWith(prefix) ? partName.slice(prefix.length) : partName;
}

interface ExistingRelationship {
  id: string;
  type: string;
  target: string;
}

function readRelationships(relsXml: string): ExistingRelationship[] {
  return [...relsXml.matchAll(/<Relationship\b[^>]*?\/?>/gi)].map((match) => ({
    id: decodeXmlAttribute(match[0].match(/\bId\s*=\s*"([^"]*)"/i)?.[1] ?? ""),
    type: decodeXmlAttribute(match[0].match(/\bType\s*=\s*"([^"]*)"/i)?.[1] ?? ""),
    target: decodeXmlAttribute(match[0].match(/\bTarget\s*=\s*"([^"]*)"/i)?.[1] ?? ""),
  }));
}

export interface SectionPartRelationshipPatch {
  xml: string;
  /** `r:id` renames the document part must apply (collision safety). */
  renamedIds: Map<string, string>;
}

/**
 * Add the missing `header`/`footer` relationships to
 * `word/_rels/document.xml.rels`.
 *
 * A reference whose id is already taken by a different relationship is renamed
 * to a free id; the caller must rewrite `word/document.xml` with the returned
 * map so the reference and the relationship keep pointing at each other.
 */
export function applySectionPartRelationships(
  relsXml: string,
  references: DocxSectionPartReference[],
): SectionPartRelationshipPatch {
  const relationships = readRelationships(relsXml);
  const byId = new Map(relationships.map((relationship) => [relationship.id, relationship]));
  const renamedIds = new Map<string, string>();
  const additions: string[] = [];
  let nextNumericId =
    relationships.reduce((max, relationship) => {
      const numeric = Number(relationship.id.match(/^rId(\d+)$/)?.[1] ?? "");
      return Number.isFinite(numeric) ? Math.max(max, numeric) : max;
    }, 0) + 1;

  for (const reference of references) {
    if (!reference.relationshipId) continue;
    const type =
      reference.kind === "header" ? HEADER_RELATIONSHIP_TYPE : FOOTER_RELATIONSHIP_TYPE;
    const target = relationshipTarget(reference.partName);
    const existing = byId.get(reference.relationshipId);

    if (
      existing &&
      existing.type === type &&
      existing.target.toLowerCase() === target.toLowerCase()
    ) {
      continue;
    }

    let relationshipId = reference.relationshipId;
    if (existing) {
      // The id belongs to another relationship: move this reference to a free
      // id instead of hijacking it.
      do {
        relationshipId = `rId${nextNumericId}`;
        nextNumericId += 1;
      } while (byId.has(relationshipId));
      renamedIds.set(reference.relationshipId, relationshipId);
    }
    byId.set(relationshipId, { id: relationshipId, type, target });

    additions.push(
      `<Relationship Id="${escapeXmlAttribute(relationshipId)}" ` +
        `Type="${escapeXmlAttribute(type)}" ` +
        `Target="${escapeXmlAttribute(target)}"/>`,
    );
  }

  if (additions.length === 0) {
    return { xml: relsXml, renamedIds };
  }

  const closing = relsXml.lastIndexOf("</Relationships>");
  if (closing < 0) {
    return { xml: relsXml, renamedIds: new Map() };
  }
  return {
    xml: relsXml.slice(0, closing) + additions.join("") + relsXml.slice(closing),
    renamedIds,
  };
}

/** Add the missing `header`/`footer` `<Override>` content types. */
export function applySectionPartContentTypes(
  contentTypesXml: string,
  references: DocxSectionPartReference[],
): string {
  const declarations = contentTypesXml.toLowerCase();
  const additions: string[] = [];

  for (const reference of references) {
    const partName = `/${reference.partName.replace(/^\/+/, "")}`;
    if (declarations.includes(`partname="${partName.toLowerCase()}"`)) {
      continue;
    }
    additions.push(
      `<Override PartName="${escapeXmlAttribute(partName)}" ContentType="${
        reference.kind === "header" ? HEADER_CONTENT_TYPE : FOOTER_CONTENT_TYPE
      }"/>`,
    );
  }

  if (additions.length === 0) {
    return contentTypesXml;
  }

  const closing = contentTypesXml.lastIndexOf("</Types>");
  if (closing < 0) {
    return contentTypesXml;
  }
  return contentTypesXml.slice(0, closing) + additions.join("") + contentTypesXml.slice(closing);
}

/** True when the model carries at least one header/footer part. */
export function modelHasSectionParts(model: DocModel): boolean {
  return (
    (model.metadata.headerSections?.length ?? 0) > 0 ||
    (model.metadata.footerSections?.length ?? 0) > 0 ||
    (model.metadata.sections ?? []).some(
      (section) =>
        (section.headerSections?.length ?? 0) > 0 ||
        (section.footerSections?.length ?? 0) > 0,
    )
  );
}
