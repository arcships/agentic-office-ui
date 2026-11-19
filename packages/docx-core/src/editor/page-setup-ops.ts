// Page-setup and header/footer authoring commands.
//
// `metadata.sections` is the layout source of truth: `sectionPropertiesXml`
// drives pagination and the page surface, so every page-setup edit rewrites
// that XML. A mid-document section also carries the same fragment on the
// paragraph that ends it (see `engine/section-break.ts`), and the final section
// carries it in `metadata.sectionPropertiesXml` — the body-level `<w:sectPr>`
// the serializer writes. These commands keep all three representations in step
// so the change survives export/reload.

import type {
  DocModel,
  DocNode,
  DocumentSection,
  FooterSection,
  HeaderSection,
  ParagraphNode
} from "../engine/types";
import { cloneDocModelShallow } from "../engine/clone";
import {
  DEFAULT_SECTION_PROPERTIES_XML,
  readSectionColumns,
  readSectionPageMargins,
  readSectionPageSize,
  addSectionPartReference,
  writeSectionColumns,
  writeSectionPageMargins,
  writeSectionPageSize,
  type DocxPageMargins,
  type DocxPageMarginsUpdate,
  type DocxPageSize,
  type DocxPageSizeUpdate,
  type DocxSectionColumns,
  type DocxSectionColumnsUpdate
} from "../engine/section-properties";
import {
  paragraphSectionPropertiesXml,
  setParagraphSectionBreak
} from "../engine/section-break";
import type { DocxSectionRegion } from "./helpers/editor-types";

/** A section plus the page geometry it currently uses. */
export interface DocxSectionPageSetup {
  index: number;
  startNodeIndex: number;
  pageSize: DocxPageSize;
  margins: DocxPageMargins;
  /** Body columns of the section (`w:cols`), 1 when single-column. */
  columns: DocxSectionColumns;
  sectionPropertiesXml: string;
}

export interface AddHeaderFooterOptions {
  /**
   * OOXML header/footer variant: `default`, `first` (title page) or `even`.
   *
   * @defaultValue `"default"`
   */
  referenceType?: string;
}

const EMPTY_PARAGRAPH: ParagraphNode = { type: "paragraph", children: [] };

const HEADER_FOOTER_REFERENCE_TYPES = new Set(["default", "first", "even"]);

function normalizeReferenceType(referenceType?: string): string {
  const normalized = referenceType?.trim().toLowerCase();
  return normalized && HEADER_FOOTER_REFERENCE_TYPES.has(normalized)
    ? normalized
    : "default";
}

function normalizeSectionIndex(sectionIndex: number, sectionCount: number): number {
  const index = Number.isFinite(sectionIndex) ? Math.round(sectionIndex) : 0;
  return Math.max(0, Math.min(index, Math.max(0, sectionCount - 1)));
}

/** Deep copy of the section list so a command never mutates its input model. */
function cloneDocumentSections(model: DocModel): {
  model: DocModel;
  sections: DocumentSection[];
} {
  const next = cloneDocModelShallow(model);
  const sections: DocumentSection[] = (model.metadata.sections ?? []).map((section) => ({
    startNodeIndex: section.startNodeIndex,
    sectionPropertiesXml: section.sectionPropertiesXml,
    headerSections: (section.headerSections ?? []).map((entry) => ({
      ...entry,
      nodes: [...entry.nodes]
    })),
    footerSections: (section.footerSections ?? []).map((entry) => ({
      ...entry,
      nodes: [...entry.nodes]
    }))
  }));

  if (sections.length === 0) {
    // Documents without section metadata still have one implicit section: the
    // body-level `w:sectPr` in `metadata.sectionPropertiesXml`.
    sections.push({
      startNodeIndex: 0,
      sectionPropertiesXml: model.metadata.sectionPropertiesXml,
      headerSections: model.metadata.headerSections ?? [],
      footerSections: model.metadata.footerSections ?? []
    });
  }

  sections.sort((left, right) => left.startNodeIndex - right.startNodeIndex);
  next.metadata.sections = sections;
  return { model: next, sections };
}

/**
 * Index of the closest top-level paragraph before an exclusive node index —
 * the paragraph whose `<w:sectPr>` ends the section that stops there.
 */
function findSectionEndParagraphIndex(
  nodes: readonly DocNode[],
  endExclusive: number
): number | undefined {
  for (let index = endExclusive - 1; index >= 0; index -= 1) {
    if (nodes[index]?.type === "paragraph") {
      return index;
    }
  }
  return undefined;
}

/** The `w:sectPr` fragment a section currently uses. */
export function resolveSectionPropertiesXml(
  model: DocModel,
  sectionIndex: number
): string {
  const sections = model.metadata.sections ?? [];
  const resolvedIndex = sections.length > 0
    ? normalizeSectionIndex(sectionIndex, sections.length)
    : 0;
  const section = sections[resolvedIndex];
  const isFinalSection = sections.length === 0 || resolvedIndex === sections.length - 1;

  const candidates = [
    section?.sectionPropertiesXml,
    isFinalSection ? model.metadata.sectionPropertiesXml : undefined,
    isFinalSection
      ? undefined
      : sectionBreakSectionPropertiesXml(model, sections, resolvedIndex),
    DEFAULT_SECTION_PROPERTIES_XML
  ];

  for (const candidate of candidates) {
    if (candidate && candidate.trim().length > 0) {
      return candidate;
    }
  }
  return DEFAULT_SECTION_PROPERTIES_XML;
}

function sectionBreakSectionPropertiesXml(
  model: DocModel,
  sections: DocumentSection[],
  sectionIndex: number
): string | undefined {
  const nextStart = sections[sectionIndex + 1]?.startNodeIndex;
  if (nextStart === undefined) {
    return undefined;
  }
  const paragraphIndex = findSectionEndParagraphIndex(model.nodes, nextStart);
  const paragraph = paragraphIndex === undefined ? undefined : model.nodes[paragraphIndex];
  return paragraph?.type === "paragraph"
    ? paragraphSectionPropertiesXml(paragraph)
    : undefined;
}

/**
 * Write a section's `w:sectPr` fragment back to every place that carries it.
 *
 * The model object is mutated in place: callers pass a copy-on-write clone.
 */
function writeSectionPropertiesXml(
  model: DocModel,
  sections: DocumentSection[],
  sectionIndex: number,
  sectionPropertiesXml: string
): void {
  const section = sections[sectionIndex];
  if (!section) {
    return;
  }
  section.sectionPropertiesXml = sectionPropertiesXml;

  if (sectionIndex === sections.length - 1) {
    model.metadata.sectionPropertiesXml = sectionPropertiesXml;
    return;
  }

  const nextStart = sections[sectionIndex + 1]?.startNodeIndex;
  if (nextStart === undefined) {
    return;
  }
  const paragraphIndex = findSectionEndParagraphIndex(model.nodes, nextStart);
  if (paragraphIndex === undefined) {
    return;
  }
  const paragraph = model.nodes[paragraphIndex];
  if (paragraph?.type !== "paragraph") {
    return;
  }
  // The copy-on-write clone shares node objects with the source model, so the
  // edited paragraph is replaced by a clone before its marker changes.
  const clonedParagraph: ParagraphNode = {
    ...paragraph,
    children: [...paragraph.children]
  };
  model.nodes[paragraphIndex] = clonedParagraph;
  setParagraphSectionBreak(clonedParagraph, sectionPropertiesXml);
}

function applySectionPropertiesUpdate(
  model: DocModel,
  sectionIndex: number,
  update: (sectionPropertiesXml: string) => string
): DocModel {
  const { model: next, sections } = cloneDocumentSections(model);
  const resolvedIndex = normalizeSectionIndex(sectionIndex, sections.length);
  const current = resolveSectionPropertiesXml(next, resolvedIndex);
  const updated = update(current);
  if (updated === current) {
    return model;
  }
  writeSectionPropertiesXml(next, sections, resolvedIndex, updated);
  return next;
}

/**
 * Set page margins (twips) on one section. Unspecified margins keep their
 * current value and every other `w:sectPr` child is preserved.
 */
export function setPageMargins(
  model: DocModel,
  sectionIndex: number,
  margins: DocxPageMarginsUpdate
): DocModel {
  return applySectionPropertiesUpdate(model, sectionIndex, (sectionPropertiesXml) =>
    writeSectionPageMargins(sectionPropertiesXml, margins)
  );
}

/**
 * Set the paper size/orientation of one section. `orientation` swaps the
 * width/height pair when they contradict it, matching the way Word stores a
 * landscape page.
 */
export function setPageSize(
  model: DocModel,
  sectionIndex: number,
  pageSize: DocxPageSizeUpdate
): DocModel {
  return applySectionPropertiesUpdate(model, sectionIndex, (sectionPropertiesXml) =>
    writeSectionPageSize(sectionPropertiesXml, pageSize)
  );
}

/**
 * Set the body columns of one section (`w:cols`). Changing the count rewrites
 * `w:num` and clears explicit per-column widths, so the section flows through
 * the requested number of equal columns; every other `sectPr` child is
 * preserved. One undo entry.
 */
export function setSectionColumns(
  model: DocModel,
  sectionIndex: number,
  columns: DocxSectionColumnsUpdate
): DocModel {
  return applySectionPropertiesUpdate(model, sectionIndex, (sectionPropertiesXml) =>
    writeSectionColumns(sectionPropertiesXml, columns)
  );
}

/** Page geometry of every section, in document order. */
export function resolvePageSetup(model: DocModel): DocxSectionPageSetup[] {
  const sections = model.metadata.sections ?? [];
  if (sections.length === 0) {
    const sectionPropertiesXml = resolveSectionPropertiesXml(model, 0);
    return [
      {
        index: 0,
        startNodeIndex: 0,
        pageSize: readSectionPageSize(sectionPropertiesXml),
        margins: readSectionPageMargins(sectionPropertiesXml),
        columns: readSectionColumns(sectionPropertiesXml),
        sectionPropertiesXml
      }
    ];
  }

  return sections.map((section, index) => {
    const sectionPropertiesXml = resolveSectionPropertiesXml(model, index);
    return {
      index,
      startNodeIndex: section.startNodeIndex,
      pageSize: readSectionPageSize(sectionPropertiesXml),
      margins: readSectionPageMargins(sectionPropertiesXml),
      columns: readSectionColumns(sectionPropertiesXml),
      sectionPropertiesXml
    };
  });
}

function sectionParts(
  model: DocModel,
  region: DocxSectionRegion
): HeaderSection[] | FooterSection[] {
  return region === "header"
    ? model.metadata.headerSections ?? []
    : model.metadata.footerSections ?? [];
}

function allSectionParts(model: DocModel): Array<HeaderSection | FooterSection> {
  const perSection = (model.metadata.sections ?? []).flatMap((section) => [
    ...(section.headerSections ?? []),
    ...(section.footerSections ?? [])
  ]);
  return [
    ...perSection,
    ...(model.metadata.headerSections ?? []),
    ...(model.metadata.footerSections ?? [])
  ];
}

/** Next free part name for a header/footer part, e.g. `word/header3.xml`. */
function nextSectionPartName(model: DocModel, region: DocxSectionRegion): string {
  const used = new Set(
    allSectionParts(model).map((entry) => entry.partName.trim().toLowerCase())
  );
  let index = 1;
  for (;;) {
    const candidate = `word/${region}${index}.xml`;
    if (!used.has(candidate)) {
      return candidate;
    }
    index += 1;
  }
}

/** Relationship id used by a section part reference, unique per document. */
function nextSectionPartRelationshipId(model: DocModel, partName: string): string {
  const used = new Set<string>();
  const collect = (xml?: string) => {
    if (!xml) return;
    for (const match of xml.matchAll(/\br:id\s*=\s*"([^"]+)"/gi)) {
      used.add(match[1]);
    }
  };

  collect(model.metadata.sectionPropertiesXml);
  (model.metadata.sections ?? []).forEach((section) => collect(section.sectionPropertiesXml));
  model.nodes.forEach((node) => {
    collect(node.sourceXml);
    if (node.type === "paragraph") collect(node.sectionBreak?.sectionPropertiesXml);
  });
  allSectionParts(model).forEach((entry) =>
    entry.nodes.forEach((node) => {
      collect(node.sourceXml);
      if (node.type === "paragraph") collect(node.sectionBreak?.sectionPropertiesXml);
    })
  );

  // Word writes `rId<number>` for its own relationships; the readable prefix
  // keeps a generated id clear of those while staying a valid NCName.
  const kind = partName.includes("/footer") ? "Ftr" : "Hdr";
  const stem = `rId${kind}${partName.match(/(\d+)\.xml$/)?.[1] ?? "1"}`;
  let candidate = stem;
  let suffix = 1;
  while (used.has(candidate)) {
    suffix += 1;
    candidate = `${stem}_${suffix}`;
  }
  return candidate;
}

/**
 * Create a header/footer part for a section and reference it from the section's
 * `w:sectPr`.
 *
 * The new part starts with one empty paragraph (what Word inserts for "Edit
 * header") and is added to `metadata.sections[index]`, `metadata.headerSections`
 * / `metadata.footerSections` and — through the `w:headerReference` /
 * `w:footerReference` the command writes — to the section properties. Adding a
 * variant the section already has is a no-op: the returned model is the input
 * model, so the caller can treat it as "nothing to do, edit the existing part".
 */
export function addHeaderFooter(
  model: DocModel,
  sectionIndex: number,
  kind: DocxSectionRegion,
  options: AddHeaderFooterOptions = {}
): DocModel {
  const { model: next, sections } = cloneDocumentSections(model);
  const resolvedIndex = normalizeSectionIndex(sectionIndex, sections.length);
  const section = sections[resolvedIndex];
  if (!section) {
    return model;
  }

  const referenceType = normalizeReferenceType(options.referenceType);
  const targetKey = kind === "header" ? "headerSections" : "footerSections";
  const existing = (section[targetKey] ?? []).find(
    (entry) => normalizeReferenceType(entry.referenceType) === referenceType
  );
  if (existing) {
    return model;
  }

  const partName = nextSectionPartName(model, kind);
  const relationshipId = nextSectionPartRelationshipId(model, partName);
  const part: HeaderSection = {
    partName,
    referenceType,
    nodes: [{ ...EMPTY_PARAGRAPH, children: [] }]
  };

  if (kind === "header") {
    section.headerSections = [...(section.headerSections ?? []), part];
  } else {
    section.footerSections = [...(section.footerSections ?? []), part as FooterSection];
  }

  // The top-level lists are the part source for the serializer and the editing
  // fallback; both must know the new part.
  const topLevel = sectionParts(next, kind);
  if (!topLevel.some((entry) => entry.partName === partName)) {
    if (kind === "header") {
      next.metadata.headerSections = [...topLevel, part];
    } else {
      next.metadata.footerSections = [...topLevel, part as FooterSection];
    }
  }

  const current = resolveSectionPropertiesXml(next, resolvedIndex);
  const updated = addSectionPartReference(
    current,
    kind,
    referenceType,
    relationshipId
  );
  writeSectionPropertiesXml(next, sections, resolvedIndex, updated);

  return next;
}
