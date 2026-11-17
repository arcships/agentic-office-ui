// Paragraph-level section-break preservation helpers.
//
// A mid-document section break lives in the `<w:sectPr>` child of the ending
// paragraph's `<w:pPr>`. Editing that paragraph clears `sourceXml`, so the raw
// XML can no longer be the only place the break exists. These helpers promote
// the break onto the paragraph model while leaving `metadata.sections` as the
// layout/serialization source of truth.

import type {
  DocModel,
  ParagraphNode,
  ParagraphSectionBreak,
  SectionBreakType
} from "./types";

const SECTION_PROPERTIES_XML_PATTERN = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/i;
const SELF_CLOSING_SECTION_PROPERTIES_XML_PATTERN = /<w:sectPr\b[^>]*\/>/i;
const SECTION_TYPE_XML_PATTERN = /<w:type\b[^>]*w:val="([^"]*)"/i;

export function sectionPropertiesXmlFromSourceXml(
  sourceXml?: string
): string | undefined {
  if (!sourceXml) {
    return undefined;
  }

  return (
    sourceXml.match(SECTION_PROPERTIES_XML_PATTERN)?.[0] ??
    sourceXml.match(SELF_CLOSING_SECTION_PROPERTIES_XML_PATTERN)?.[0]
  );
}

export function normalizeSectionBreakType(value?: string): SectionBreakType {
  switch (value?.trim().toLowerCase()) {
    case "continuous":
      return "continuous";
    case "evenpage":
      return "evenPage";
    case "oddpage":
      return "oddPage";
    case "nextcolumn":
      return "nextColumn";
    default:
      // ECMA-376 §2.6.22: an omitted <w:type> defaults to nextPage.
      return "nextPage";
  }
}

export function parseSectionBreakTypeFromPropertiesXml(
  sectionPropertiesXml?: string
): SectionBreakType {
  return normalizeSectionBreakType(
    sectionPropertiesXml?.match(SECTION_TYPE_XML_PATTERN)?.[1]
  );
}

export function sectionBreakFromSourceXml(
  sourceXml?: string
): ParagraphSectionBreak | undefined {
  const sectionPropertiesXml = sectionPropertiesXmlFromSourceXml(sourceXml);
  return sectionPropertiesXml ? { sectionPropertiesXml } : undefined;
}

/** The section break ending at this paragraph, if any. */
export function paragraphSectionBreak(
  paragraph: ParagraphNode
): ParagraphSectionBreak | undefined {
  if (paragraph.sectionBreak) {
    return paragraph.sectionBreak;
  }

  return sectionBreakFromSourceXml(paragraph.sourceXml);
}

export function paragraphSectionPropertiesXml(
  paragraph: ParagraphNode
): string | undefined {
  return paragraphSectionBreak(paragraph)?.sectionPropertiesXml;
}

export function paragraphHasSectionBreak(paragraph: ParagraphNode): boolean {
  return paragraphSectionBreak(paragraph) !== undefined;
}

/**
 * Promote `<w:sectPr>` out of `sourceXml` before the XML is discarded. Existing
 * markers win because they may already be the reconciled model value.
 */
export function preserveParagraphSectionBreak(paragraph: ParagraphNode): void {
  if (paragraph.sectionBreak) {
    return;
  }

  const sectionBreak = sectionBreakFromSourceXml(paragraph.sourceXml);
  if (sectionBreak) {
    paragraph.sectionBreak = sectionBreak;
  }
}

/** Clear `sourceXml` without losing the section break it carries. */
export function invalidateParagraphSourceXml(paragraph: ParagraphNode): void {
  preserveParagraphSectionBreak(paragraph);
  paragraph.sourceXml = undefined;
}

export function setSectionBreakTypeInPropertiesXml(
  sectionPropertiesXml: string,
  type: SectionBreakType
): string {
  const typeTag = `<w:type w:val="${type}"/>`;

  if (SECTION_TYPE_XML_PATTERN.test(sectionPropertiesXml)) {
    return sectionPropertiesXml.replace(SECTION_TYPE_XML_PATTERN, typeTag);
  }

  const selfClosingSectPr = sectionPropertiesXml.match(/<w:sectPr\b[^>]*\/>/i);
  if (selfClosingSectPr) {
    return sectionPropertiesXml.replace(
      selfClosingSectPr[0],
      `${selfClosingSectPr[0].replace(/\/>\s*$/, ">")}${typeTag}</w:sectPr>`
    );
  }

  // CT_SectPr places <w:type> after header/footer references and before
  // <w:pgSz>. Insert there so Word does not have to repair the document.
  const pageSizeIndex = sectionPropertiesXml.search(/<w:pgSz\b/i);
  if (pageSizeIndex >= 0) {
    return (
      sectionPropertiesXml.slice(0, pageSizeIndex) +
      typeTag +
      sectionPropertiesXml.slice(pageSizeIndex)
    );
  }

  if (/<\/w:sectPr>\s*$/i.test(sectionPropertiesXml)) {
    return sectionPropertiesXml.replace(
      /<\/w:sectPr>\s*$/i,
      `${typeTag}</w:sectPr>`
    );
  }

  return sectionPropertiesXml;
}

/** Remove every `<w:sectPr>` fragment from a paragraph XML string. */
export function removeSectionPropertiesXmlFromSourceXml(
  sourceXml?: string
): string | undefined {
  if (!sourceXml) {
    return undefined;
  }

  return sourceXml
    .replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/gi, "")
    .replace(/<w:sectPr\b[^>]*\/>/gi, "");
}

function replaceSectionPropertiesXmlInSourceXml(
  sourceXml: string,
  sectionPropertiesXml: string
): string {
  if (/<w:sectPr\b[\s\S]*?<\/w:sectPr>/i.test(sourceXml)) {
    return sourceXml.replace(
      /<w:sectPr\b[\s\S]*?<\/w:sectPr>/i,
      sectionPropertiesXml
    );
  }

  if (/<w:sectPr\b[^>]*\/>/i.test(sourceXml)) {
    return sourceXml.replace(/<w:sectPr\b[^>]*\/>/i, sectionPropertiesXml);
  }

  return sourceXml;
}

/**
 * Attach or replace the section break ending at a paragraph. When the source
 * XML cannot carry the fragment it is invalidated so serialization falls back
 * to the model (and the serializer re-attaches the preserved break).
 */
export function setParagraphSectionBreak(
  paragraph: ParagraphNode,
  sectionPropertiesXml: string
): void {
  paragraph.sectionBreak = { sectionPropertiesXml };

  if (paragraph.sourceXml) {
    const replaced = replaceSectionPropertiesXmlInSourceXml(
      paragraph.sourceXml,
      sectionPropertiesXml
    );
    paragraph.sourceXml = replaced === paragraph.sourceXml
      ? undefined
      : replaced;
  }
}

/** Remove the section break ending at a paragraph without clearing its text. */
export function clearParagraphSectionBreak(paragraph: ParagraphNode): void {
  paragraph.sectionBreak = undefined;
  paragraph.sourceXml = removeSectionPropertiesXmlFromSourceXml(
    paragraph.sourceXml
  );
}

/**
 * Keep `metadata.sections[].startNodeIndex` aligned when nodes are inserted or
 * removed. `fromNodeIndex` is the insertion/removal point; `delta` is positive
 * for insertions and negative for removals.
 *
 * Insertions and removals both shift sections that start after the affected
 * node. A section that starts exactly at the insertion/removal point keeps its
 * start index: the inserted node becomes its new first node, and after removal
 * the following node becomes its first node.
 */
export function shiftDocumentSectionStartNodeIndexes(
  model: DocModel,
  fromNodeIndex: number,
  delta: number
): void {
  if (!model.metadata.sections || delta === 0) {
    return;
  }

  const threshold = fromNodeIndex + 1;
  model.metadata.sections.forEach((section) => {
    if (section.startNodeIndex >= threshold) {
      section.startNodeIndex = Math.max(0, section.startNodeIndex + delta);
    }
  });
}
