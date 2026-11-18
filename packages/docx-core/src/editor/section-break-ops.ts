// Section-break insertion/deletion commands.
//
// `metadata.sections` stays the source of truth used by layout and header/footer
// resolution. A break is represented twice while it is structural: as the raw
// `<w:sectPr>` marker on the paragraph that ends the section, and as the
// matching `metadata.sections` entry for that section. These commands keep both
// representations and every later `startNodeIndex` consistent.

import type {
  DocModel,
  DocumentSection,
  ParagraphNode,
  SectionBreakType
} from "../engine/types";
import { cloneDocModel } from "../engine/clone";
import {
  clearParagraphSectionBreak,
  normalizeSectionBreakType,
  paragraphSectionBreak,
  setParagraphSectionBreak,
  setSectionBreakTypeInPropertiesXml
} from "../engine/section-break";
import { getParagraph, paragraphFromText } from "./paragraph-ops";

const DEFAULT_SECTION_PROPERTIES_XML =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

function cloneSectionHeaderFooter(
  section: DocumentSection
): Pick<DocumentSection, "headerSections" | "footerSections"> {
  return {
    headerSections: section.headerSections.map((header) => ({ ...header })),
    footerSections: section.footerSections.map((footer) => ({ ...footer }))
  };
}

function ensureDocumentSections(model: DocModel): DocumentSection[] {
  const sections = model.metadata.sections;
  if (sections && sections.length > 0) {
    sections.sort((left, right) => left.startNodeIndex - right.startNodeIndex);
    return sections;
  }

  const synthesized: DocumentSection[] = [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: model.metadata.sectionPropertiesXml,
      headerSections: model.metadata.headerSections ?? [],
      footerSections: model.metadata.footerSections ?? []
    }
  ];
  model.metadata.sections = synthesized;
  return synthesized;
}

function resolveSectionIndexForNode(
  sections: DocumentSection[],
  nodeIndex: number
): number {
  let sectionIndex = 0;
  for (let index = 0; index < sections.length; index += 1) {
    if (sections[index].startNodeIndex <= nodeIndex) {
      sectionIndex = index;
    } else {
      break;
    }
  }
  return sectionIndex;
}

function sectionPropertiesForIndex(
  model: DocModel,
  sections: DocumentSection[],
  sectionIndex: number
): string {
  return (
    sections[sectionIndex]?.sectionPropertiesXml ??
    (sectionIndex === sections.length - 1
      ? model.metadata.sectionPropertiesXml
      : undefined) ??
    DEFAULT_SECTION_PROPERTIES_XML
  );
}

/**
 * Closest top-level paragraph before an exclusive node index. This is the
 * paragraph whose `<w:sectPr>` ends the section that stops at that node.
 */
function findSectionEndParagraph(
  model: DocModel,
  endExclusive: number
): ParagraphNode | undefined {
  for (let index = endExclusive - 1; index >= 0; index -= 1) {
    const node = model.nodes[index];
    if (node?.type === "paragraph") {
      return node;
    }
  }
  return undefined;
}

/**
 * Insert a section break after `nodeIndex` (the paragraph becomes the last
 * paragraph of the current section). `metadata.sections`, the preserved marker
 * and the next section's properties are all kept in sync.
 */
export function insertSectionBreak(
  model: DocModel,
  nodeIndex: number,
  type: SectionBreakType = "nextPage"
): DocModel {
  const next = cloneDocModel(model);
  const paragraph = getParagraph(next, nodeIndex);
  if (!paragraph) {
    return next;
  }

  const breakType = normalizeSectionBreakType(type);
  const sections = ensureDocumentSections(next);
  const sectionIndex = resolveSectionIndexForNode(sections, nodeIndex);
  const explicitBreak = paragraphSectionBreak(paragraph);
  const currentProperties =
    explicitBreak?.sectionPropertiesXml ??
    sectionPropertiesForIndex(next, sections, sectionIndex);

  sections[sectionIndex].sectionPropertiesXml = currentProperties;

  const followingSection = sections[sectionIndex + 1];
  if (followingSection && followingSection.startNodeIndex === nodeIndex + 1) {
    // The paragraph already ends a section: treat the command as changing the
    // type of the section that starts after it.
    const updatedProperties = setSectionBreakTypeInPropertiesXml(
      followingSection.sectionPropertiesXml ??
        sectionPropertiesForIndex(next, sections, sectionIndex + 1),
      breakType
    );
    followingSection.sectionPropertiesXml = updatedProperties;
    setParagraphSectionBreak(paragraph, currentProperties);

    const followingEnd = sections[sectionIndex + 2]?.startNodeIndex;
    if (followingEnd !== undefined) {
      const followingEndParagraph = findSectionEndParagraph(next, followingEnd);
      if (followingEndParagraph) {
        setParagraphSectionBreak(followingEndParagraph, updatedProperties);
      }
    } else {
      next.metadata.sectionPropertiesXml = updatedProperties;
    }

    return next;
  }

  if (nodeIndex === next.nodes.length - 1) {
    next.nodes.push(paragraphFromText(""));
  }

  const newSectionProperties = setSectionBreakTypeInPropertiesXml(
    currentProperties,
    breakType
  );
  setParagraphSectionBreak(paragraph, currentProperties);

  if (followingSection) {
    const previousEndParagraph = findSectionEndParagraph(
      next,
      followingSection.startNodeIndex
    );
    if (previousEndParagraph) {
      setParagraphSectionBreak(previousEndParagraph, newSectionProperties);
    }
  } else {
    next.metadata.sectionPropertiesXml = newSectionProperties;
  }

  sections.splice(sectionIndex + 1, 0, {
    startNodeIndex: nodeIndex + 1,
    sectionPropertiesXml: newSectionProperties,
    ...cloneSectionHeaderFooter(sections[sectionIndex])
  });

  return next;
}

/**
 * Delete the section break after `nodeIndex` and merge the current section into
 * the following one. As in Word, the merged section adopts the following
 * section's page setup and header/footer references.
 */
export function deleteSectionBreak(
  model: DocModel,
  nodeIndex: number
): DocModel {
  const next = cloneDocModel(model);
  const paragraph = getParagraph(next, nodeIndex);
  if (!paragraph || !paragraphSectionBreak(paragraph)) {
    return next;
  }

  clearParagraphSectionBreak(paragraph);

  const sections = next.metadata.sections;
  if (!sections || sections.length === 0) {
    return next;
  }

  const followingIndex = sections.findIndex(
    (section) => section.startNodeIndex === nodeIndex + 1
  );
  if (followingIndex > 0) {
    const currentSection = sections[followingIndex - 1];
    sections[followingIndex].startNodeIndex = currentSection.startNodeIndex;
    sections.splice(followingIndex - 1, 1);
  }

  return next;
}
