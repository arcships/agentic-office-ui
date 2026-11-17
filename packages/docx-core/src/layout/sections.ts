/**
 * Section metadata helpers shared by the layout engine, pagination and page
 * segmentation.
 *
 * This module deliberately has no module-level caches: it is reachable from
 * the platform-neutral `core.ts` runtime graph through `layout-engine.ts`,
 * where top-level mutable state is not allowed.
 */

import type {
  DocModel,
  FooterSection,
  HeaderSection
} from "../engine/types";

const SECTION_TYPE_XML_PATTERN = /<w:type\b[^>]*w:val="([^"]+)"/i;

export interface ResolvedModelSection {
  startNodeIndex: number;
  sectionPropertiesXml?: string;
  headerSections: HeaderSection[];
  footerSections: FooterSection[];
}

export type SectionStartPageParity = "even" | "odd";

function normalizeSectionReferenceType(referenceType?: string): string {
  const normalized = referenceType?.trim().toLowerCase();
  return normalized && normalized.length > 0 ? normalized : "default";
}

function inheritSectionReferences<T extends HeaderSection | FooterSection>(
  sections: ResolvedModelSection[],
  sectionKey: "headerSections" | "footerSections"
): ResolvedModelSection[] {
  const inheritedByType = new Map<string, T>();

  return sections.map((section) => {
    const explicitSections = section[sectionKey] as T[];
    if (explicitSections.length > 0) {
      explicitSections.forEach((entry) => {
        inheritedByType.set(normalizeSectionReferenceType(entry.referenceType), entry);
      });
    }

    return {
      ...section,
      [sectionKey]: [...inheritedByType.values()]
    };
  });
}

function resolveInheritedSectionHeaderFooterReferences(
  sections: ResolvedModelSection[]
): ResolvedModelSection[] {
  if (sections.length === 0) {
    return sections;
  }

  return inheritSectionReferences(
    inheritSectionReferences(sections, "headerSections"),
    "footerSections"
  );
}

export function sectionBreakPropertiesStartNewPage(sectionPropertiesXml: string): boolean {
  const sectionType =
    sectionPropertiesXml.match(SECTION_TYPE_XML_PATTERN)?.[1]?.trim().toLowerCase() ?? "nextpage";

  if (sectionType === "continuous") {
    return false;
  }

  if (sectionType === "nextcolumn") {
    const columnsTag = sectionPropertiesXml.match(/<w:cols\b[^>]*\/?>/i)?.[0];
    const columnsCount = Number.parseInt(columnsTag?.match(/\bw:num="(\d+)"/i)?.[1] ?? "", 10);
    return !Number.isFinite(columnsCount) || columnsCount <= 1;
  }

  return true;
}

/**
 * Word page-number parity a section break must land on. `evenPage` requires
 * the next section to start on an even-numbered page, `oddPage` on an odd one;
 * both may need a blank page inserted when the previous section ends on the
 * wrong parity.
 */
export function sectionBreakRequiredPageParity(
  sectionPropertiesXml?: string
): SectionStartPageParity | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const sectionType =
    sectionPropertiesXml.match(SECTION_TYPE_XML_PATTERN)?.[1]?.trim().toLowerCase();
  if (sectionType === "evenpage") {
    return "even";
  }
  if (sectionType === "oddpage") {
    return "odd";
  }

  return undefined;
}

export function resolveDocumentSectionsFromMetadata(
  metadata: DocModel["metadata"]
): ResolvedModelSection[] {
  const normalizedSections = (metadata.sections ?? [])
    .map((section): ResolvedModelSection => ({
      startNodeIndex:
        Number.isFinite(section.startNodeIndex) && (section.startNodeIndex as number) >= 0
          ? Math.round(section.startNodeIndex as number)
          : 0,
      sectionPropertiesXml: section.sectionPropertiesXml,
      headerSections: section.headerSections ?? [],
      footerSections: section.footerSections ?? []
    }))
    .sort((left, right) => left.startNodeIndex - right.startNodeIndex);

  if (normalizedSections.length > 0) {
    if (normalizedSections[0].startNodeIndex > 0) {
      normalizedSections.unshift({
        startNodeIndex: 0,
        sectionPropertiesXml: normalizedSections[0].sectionPropertiesXml,
        headerSections: normalizedSections[0].headerSections,
        footerSections: normalizedSections[0].footerSections
      });
    }
    return resolveInheritedSectionHeaderFooterReferences(normalizedSections);
  }

  return [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: metadata.sectionPropertiesXml,
      headerSections: metadata.headerSections ?? [],
      footerSections: metadata.footerSections ?? []
    }
  ];
}

export function resolveSectionIndexForNodeIndex(
  sections: Pick<ResolvedModelSection, "startNodeIndex">[],
  nodeIndex: number,
  previousSectionIndex: number
): number {
  if (sections.length === 0) {
    return 0;
  }

  const safePrevious = Math.max(0, Math.min(previousSectionIndex, sections.length - 1));
  let sectionIndex = safePrevious;

  if (nodeIndex < sections[sectionIndex].startNodeIndex) {
    sectionIndex = 0;
  }

  while (
    sectionIndex + 1 < sections.length &&
    sections[sectionIndex + 1].startNodeIndex <= nodeIndex
  ) {
    sectionIndex += 1;
  }

  return sectionIndex;
}
