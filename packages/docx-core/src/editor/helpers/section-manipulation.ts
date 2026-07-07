// Section (header/footer) paragraph and image location keying, lookup, and
// mutation. Lets the editor surface edit text and floating-image state inside
// header/footer section node lists using the same DocxSectionParagraphLocation /
// DocxSectionImageLocation coordinates used elsewhere.
// Upstream editor.tsx: lines 24641-24952.

import type {
  DocModel,
  ImageRunNode,
  ParagraphNode,
  TableNode
} from "../../engine/types";
import { cloneDocModel } from "../../engine/clone";
import {
  updateParagraphText,
  updateTableCellParagraphText
} from "../editor-ops";
import type {
  DocxSectionImageLocation,
  DocxSectionParagraphLocation,
  ParagraphLocation
} from "./editor-types";
import {
  paragraphText,
  tableCellParagraphs
} from "./paragraph-inspect";

// `ParagraphTextUpdateOptions` mirrors the upstream alias
// `type ParagraphTextUpdateOptions = Parameters<typeof updateParagraphText>[3]`
// (editor.tsx:24364). We re-derive it from the local editor-ops signature so
// callers stay in sync with the canonical option shape.
export type ParagraphTextUpdateOptions = Parameters<
  typeof updateParagraphText
>[3];

export function sectionParagraphLocationKey(
  location: DocxSectionParagraphLocation
): string {
  return JSON.stringify({
    region: location.region,
    partName: location.partName,
    nodeIndex: location.nodeIndex,
    rowIndex: location.rowIndex ?? -1,
    cellIndex: location.cellIndex ?? -1,
    paragraphIndex: location.paragraphIndex ?? -1,
  });
}

export function sectionImageLocationKey(
  location: DocxSectionImageLocation
): string {
  return JSON.stringify({
    region: location.region,
    partName: location.partName,
    nodeIndex: location.nodeIndex,
    rowIndex: location.rowIndex ?? -1,
    cellIndex: location.cellIndex ?? -1,
    paragraphIndex: location.paragraphIndex ?? -1,
    childIndex: location.childIndex,
  });
}

export function tableCellParagraphDraftKey(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number
): string {
  return `${tableIndex}:${rowIndex}:${cellIndex}:${paragraphIndex}`;
}

export function parseSectionParagraphLocationKey(
  raw: string
): DocxSectionParagraphLocation | undefined {
  try {
    const parsed = JSON.parse(raw) as {
      region?: string;
      partName?: string;
      nodeIndex?: number;
      rowIndex?: number;
      cellIndex?: number;
      paragraphIndex?: number;
    };

    if (
      !parsed ||
      (parsed.region !== "header" && parsed.region !== "footer") ||
      typeof parsed.partName !== "string" ||
      !Number.isFinite(parsed.nodeIndex)
    ) {
      return undefined;
    }

    const rowIndex =
      Number.isFinite(parsed.rowIndex) && (parsed.rowIndex as number) >= 0
        ? Math.round(parsed.rowIndex as number)
        : undefined;
    const cellIndex =
      Number.isFinite(parsed.cellIndex) && (parsed.cellIndex as number) >= 0
        ? Math.round(parsed.cellIndex as number)
        : undefined;
    const paragraphIndex =
      Number.isFinite(parsed.paragraphIndex) &&
      (parsed.paragraphIndex as number) >= 0
        ? Math.round(parsed.paragraphIndex as number)
        : undefined;

    return {
      region: parsed.region,
      partName: parsed.partName,
      nodeIndex: Math.max(0, Math.round(parsed.nodeIndex as number)),
      rowIndex,
      cellIndex,
      paragraphIndex,
    };
  } catch {
    return undefined;
  }
}

export function paragraphTextFromSectionLocation(
  sectionNodes: DocModel["nodes"],
  location: Omit<DocxSectionParagraphLocation, "region" | "partName">
): string | undefined {
  const rootNode = sectionNodes[location.nodeIndex];
  if (!rootNode) {
    return undefined;
  }

  if (location.rowIndex === undefined || location.cellIndex === undefined) {
    if (rootNode.type !== "paragraph") {
      return undefined;
    }
    return paragraphText(rootNode);
  }

  if (rootNode.type !== "table") {
    return undefined;
  }

  const paragraphIndex = Math.max(0, Math.round(location.paragraphIndex ?? 0));
  const cell = rootNode.rows[location.rowIndex]?.cells[location.cellIndex];
  if (!cell) {
    return undefined;
  }

  const paragraph = tableCellParagraphs(cell.nodes)[paragraphIndex];
  return paragraph ? paragraphText(paragraph) : undefined;
}

export function sectionParagraphFromLocation(
  sectionNodes: DocModel["nodes"],
  location: Omit<DocxSectionParagraphLocation, "region" | "partName">
): {
  paragraph?: ParagraphNode;
  tableNode?: TableNode;
} {
  const rootNode = sectionNodes[location.nodeIndex];
  if (!rootNode) {
    return {};
  }

  if (location.rowIndex === undefined || location.cellIndex === undefined) {
    if (rootNode.type !== "paragraph") {
      return {};
    }
    return { paragraph: rootNode };
  }

  if (rootNode.type !== "table") {
    return {};
  }

  const paragraphIndex = Math.max(0, Math.round(location.paragraphIndex ?? 0));
  const cell = rootNode.rows[location.rowIndex]?.cells[location.cellIndex];
  if (!cell) {
    return {};
  }

  const paragraph = tableCellParagraphs(cell.nodes)[paragraphIndex];
  if (!paragraph) {
    return {};
  }

  return {
    paragraph,
    tableNode: rootNode,
  };
}

export function updateSectionParagraphTextAtLocation(
  model: DocModel,
  location: DocxSectionParagraphLocation,
  text: string,
  options?: ParagraphTextUpdateOptions
): DocModel {
  const next = cloneDocModel(model);
  const candidateSectionLists =
    location.region === "header"
      ? [
          next.metadata.headerSections ?? [],
          ...(next.metadata.sections?.map(
            (section) => section.headerSections ?? []
          ) ?? []),
        ]
      : [
          next.metadata.footerSections ?? [],
          ...(next.metadata.sections?.map(
            (section) => section.footerSections ?? []
          ) ?? []),
        ];

  const normalizedLocation = {
    nodeIndex: Math.max(0, Math.round(location.nodeIndex)),
    rowIndex:
      Number.isFinite(location.rowIndex) && (location.rowIndex as number) >= 0
        ? Math.round(location.rowIndex as number)
        : undefined,
    cellIndex:
      Number.isFinite(location.cellIndex) && (location.cellIndex as number) >= 0
        ? Math.round(location.cellIndex as number)
        : undefined,
    paragraphIndex:
      Number.isFinite(location.paragraphIndex) &&
      (location.paragraphIndex as number) >= 0
        ? Math.round(location.paragraphIndex as number)
        : undefined,
  };

  let changed = false;
  candidateSectionLists.forEach((sections) => {
    sections.forEach((section) => {
      if (section.partName !== location.partName) {
        return;
      }

      const currentText = paragraphTextFromSectionLocation(
        section.nodes,
        normalizedLocation
      );
      if (currentText === undefined || currentText === text) {
        return;
      }

      const scopedModel: DocModel = {
        ...next,
        nodes: section.nodes,
      };
      const updatedScopedModel =
        normalizedLocation.rowIndex === undefined ||
        normalizedLocation.cellIndex === undefined
          ? updateParagraphText(
              scopedModel,
              normalizedLocation.nodeIndex,
              text,
              options
            )
          : updateTableCellParagraphText(
              scopedModel,
              normalizedLocation.nodeIndex,
              normalizedLocation.rowIndex,
              normalizedLocation.cellIndex,
              normalizedLocation.paragraphIndex ?? 0,
              text,
              options
            );

      section.nodes = updatedScopedModel.nodes;
      changed = true;
    });
  });

  return changed ? next : model;
}

export function updateSectionImageFloatingAtLocation(
  model: DocModel,
  location: DocxSectionImageLocation,
  patch: Partial<NonNullable<ImageRunNode["floating"]>>
): DocModel {
  const next = cloneDocModel(model);
  const candidateSectionLists =
    location.region === "header"
      ? [
          next.metadata.headerSections ?? [],
          ...(next.metadata.sections?.map(
            (section) => section.headerSections ?? []
          ) ?? []),
        ]
      : [
          next.metadata.footerSections ?? [],
          ...(next.metadata.sections?.map(
            (section) => section.footerSections ?? []
          ) ?? []),
        ];

  const normalizedLocation = {
    nodeIndex: Math.max(0, Math.round(location.nodeIndex)),
    rowIndex:
      Number.isFinite(location.rowIndex) && (location.rowIndex as number) >= 0
        ? Math.round(location.rowIndex as number)
        : undefined,
    cellIndex:
      Number.isFinite(location.cellIndex) && (location.cellIndex as number) >= 0
        ? Math.round(location.cellIndex as number)
        : undefined,
    paragraphIndex:
      Number.isFinite(location.paragraphIndex) &&
      (location.paragraphIndex as number) >= 0
        ? Math.round(location.paragraphIndex as number)
        : undefined,
    childIndex: Math.max(0, Math.round(location.childIndex)),
  };

  let changed = false;
  candidateSectionLists.forEach((sections) => {
    sections.forEach((section) => {
      if (section.partName !== location.partName) {
        return;
      }

      const lookup = sectionParagraphFromLocation(
        section.nodes,
        normalizedLocation
      );
      const paragraph = lookup.paragraph;
      if (!paragraph) {
        return;
      }

      const child = paragraph.children[normalizedLocation.childIndex];
      if (!child || child.type !== "image") {
        return;
      }

      child.floating = {
        ...(child.floating ?? {}),
        ...patch,
      };
      paragraph.sourceXml = undefined;
      if (lookup.tableNode) {
        lookup.tableNode.sourceXml = undefined;
      }
      changed = true;
    });
  });

  return changed ? next : model;
}

export function updateParagraphTextAtLocation(
  model: DocModel,
  location: ParagraphLocation,
  text: string,
  options?: ParagraphTextUpdateOptions
): DocModel {
  if (location.kind === "paragraph") {
    return updateParagraphText(model, location.nodeIndex, text, options);
  }

  return updateTableCellParagraphText(
    model,
    location.tableIndex,
    location.rowIndex,
    location.cellIndex,
    location.paragraphIndex,
    text,
    options
  );
}
