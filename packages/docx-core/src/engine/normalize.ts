import type {
  DocModel,
  DocNode,
  ImageRunNode,
  ParagraphChildNode,
  ParagraphNode,
  TableCellContentNode
} from "./types";
import { parseFloatingAnchorProperties } from "./anchor-properties";
import { parseParagraphCommentRanges } from "./comment-ranges";
import { sectionBreakFromSourceXml } from "./section-break";

function normalizeUint8Array(value: unknown): Uint8Array | undefined {
  if (value instanceof Uint8Array) {
    return value;
  }
  if (Array.isArray(value)) {
    return Uint8Array.from(value as number[]);
  }
  return undefined;
}

function normalizeParagraphChild(child: ParagraphChildNode): ParagraphChildNode {
  if (child.type !== "image") {
    return child;
  }

  const image = child as ImageRunNode;
  const data = normalizeUint8Array(image.data);
  const anchorProperties = image.floating
    ? parseFloatingAnchorProperties(image.sourceXml)
    : undefined;
  const floating = anchorProperties
    ? { ...image.floating, ...anchorProperties }
    : image.floating;
  const dataChanged = data !== image.data;
  const floatingChanged = floating !== image.floating;
  if (!dataChanged && !floatingChanged) {
    return child;
  }

  return {
    ...image,
    data,
    floating
  };
}

function normalizeTableCellContent(node: TableCellContentNode): TableCellContentNode {
  if (node.type === "table") {
    return normalizeDocNode(node) as TableCellContentNode;
  }

  return normalizeDocNode(node) as ParagraphNode;
}

function paragraphModelTextLength(node: ParagraphNode): number {
  let length = 0;
  for (const child of node.children) {
    if (child.type === "text") {
      length += child.text.length;
    }
  }
  return length;
}

/**
 * Captures `w:commentRangeStart` / `w:commentRangeEnd` as character offsets the
 * first time a parsed paragraph is normalized. The offsets live on the model so
 * highlights and comment cards survive the edits that clear `sourceXml`.
 */
function normalizeParagraphCommentRanges(
  node: ParagraphNode
): ParagraphNode["commentRanges"] {
  if (node.commentRanges) {
    return node.commentRanges;
  }
  if (!node.sourceXml || !/commentRange|commentReference/i.test(node.sourceXml)) {
    return undefined;
  }

  const ranges = parseParagraphCommentRanges(
    node.sourceXml,
    paragraphModelTextLength(node)
  );
  return ranges.length > 0 ? ranges : undefined;
}

function normalizeDocNode(node: DocNode): DocNode {
  if (node.type === "paragraph") {
    return {
      ...node,
      sectionBreak:
        node.sectionBreak ?? sectionBreakFromSourceXml(node.sourceXml),
      commentRanges: normalizeParagraphCommentRanges(node),
      children: node.children.map(normalizeParagraphChild)
    };
  }

  return {
    ...node,
    rows: node.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        nodes: cell.nodes.map(normalizeTableCellContent)
      }))
    }))
  };
}

export function normalizeDocModel(model: DocModel): DocModel {
  const nodes = model.nodes.map(normalizeDocNode);
  const headerSections = model.metadata.headerSections.map((section) => ({
    ...section,
    nodes: section.nodes.map(normalizeDocNode)
  }));
  const footerSections = model.metadata.footerSections.map((section) => ({
    ...section,
    nodes: section.nodes.map(normalizeDocNode)
  }));
  const sections = normalizeDocumentSections(
    nodes,
    model.metadata,
    headerSections,
    footerSections
  );

  return {
    ...model,
    nodes,
    metadata: {
      ...model.metadata,
      headerSections,
      footerSections,
      sections,
      footnotes: model.metadata.footnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map(normalizeDocNode)
      })),
      endnotes: model.metadata.endnotes?.map((note) => ({
        ...note,
        nodes: note.nodes?.map(normalizeDocNode)
      }))
    }
  };
}

/**
 * Rebuild metadata.sections from preserved paragraph markers when the importer
 * (or an older snapshot) did not provide section metadata. Existing imported
 * section entries stay authoritative.
 */
function normalizeDocumentSections(
  nodes: DocNode[],
  metadata: DocModel["metadata"],
  headerSections: DocModel["metadata"]["headerSections"],
  footerSections: DocModel["metadata"]["footerSections"]
): DocModel["metadata"]["sections"] {
  const normalizedSections = metadata.sections?.map((section) => ({
    ...section,
    headerSections: section.headerSections.map((header) => ({
      ...header,
      nodes: header.nodes.map(normalizeDocNode)
    })),
    footerSections: section.footerSections.map((footer) => ({
      ...footer,
      nodes: footer.nodes.map(normalizeDocNode)
    }))
  }));

  if (normalizedSections && normalizedSections.length > 0) {
    return normalizedSections;
  }

  const markers: Array<{ startNodeIndex: number; sectionPropertiesXml: string }> = [];
  nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph" && node.sectionBreak) {
      markers.push({
        startNodeIndex: nodeIndex + 1,
        sectionPropertiesXml: node.sectionBreak.sectionPropertiesXml
      });
    }
  });

  if (markers.length === 0) {
    return normalizedSections;
  }

  const sections: NonNullable<DocModel["metadata"]["sections"]> = [];
  let startNodeIndex = 0;
  for (const marker of markers) {
    sections.push({
      startNodeIndex,
      sectionPropertiesXml: marker.sectionPropertiesXml,
      headerSections,
      footerSections
    });
    startNodeIndex = marker.startNodeIndex;
  }
  sections.push({
    startNodeIndex,
    sectionPropertiesXml: metadata.sectionPropertiesXml,
    headerSections,
    footerSections
  });

  return sections;
}
