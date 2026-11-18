// Attach drawing objects and equations to the WASM-built document model.
//
// The WASM model builder only understands text, pictures, form fields and text
// boxes (the last as a synthetic image). Everything else a Word document can
// carry — autoshapes, WordArt, shape groups, SmartArt, charts, embedded OLE
// objects, 3D models and equations — is left out of the model, so an edit that
// rebuilds a paragraph silently drops it.
//
// This pass reads the objects back from each paragraph's own markup and attaches
// them as `DrawingRunNode` / `MathRunNode` children. The children carry the
// original markup, so the serializer can write the object back on export and the
// renderer can draw the content it knows (text boxes, chart data, poster
// frames) or a labelled placeholder.

import type { OoxmlPackage } from "./ooxml-core";
import type {
  ChartData,
  DocModel,
  DocNode,
  DrawingAnchor,
  DrawingObjectKind,
  DrawingRunNode,
  ImageRunNode,
  MathRunNode,
  ParagraphChildNode,
  ParagraphNode,
  TableNode,
  TextBoxContent,
} from "./types";
import {
  drawingExtentPx,
  drawingObjectIdentity,
  drawingRelationshipId,
  mathIsDisplay,
  mathTextFromXml,
  model3dPosterRelationshipId,
  oleObjectInfo,
  paragraphDrawingObjects,
  parseChartData,
  parseDrawingAnchor,
  parseDrawingShapeStyle,
  parseTextBoxContent,
} from "./drawing-objects";

const DOCUMENT_PART = "word/document.xml";

export interface DrawingObjectEnrichmentContext {
  /** Package parts, used to resolve chart parts and 3D poster images. */
  parts: ReadonlyMap<string, { content: string }>;
  binaryAssets: ReadonlyMap<string, Uint8Array>;
  /** Parsed `r:id` → part maps per story part, shared by every paragraph. */
  relationshipCache?: Map<string, Map<string, string>>;
}

// ── Package relationship helpers ────────────────────────────────────

function partDirectory(partName: string): string {
  const index = partName.lastIndexOf("/");
  return index >= 0 ? partName.slice(0, index) : "";
}

function partFileName(partName: string): string {
  const index = partName.lastIndexOf("/");
  return index >= 0 ? partName.slice(index + 1) : partName;
}

/** `word/_rels/document.xml.rels` for `word/document.xml`. */
export function relationshipPartName(partName: string): string {
  const directory = partDirectory(partName);
  const rels = `${directory ? `${directory}/` : ""}_rels/${partFileName(partName)}.rels`;
  return rels;
}

function resolveRelationshipTarget(partName: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const directory = partDirectory(partName);
  const segments = `${directory ? `${directory}/` : ""}${target}`.split("/");
  const resolved: string[] = [];
  for (const segment of segments) {
    if (segment === "." || segment === "") continue;
    if (segment === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }
  return resolved.join("/");
}

/** Maps `r:id` → package part name for one story part (cached per context). */
export function storyRelationshipTargets(
  context: DrawingObjectEnrichmentContext,
  partName: string,
): Map<string, string> {
  const cached = context.relationshipCache?.get(partName);
  if (cached) return cached;
  const targets = new Map<string, string>();
  context.relationshipCache?.set(partName, targets);
  const relsXml = context.parts.get(relationshipPartName(partName))?.content;
  if (!relsXml) return targets;
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const tag = match[0];
    const id = /\bId="([^"]+)"/.exec(tag)?.[1];
    const target = /\bTarget="([^"]+)"/.exec(tag)?.[1];
    if (!id || !target) continue;
    targets.set(id, resolveRelationshipTarget(partName, target));
  }
  return targets;
}

const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  tif: "image/tiff",
  tiff: "image/tiff",
  emf: "image/emf",
  wmf: "image/wmf",
};

const BASE64_ALPHABET =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

function base64Encode(bytes: Uint8Array): string {
  let output = "";
  for (let index = 0; index < bytes.length; index += 3) {
    const first = bytes[index] ?? 0;
    const second = bytes[index + 1];
    const third = bytes[index + 2];
    const triple = (first << 16) | ((second ?? 0) << 8) | (third ?? 0);
    output += BASE64_ALPHABET[(triple >> 18) & 63];
    output += BASE64_ALPHABET[(triple >> 12) & 63];
    output += second === undefined ? "=" : BASE64_ALPHABET[(triple >> 6) & 63];
    output += third === undefined ? "=" : BASE64_ALPHABET[triple & 63];
  }
  return output;
}

/** Data URI of a package media part, when the bytes are available. */
export function mediaPartDataUri(
  context: DrawingObjectEnrichmentContext,
  partName: string | undefined,
): string | undefined {
  if (!partName) return undefined;
  const bytes = context.binaryAssets.get(partName);
  if (!bytes || bytes.length === 0) return undefined;
  const extension = partName.slice(partName.lastIndexOf(".") + 1).toLowerCase();
  const contentType = IMAGE_CONTENT_TYPES[extension];
  if (!contentType) return undefined;
  if (contentType === "image/svg+xml") {
    return `data:${contentType};charset=utf-8,${encodeURIComponent(
      new TextDecoder().decode(bytes),
    )}`;
  }
  return `data:${contentType};base64,${base64Encode(bytes)}`;
}

// ── Object → model node ─────────────────────────────────────────────

function normalizeXml(xml: string): string {
  return xml.replace(/\s+/g, " ").trim();
}

function childRepresentsXml(child: ParagraphChildNode, sourceXml: string): boolean {
  if (child.type === "image" || child.type === "drawing" || child.type === "math") {
    return Boolean(child.sourceXml) && normalizeXml(child.sourceXml ?? "") === normalizeXml(sourceXml);
  }
  return false;
}

function chartDataForObject(
  sourceXml: string,
  kind: DrawingObjectKind,
  context: DrawingObjectEnrichmentContext,
  partName: string,
  chartCache: Map<string, ChartData | undefined>,
): ChartData | undefined {
  if (kind !== "chart") return undefined;
  const relationshipId = drawingRelationshipId(sourceXml, "c:chart");
  if (!relationshipId) return undefined;
  const target = storyRelationshipTargets(context, partName).get(relationshipId);
  if (!target) return undefined;
  if (chartCache.has(target)) return chartCache.get(target);
  const chartXml = context.parts.get(target)?.content;
  const data = parseChartData(chartXml);
  chartCache.set(target, data);
  return data;
}

function posterDataUriForObject(
  sourceXml: string,
  kind: DrawingObjectKind,
  context: DrawingObjectEnrichmentContext,
  partName: string,
): string | undefined {
  if (kind !== "model3d") return undefined;
  const relationshipId = model3dPosterRelationshipId(sourceXml);
  if (!relationshipId) return undefined;
  const target = storyRelationshipTargets(context, partName).get(relationshipId);
  return mediaPartDataUri(context, target);
}

function drawingNodeFromXml(
  sourceXml: string,
  kind: DrawingObjectKind,
  context: DrawingObjectEnrichmentContext,
  partName: string,
  chartCache: Map<string, ChartData | undefined>,
): DrawingRunNode {
  const extent = drawingExtentPx(sourceXml);
  const floating = parseDrawingAnchor(sourceXml);
  const identity = drawingObjectIdentity(sourceXml);
  const chart = chartDataForObject(sourceXml, kind, context, partName, chartCache);
  const poster = posterDataUriForObject(sourceXml, kind, context, partName);
  const shape =
    kind === "shape" || kind === "wordArt" || kind === "textBox"
      ? parseDrawingShapeStyle(sourceXml)
      : undefined;
  const ole = kind === "oleObject" ? oleObjectInfo(sourceXml) : undefined;
  const textBox = kind === "textBox" ? parseTextBoxContent(sourceXml) : undefined;

  return {
    type: "drawing",
    kind,
    ...identity,
    ...extent,
    ...(floating ? { floating } : {}),
    ...(floating?.zIndex !== undefined ? { zIndex: floating.zIndex } : {}),
    ...(poster ? { src: poster } : {}),
    ...(textBox ? { textBox } : {}),
    ...(chart ? { chart } : {}),
    ...(shape ? { shape } : {}),
    ...(ole ? { ole } : {}),
    sourceXml,
  };
}

function mathNodeFromXml(sourceXml: string): MathRunNode {
  return {
    type: "math",
    text: mathTextFromXml(sourceXml),
    ...(mathIsDisplay(sourceXml) ? { display: true } : {}),
    sourceXml,
  };
}

/** Model text length before `childIndex`, ignoring non-text children. */
function insertionIndexForTextOffset(
  children: readonly ParagraphChildNode[],
  textOffset: number,
): number {
  let consumed = 0;
  let insertionIndex = children.length;
  for (let index = 0; index < children.length; index += 1) {
    const child = children[index];
    if (!child || child.type !== "text") continue;
    const next = consumed + child.text.length;
    if (textOffset <= consumed) {
      insertionIndex = index;
      break;
    }
    if (textOffset < next) {
      insertionIndex = index + 1;
      break;
    }
    consumed = next;
  }
  return insertionIndex;
}

function withTextBoxContent(image: ImageRunNode): ImageRunNode {
  if (!image.sourceXml) return image;
  if (!image.syntheticTextBox && !/<w:txbxContent\b/i.test(image.sourceXml) && !/<v:textbox\b/i.test(image.sourceXml)) {
    return image;
  }
  const textBox = image.textBox ?? parseTextBoxContent(image.sourceXml);
  const textBoxSourceXml = image.textBoxSourceXml ?? image.sourceXml;
  if (textBox === image.textBox && textBoxSourceXml === image.textBoxSourceXml) {
    return image;
  }
  return {
    ...image,
    ...(textBox ? { textBox } : {}),
    textBoxSourceXml,
  };
}

interface ParagraphEnrichment {
  paragraph: ParagraphNode;
  context: DrawingObjectEnrichmentContext;
  partName: string;
  chartCache: Map<string, ChartData | undefined>;
}

function enrichParagraph(settings: ParagraphEnrichment): ParagraphNode {
  const { paragraph, context, partName, chartCache } = settings;
  let children = paragraph.children;
  let changed = false;

  for (let index = 0; index < children.length; index += 1) {
    const child = children[index];
    if (!child || child.type !== "image") continue;
    const enriched = withTextBoxContent(child);
    if (enriched === child) continue;
    if (!changed) {
      children = [...children];
      changed = true;
    }
    children[index] = enriched;
  }

  if (!paragraph.sourceXml) {
    return changed ? { ...paragraph, children } : paragraph;
  }

  const objects = paragraphDrawingObjects(paragraph.sourceXml);
  if (objects.length === 0) {
    return changed ? { ...paragraph, children } : paragraph;
  }

  // Insert from the last object to the first so several objects sharing one
  // text offset keep their document order.
  for (let index = objects.length - 1; index >= 0; index -= 1) {
    const object = objects[index];
    if (!object || object.kind === "picture") continue;
    const existingImageIndex = children.findIndex((child) =>
      childRepresentsXml(child, object.sourceXml),
    );
    if (existingImageIndex >= 0) {
      const existing = children[existingImageIndex];
      // The WASM builder also emits an image run for drawings it rendered as a
      // flat SVG (a shape, a chart, a 3D group). Those images keep no model
      // node for the object, so an edit that rebuilds the paragraph drops it.
      // Replace the image with a drawing node that keeps the rendered preview
      // (`src`) and the stored markup. A text box — or a group that carries text
      // box content — stays an image run: its synthetic SVG already models the
      // text and the pagination heuristics know those images.
      if (
        existing?.type === "image" &&
        object.kind !== "math" &&
        object.kind !== "textBox" &&
        existing.syntheticTextBox !== true
      ) {
        const drawingNode = drawingNodeFromXml(
          object.sourceXml,
          object.kind,
          context,
          partName,
          chartCache,
        );
        const withImageLayout: DrawingRunNode = {
          ...drawingNode,
          // The rendered SVG preview is the best picture for a shape or group; a
          // 3D model keeps its poster frame, which is a real image.
          ...(existing.src && object.kind !== "model3d"
            ? { src: existing.src }
            : {}),
          ...(existing.widthPx !== undefined ? { widthPx: existing.widthPx } : {}),
          ...(existing.heightPx !== undefined ? { heightPx: existing.heightPx } : {}),
          ...(existing.alt ? { alt: existing.alt } : {}),
          ...(existing.floating ? { floating: existing.floating } : {}),
        };
        if (!changed) {
          children = [...children];
          changed = true;
        }
        children[existingImageIndex] = withImageLayout;
      }
      continue;
    }

    const node: ParagraphChildNode =
      object.kind === "math"
        ? mathNodeFromXml(object.sourceXml)
        : drawingNodeFromXml(
            object.sourceXml,
            object.kind,
            context,
            partName,
            chartCache,
          );
    const insertionIndex = insertionIndexForTextOffset(children, object.textOffset);
    if (!changed) {
      children = [...children];
      changed = true;
    }
    children.splice(insertionIndex, 0, node);
  }

  return changed ? { ...paragraph, children } : paragraph;
}

function enrichNodes(
  nodes: readonly DocNode[],
  context: DrawingObjectEnrichmentContext,
  partName: string,
  chartCache: Map<string, ChartData | undefined>,
): DocNode[] {
  let changed = false;
  const next = nodes.map((node) => {
    if (node.type === "table") {
      const table = enrichTable(node, context, partName, chartCache);
      if (table !== node) changed = true;
      return table;
    }
    const paragraph = enrichParagraph({ paragraph: node, context, partName, chartCache });
    if (paragraph !== node) changed = true;
    return paragraph;
  });
  return changed ? next : (nodes as DocNode[]);
}

function enrichTable(
  table: TableNode,
  context: DrawingObjectEnrichmentContext,
  partName: string,
  chartCache: Map<string, ChartData | undefined>,
): TableNode {
  let changed = false;
  const rows = table.rows.map((row) => {
    const cells = row.cells.map((cell) => {
      const nodes = enrichNodes(cell.nodes, context, partName, chartCache);
      return nodes === cell.nodes ? cell : { ...cell, nodes };
    });
    if (cells.some((cell, index) => cell !== row.cells[index])) {
      changed = true;
      return { ...row, cells };
    }
    return row;
  });
  return changed ? { ...table, rows } : table;
}

function enrichStory(
  context: DrawingObjectEnrichmentContext,
  nodes: DocNode[],
  partName: string,
  chartCache: Map<string, ChartData | undefined>,
): DocNode[] {
  return enrichNodes(nodes, context, partName, chartCache);
}

/**
 * Adds the missing drawing objects and equations to a freshly built model.
 *
 * `partName` selects the relationship part used to resolve chart parts and 3D
 * poster images: `word/document.xml` for the body and the header/footer part
 * name for section stories.
 */
export function enrichDocModelDrawingObjects(
  model: DocModel,
  pkg: OoxmlPackage,
): DocModel {
  const context: DrawingObjectEnrichmentContext = {
    parts: pkg.parts,
    binaryAssets: pkg.binaryAssets,
    relationshipCache: new Map<string, Map<string, string>>(),
  };
  const chartCache = new Map<string, ChartData | undefined>();
  const nodes = enrichStory(context, model.nodes, DOCUMENT_PART, chartCache);

  const metadata = model.metadata;
  const headerSections = metadata.headerSections?.map((section) => {
    const sectionNodes = enrichStory(
      context,
      section.nodes,
      section.partName,
      chartCache,
    );
    return sectionNodes === section.nodes ? section : { ...section, nodes: sectionNodes };
  });
  const footerSections = metadata.footerSections?.map((section) => {
    const sectionNodes = enrichStory(
      context,
      section.nodes,
      section.partName,
      chartCache,
    );
    return sectionNodes === section.nodes ? section : { ...section, nodes: sectionNodes };
  });
  const sections = metadata.sections?.map((section) => {
    const sectionHeaders = section.headerSections.map((header) => {
      const sectionNodes = enrichStory(
        context,
        header.nodes,
        header.partName,
        chartCache,
      );
      return sectionNodes === header.nodes ? header : { ...header, nodes: sectionNodes };
    });
    const sectionFooters = section.footerSections.map((footer) => {
      const sectionNodes = enrichStory(
        context,
        footer.nodes,
        footer.partName,
        chartCache,
      );
      return sectionNodes === footer.nodes ? footer : { ...footer, nodes: sectionNodes };
    });
    const headerChanged = sectionHeaders.some(
      (header, index) => header !== section.headerSections[index],
    );
    const footerChanged = sectionFooters.some(
      (footer, index) => footer !== section.footerSections[index],
    );
    return headerChanged || footerChanged
      ? { ...section, headerSections: sectionHeaders, footerSections: sectionFooters }
      : section;
  });

  const noteStories = <T extends { nodes?: DocNode[] }>(
    notes: T[] | undefined,
    partName: string,
  ): T[] | undefined =>
    notes?.map((note) => {
      if (!note.nodes) return note;
      const noteNodes = enrichStory(context, note.nodes, partName, chartCache);
      return noteNodes === note.nodes ? note : { ...note, nodes: noteNodes };
    });

  const footnotes = noteStories(metadata.footnotes, "word/footnotes.xml");
  const endnotes = noteStories(metadata.endnotes, "word/endnotes.xml");

  const headerChanged = headerSections?.some(
    (section, index) => section !== metadata.headerSections?.[index],
  );
  const footerChanged = footerSections?.some(
    (section, index) => section !== metadata.footerSections?.[index],
  );
  const sectionsChanged = sections?.some(
    (section, index) => section !== metadata.sections?.[index],
  );

  if (
    nodes === model.nodes &&
    !headerChanged &&
    !footerChanged &&
    !sectionsChanged &&
    footnotes === metadata.footnotes &&
    endnotes === metadata.endnotes
  ) {
    return model;
  }

  return {
    ...model,
    nodes,
    metadata: {
      ...metadata,
      ...(headerSections ? { headerSections } : {}),
      ...(footerSections ? { footerSections } : {}),
      ...(sections ? { sections } : {}),
      ...(footnotes ? { footnotes } : {}),
      ...(endnotes ? { endnotes } : {}),
    },
  };
}

/** Exported for tests: the objects one paragraph markup carries. */
export { paragraphDrawingObjects, parseTextBoxContent };

