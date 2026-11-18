import { strToU8 } from "fflate";
import type {
  XlsxChart,
  XlsxDrawingClipboardItem,
  XlsxDrawingClipboardKind,
  XlsxImage,
  XlsxImageAnchor,
  XlsxShape,
} from "../types";
import type { WorkbookChartAssets, WorkbookChartOrigin } from "../charts/chart-types";
import { CHART_EX_REL_TYPE, CHART_REL_TYPE } from "../charts/chart-series";
import { updateAnchorNode } from "./image-export";
import { appendWorkbookSheetDrawingAttachment, ensureContentTypeOverride, insertSheetChild } from "../authoring/sheet-authoring";
import type { WorkbookImageAssets, WorkbookSheetState } from "./grid-render";
import {
  CONTENT_TYPES_NS,
  DRAWING_CONTENT_TYPE,
  DRAWING_REL_TYPE,
  PKG_REL_NS,
  REL_NS,
  SPREADSHEET_NS,
  type ArchiveEntries,
  type WorkbookImageOrigin,
  type WorkbookShapeOrigin,
  cloneBytes,
  createImageAnchorNodes,
  dirname,
  getFirstDescendant,
  getLocalElements,
  normalizeArchivePath,
  parseXml,
  readArchiveText,
  relativeArchivePath,
  relsPathForDocument,
  resolveArchiveTarget,
  serializeXml,
} from "./image-parser";
import { anchorToAbsoluteRect } from "./column-width";

const SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

/** What a caller has to identify to clone one drawing. */
export type WorkbookDrawingCloneSource = {
  /** Stable model identity (`image`, `shape`, or `chart` id). */
  id: string;
  kind: XlsxDrawingClipboardKind;
  /** Workbook sheet index the drawing belongs to. */
  workbookSheetIndex: number;
};

export type WorkbookDrawingCloneOptions = {
  /** Force the new model id; a unique id is generated when omitted. */
  id?: string;
  /** Visible sheet index used by viewer model arrays; defaults to the workbook index. */
  visibleSheetIndex?: number;
};

export type WorkbookDrawingCloneResult = {
  id: string;
  kind: XlsxDrawingClipboardKind;
  model: XlsxChart | XlsxImage | XlsxShape;
};

type DrawingModel = XlsxChart | XlsxImage | XlsxShape;

type SourceDrawing = {
  anchorNode: Element;
  drawingDocument: XMLDocument;
  drawingPath: string;
  model: DrawingModel;
  origin: WorkbookImageOrigin | WorkbookChartOrigin;
};

type TargetDrawing = {
  drawingDocument: XMLDocument;
  drawingPath: string;
  relationshipsDocument: XMLDocument;
  relationshipsPath: string;
};

type SourceRelationship = {
  id: string;
  target: string;
  targetMode: string | null;
  type: string;
};

function isElementNode(node: Node | null | undefined): node is Element {
  return Boolean(node && node.nodeType === 1);
}

function descendantsByLocalName(root: Element, localName: string): Element[] {
  const matches: Element[] = [];
  const visit = (node: Element) => {
    for (const child of Array.from(node.childNodes)) {
      if (!isElementNode(child)) continue;
      const element = child as Element;
      if (element.localName === localName) matches.push(element);
      visit(element);
    }
  };
  visit(root);
  return matches;
}

function relationshipIdsInSubtree(root: Element): Set<string> {
  const ids = new Set<string>();
  const visit = (node: Element) => {
    for (const attribute of Array.from(node.attributes)) {
      if (
        attribute.namespaceURI === REL_NS
        || attribute.name === "r:id"
        || attribute.name === "r:embed"
        || attribute.name === "r:link"
      ) {
        if (attribute.value && !attribute.value.startsWith("#")) ids.add(attribute.value);
      }
    }
    for (const child of Array.from(node.childNodes)) {
      if (isElementNode(child)) visit(child as Element);
    }
  };
  visit(root);
  return ids;
}

function replaceRelationshipId(root: Element, oldId: string, newId: string) {
  const visit = (node: Element) => {
    for (const attribute of Array.from(node.attributes)) {
      if (
        attribute.value === oldId
        && (
          attribute.namespaceURI === REL_NS
          || attribute.name === "r:id"
          || attribute.name === "r:embed"
          || attribute.name === "r:link"
        )
      ) {
        attribute.value = newId;
      }
    }
    for (const child of Array.from(node.childNodes)) {
      if (isElementNode(child)) visit(child as Element);
    }
  };
  visit(root);
}

function findNonVisualObject(root: Element, objectId: string, objectKind?: string): Element | null {
  for (const child of Array.from(root.childNodes)) {
    if (!isElementNode(child)) continue;
    const element = child as Element;
    if (
      (!objectKind || element.localName === objectKind)
      && ["pic", "sp", "cxnSp", "graphicFrame"].includes(element.localName)
      && getFirstDescendant(element, "cNvPr")?.getAttribute("id") === objectId
    ) {
      return element;
    }
    const nested = findNonVisualObject(element, objectId, objectKind);
    if (nested) return nested;
  }
  return null;
}

function findAnchorForOrigin(document: XMLDocument, origin: WorkbookImageOrigin) {
  const anchors = createImageAnchorNodes(document);
  if (origin.objectId) {
    const match = anchors.find((anchor) => Boolean(findNonVisualObject(anchor, origin.objectId!, origin.objectKind)));
    if (match) return match;
  }
  return anchors[origin.anchorIndex] ?? null;
}

function readRelationships(archive: ArchiveEntries, relsPath: string, basePath: string): Map<string, SourceRelationship> {
  const xml = readArchiveText(archive, relsPath);
  const document = xml ? parseXml(xml) : null;
  const relationships = new Map<string, SourceRelationship>();
  if (!document) return relationships;
  for (const node of getLocalElements(document, "Relationship")) {
    const id = node.getAttribute("Id");
    const target = node.getAttribute("Target");
    const type = node.getAttribute("Type");
    if (!id || !target || !type) continue;
    const targetMode = node.getAttribute("TargetMode");
    relationships.set(id, {
      id,
      target: targetMode === "External" ? target : resolveArchiveTarget(basePath, target),
      targetMode,
      type,
    });
  }
  return relationships;
}

function ensureRelationshipsDocument(archive: ArchiveEntries, relsPath: string): XMLDocument | null {
  const xml = readArchiveText(archive, relsPath);
  if (xml) {
    const document = parseXml(xml);
    if (document) return document;
  }
  const document = parseXml(`<Relationships xmlns="${PKG_REL_NS}"></Relationships>`);
  return document;
}

function nextRelationshipId(document: XMLDocument): string {
  const used = new Set(
    getLocalElements(document, "Relationship")
      .map((node) => node.getAttribute("Id"))
      .filter((value): value is string => Boolean(value)),
  );
  let index = 1;
  while (used.has(`rIdClipboard${index}`)) index += 1;
  return `rIdClipboard${index}`;
}

function nextIndexedPath(archive: ArchiveEntries, directory: string, baseName: string, extension: string) {
  const dir = directory ? `${normalizeArchivePath(directory)}/` : "";
  let index = 1;
  while (archive[`${dir}${baseName}${index}.${extension}`]) index += 1;
  return `${dir}${baseName}${index}.${extension}`;
}

function copyContentTypeOverride(archive: ArchiveEntries, sourcePath: string, targetPath: string) {
  const xml = readArchiveText(archive, "[Content_Types].xml");
  const document = xml ? parseXml(xml) : null;
  if (!document) return;
  const sourceOverride = getLocalElements(document, "Override").find(
    (node) => normalizeArchivePath(node.getAttribute("PartName") ?? "") === normalizeArchivePath(sourcePath),
  );
  if (!sourceOverride) return;
  const targetPartName = `/${normalizeArchivePath(targetPath)}`;
  const existing = getLocalElements(document, "Override").find(
    (node) => node.getAttribute("PartName") === targetPartName,
  );
  if (existing) return;
  const override = document.createElementNS(CONTENT_TYPES_NS, "Override");
  override.setAttribute("PartName", targetPartName);
  override.setAttribute("ContentType", sourceOverride.getAttribute("ContentType") ?? "");
  document.documentElement?.appendChild(override);
  archive["[Content_Types].xml"] = strToU8(serializeXml(document));
}

function cloneChartPart(archive: ArchiveEntries, sourceChartPath: string): string | null {
  const normalizedSourcePath = normalizeArchivePath(sourceChartPath);
  const bytes = archive[normalizedSourcePath];
  if (!bytes) return null;
  const extension = normalizedSourcePath.split(".").pop() || "xml";
  const directory = dirname(normalizedSourcePath);
  const targetPath = nextIndexedPath(archive, directory, "chartClipboard", extension);
  archive[targetPath] = cloneBytes(bytes);
  const sourceRelationshipsPath = relsPathForDocument(normalizedSourcePath);
  const sourceRelationships = archive[sourceRelationshipsPath];
  if (sourceRelationships) {
    archive[relsPathForDocument(targetPath)] = cloneBytes(sourceRelationships);
  }
  copyContentTypeOverride(archive, normalizedSourcePath, targetPath);
  return targetPath;
}

function ensureDrawingNamespace(document: XMLDocument) {
  const root = document.documentElement;
  if (!root) return;
  if (!root.getAttribute("xmlns:r")) {
    root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", REL_NS);
  }
}

function ensureSheetDrawingReference(assets: WorkbookImageAssets, workbookSheetIndex: number, drawingPath: string) {
  const sheetPath = assets.sheetPathsByWorkbookSheetIndex[workbookSheetIndex];
  if (!sheetPath) return false;
  const sheetXml = readArchiveText(assets.archive, sheetPath);
  const sheetDocument = sheetXml ? parseXml(sheetXml) : null;
  if (!sheetDocument) return false;
  if (getLocalElements(sheetDocument, "drawing").length > 0) return true;

  const relationshipsPath = relsPathForDocument(sheetPath);
  const relationshipsDocument = ensureRelationshipsDocument(assets.archive, relationshipsPath);
  if (!relationshipsDocument) return false;
  const usedIds = new Set(
    getLocalElements(relationshipsDocument, "Relationship")
      .map((node) => node.getAttribute("Id"))
      .filter((value): value is string => Boolean(value)),
  );
  let index = 1;
  while (usedIds.has(`rIdXlsxDrawing${index}`)) index += 1;
  const relationshipId = `rIdXlsxDrawing${index}`;
  const relationshipNode = relationshipsDocument.createElementNS(PKG_REL_NS, "Relationship");
  relationshipNode.setAttribute("Id", relationshipId);
  relationshipNode.setAttribute("Type", DRAWING_REL_TYPE);
  relationshipNode.setAttribute("Target", relativeArchivePath(sheetPath, drawingPath));
  relationshipsDocument.documentElement?.appendChild(relationshipNode);

  ensureDrawingNamespace(sheetDocument);
  const drawingNode = sheetDocument.createElementNS(SPREADSHEET_NS, "drawing");
  drawingNode.setAttributeNS(REL_NS, "r:id", relationshipId);
  insertSheetChild(sheetDocument, drawingNode);

  assets.archive[sheetPath] = strToU8(serializeXml(sheetDocument));
  assets.archive[normalizeArchivePath(relationshipsPath)] = strToU8(serializeXml(relationshipsDocument));
  return true;
}

function ensureTargetDrawing(assets: WorkbookImageAssets, workbookSheetIndex: number): TargetDrawing | null {
  const sheetPath = assets.sheetPathsByWorkbookSheetIndex[workbookSheetIndex];
  if (!sheetPath) return null;

  const existingAttachment = assets.sheetOrigins[workbookSheetIndex]?.attachments[0] ?? null;
  if (existingAttachment) {
    const drawingXml = readArchiveText(assets.archive, existingAttachment.drawingPath);
    const drawingDocument = drawingXml ? parseXml(drawingXml) : null;
    const relationshipsPath = existingAttachment.drawingRelsPath ?? relsPathForDocument(existingAttachment.drawingPath);
    const relationshipsDocument = ensureRelationshipsDocument(assets.archive, relationshipsPath);
    if (drawingDocument && relationshipsDocument) {
      return {
        drawingDocument,
        drawingPath: normalizeArchivePath(existingAttachment.drawingPath),
        relationshipsDocument,
        relationshipsPath: normalizeArchivePath(relationshipsPath),
      };
    }
  }

  const drawingPath = nextIndexedPath(assets.archive, "xl/drawings", "drawingClipboard", "xml");
  const relationshipsPath = relsPathForDocument(drawingPath);
  const drawingDocument = parseXml(`<xdr:wsDr xmlns:xdr="${SPREADSHEET_DRAWING_NS}"></xdr:wsDr>`);
  const relationshipsDocument = parseXml(`<Relationships xmlns="${PKG_REL_NS}"></Relationships>`);
  if (!drawingDocument || !relationshipsDocument) return null;

  ensureContentTypeOverride(assets.archive, drawingPath, DRAWING_CONTENT_TYPE);
  assets.archive[drawingPath] = strToU8(serializeXml(drawingDocument));
  assets.archive[relationshipsPath] = strToU8(serializeXml(relationshipsDocument));
  if (!ensureSheetDrawingReference(assets, workbookSheetIndex, drawingPath)) return null;
  appendWorkbookSheetDrawingAttachment(assets, workbookSheetIndex, {
    drawingPath,
    drawingRelsPath: relationshipsPath,
    mediaPaths: [],
  });
  return {
    drawingDocument,
    drawingPath,
    relationshipsDocument,
    relationshipsPath,
  };
}

function resolveImageSource(
  assets: WorkbookImageAssets,
  source: WorkbookDrawingCloneSource,
): SourceDrawing | null {
  const origin = source.kind === "image"
    ? assets.imageOriginsById.get(source.id)
    : assets.shapeOriginsById.get(source.id);
  if (!origin) return null;
  if (origin.workbookSheetIndex !== source.workbookSheetIndex) return null;

  const model = source.kind === "image"
    ? (assets.imagesByWorkbookSheetIndex[source.workbookSheetIndex] ?? []).find((entry) => entry.id === source.id)
    : (assets.shapesByWorkbookSheetIndex[source.workbookSheetIndex] ?? []).find((entry) => entry.id === source.id);
  if (!model) return null;

  const attachments = assets.sheetOrigins[source.workbookSheetIndex]?.attachments ?? [];
  const candidates = origin.drawingPath
    ? attachments.filter((attachment) => attachment.drawingPath === origin.drawingPath)
    : attachments;
  for (const attachment of candidates) {
    const drawingXml = readArchiveText(assets.archive, attachment.drawingPath);
    const drawingDocument = drawingXml ? parseXml(drawingXml) : null;
    if (!drawingDocument) continue;
    const anchorNode = findAnchorForOrigin(drawingDocument, origin as WorkbookShapeOrigin);
    if (!anchorNode) continue;
    return {
      anchorNode,
      drawingDocument,
      drawingPath: normalizeArchivePath(attachment.drawingPath),
      model,
      origin,
    };
  }
  return null;
}

function resolveChartSource(
  assets: WorkbookImageAssets,
  chartAssets: WorkbookChartAssets,
  source: WorkbookDrawingCloneSource,
): SourceDrawing | null {
  const origin = chartAssets.chartOriginsById.get(source.id);
  if (!origin) return null;
  const chart = (chartAssets.chartsByWorkbookSheetIndex[origin.workbookSheetIndex] ?? [])
    .find((entry) => entry.id === source.id);
  if (!chart) return null;

  const drawingXml = readArchiveText(assets.archive, origin.drawingPath);
  const drawingDocument = drawingXml ? parseXml(drawingXml) : null;
  if (!drawingDocument) return null;
  const anchors = createImageAnchorNodes(drawingDocument);
  const anchorNode = origin.objectId
    ? anchors.find((anchor) => getFirstDescendant(anchor, "cNvPr")?.getAttribute("id") === origin.objectId)
      ?? anchors[origin.anchorIndex]
    : anchors[origin.anchorIndex];
  if (!anchorNode) return null;
  return {
    anchorNode,
    drawingDocument,
    drawingPath: normalizeArchivePath(origin.drawingPath),
    model: chart,
    origin,
  };
}

function uniqueModelId(kind: XlsxDrawingClipboardKind, workbookSheetIndex: number, used: Set<string>) {
  let index = 1;
  let candidate = `clipboard-${kind}-${workbookSheetIndex}-${index}`;
  while (used.has(candidate)) {
    index += 1;
    candidate = `clipboard-${kind}-${workbookSheetIndex}-${index}`;
  }
  return candidate;
}

function usedImageIds(assets: WorkbookImageAssets) {
  return new Set([
    ...assets.imageOriginsById.keys(),
    ...assets.imagesByWorkbookSheetIndex.flat().map((image) => image.id),
  ]);
}

function usedShapeIds(assets: WorkbookImageAssets) {
  return new Set([
    ...assets.shapeOriginsById.keys(),
    ...assets.shapesByWorkbookSheetIndex.flat().map((shape) => shape.id),
  ]);
}

function usedChartIds(chartAssets: WorkbookChartAssets) {
  return new Set([
    ...chartAssets.chartOriginsById.keys(),
    ...chartAssets.chartsByWorkbookSheetIndex.flat().map((chart) => chart.id),
  ]);
}

function nextZIndex(models: Array<{ zIndex?: number }>) {
  return models.reduce((max, model) => Math.max(max, Number(model.zIndex) || 0), 0) + 1;
}

function cloneSeries(model: XlsxChart, chartId: string) {
  return model.series.map((series, index) => ({
    ...series,
    id: `${chartId}-series-${index}`,
  }));
}

function cloneChartModel(model: XlsxChart, chartId: string) {
  return {
    ...model,
    axes: model.axes.map((axis) => ({ ...axis })),
    categoryAxis: model.categoryAxis ? { ...model.categoryAxis } : model.categoryAxis,
    chartPath: undefined,
    id: chartId,
    series: cloneSeries(model, chartId),
    seriesAxis: model.seriesAxis ? { ...model.seriesAxis } : model.seriesAxis,
    typeGroups: model.typeGroups?.map((group) => ({
      ...group,
      series: group.series.map((series, index) => ({
        ...series,
        id: `${chartId}-series-${index}`,
      })),
    })),
    valueAxis: model.valueAxis ? { ...model.valueAxis } : model.valueAxis,
  } satisfies XlsxChart;
}

function removeUnusedOrigin(assets: WorkbookImageAssets, kind: XlsxDrawingClipboardKind, id: string, workbookSheetIndex: number) {
  if (kind === "image") {
    assets.imageOriginsById.delete(id);
    assets.imagesByWorkbookSheetIndex[workbookSheetIndex] = (
      assets.imagesByWorkbookSheetIndex[workbookSheetIndex] ?? []
    ).filter((image) => image.id !== id);
    return;
  }
  assets.shapeOriginsById.delete(id);
  assets.shapesByWorkbookSheetIndex[workbookSheetIndex] = (
    assets.shapesByWorkbookSheetIndex[workbookSheetIndex] ?? []
  ).filter((shape) => shape.id !== id);
}

/**
 * Captures one drawing into a clipboard item.
 *
 * The source sheet state is only needed to express the drawing's rendered size
 * and its offset from the copied group in EMU, so a clipboard payload pastes
 * with the same geometry even when the target sheet has different column
 * widths.
 */
export function createDrawingClipboardItem(
  kind: XlsxDrawingClipboardKind,
  model: DrawingModel,
  sheetState: WorkbookSheetState | null,
): XlsxDrawingClipboardItem {
  const rect = anchorToAbsoluteRect(model.anchor, sheetState);
  return {
    anchor: model.anchor,
    kind,
    model,
    relativeAnchor: { x: 0, y: 0 },
    size: { cx: rect.cx, cy: rect.cy },
    sourceId: model.id,
    sourceSheetId: model.workbookSheetIndex,
  };
}

/**
 * Captures several drawings as one group.
 *
 * Offsets are measured from the group's top-left, so pasting the group at a
 * cell keeps the drawings' relative arrangement.
 */
export function createDrawingClipboardGroup(
  drawings: Array<{ kind: XlsxDrawingClipboardKind; model: DrawingModel }>,
  sheetState: WorkbookSheetState | null,
): XlsxDrawingClipboardItem[] {
  const items = drawings.map((drawing) => createDrawingClipboardItem(drawing.kind, drawing.model, sheetState));
  if (items.length === 0) return items;
  const rects = items.map((item) => anchorToAbsoluteRect(item.anchor, sheetState));
  const originX = Math.min(...rects.map((rect) => rect.x));
  const originY = Math.min(...rects.map((rect) => rect.y));
  return items.map((item, index) => {
    const rect = rects[index];
    return {
      ...item,
      relativeAnchor: {
        x: Math.max(0, rect.x - originX),
        y: Math.max(0, rect.y - originY),
      },
    };
  });
}

/**
 * Clones one drawing from its source sheet into another drawing part.
 *
 * The DrawingML anchor is copied (including pictures, shapes and chart
 * references), relationships are remapped, and a new model + origin are
 * registered in `imageAssets`/`chartAssets` so the viewer and the export path
 * see the pasted drawing immediately.
 *
 * @returns The new model, or `null` when the source drawing or target sheet
 *   cannot be resolved.
 */
export function cloneWorkbookDrawing(
  imageAssets: WorkbookImageAssets,
  chartAssets: WorkbookChartAssets | null,
  source: WorkbookDrawingCloneSource,
  targetWorkbookSheetIndex: number,
  targetAnchor: XlsxImageAnchor,
  options: WorkbookDrawingCloneOptions = {},
): WorkbookDrawingCloneResult | null {
  const resolved = source.kind === "chart"
    ? (chartAssets ? resolveChartSource(imageAssets, chartAssets, source) : null)
    : resolveImageSource(imageAssets, source);
  if (!resolved) return null;
  const target = ensureTargetDrawing(imageAssets, targetWorkbookSheetIndex);
  if (!target) return null;

  const sourceAnchorObjectId = (resolved.origin as WorkbookImageOrigin).objectId;
  const sourceObjectKind = (resolved.origin as WorkbookImageOrigin).objectKind;
  const sourceCNvPrs = descendantsByLocalName(resolved.anchorNode, "cNvPr");
  const selectedCNvPrIndex = sourceAnchorObjectId
    ? sourceCNvPrs.findIndex((node) => node.getAttribute("id") === sourceAnchorObjectId)
    : -1;

  const clone = target.drawingDocument.importNode(resolved.anchorNode, true) as Element;
  const targetCNvPrs = descendantsByLocalName(clone, "cNvPr");
  let nextObjectId = getLocalElements(target.drawingDocument, "cNvPr")
    .reduce((max, node) => Math.max(max, Number(node.getAttribute("id")) || 0), 0) + 1;
  for (const node of targetCNvPrs) {
    node.setAttribute("id", String(nextObjectId));
    nextObjectId += 1;
  }
  const newObjectId = selectedCNvPrIndex >= 0
    ? targetCNvPrs[selectedCNvPrIndex]?.getAttribute("id") ?? undefined
    : targetCNvPrs[0]?.getAttribute("id") ?? undefined;

  const sourceRelationships = readRelationships(
    imageAssets.archive,
    relsPathForDocument(resolved.drawingPath),
    resolved.drawingPath,
  );
  const relationshipIdMap = new Map<string, string>();
  let clonedChartPath: string | null = null;
  for (const sourceRelationshipId of relationshipIdsInSubtree(clone)) {
    const relationship = sourceRelationships.get(sourceRelationshipId);
    if (!relationship) continue;
    const relationshipId = nextRelationshipId(target.relationshipsDocument);
    let targetValue = relationship.target;
    if (relationship.targetMode !== "External") {
      if (relationship.type === CHART_REL_TYPE || relationship.type === CHART_EX_REL_TYPE) {
        const chartPath = cloneChartPart(imageAssets.archive, relationship.target);
        if (!chartPath) continue;
        clonedChartPath = chartPath;
        targetValue = relativeArchivePath(target.drawingPath, chartPath);
      } else {
        targetValue = relativeArchivePath(target.drawingPath, relationship.target);
      }
    }
    const node = target.relationshipsDocument.createElementNS(PKG_REL_NS, "Relationship");
    node.setAttribute("Id", relationshipId);
    node.setAttribute("Type", relationship.type);
    node.setAttribute("Target", targetValue);
    if (relationship.targetMode) node.setAttribute("TargetMode", relationship.targetMode);
    target.relationshipsDocument.documentElement?.appendChild(node);
    relationshipIdMap.set(sourceRelationshipId, relationshipId);
  }
  for (const [sourceId, relationshipId] of relationshipIdMap) {
    replaceRelationshipId(clone, sourceId, relationshipId);
  }
  const anchoredClone = updateAnchorNode(clone, targetAnchor);

  target.drawingDocument.documentElement?.appendChild(anchoredClone);
  imageAssets.archive[normalizeArchivePath(target.drawingPath)] = strToU8(serializeXml(target.drawingDocument));
  imageAssets.archive[normalizeArchivePath(target.relationshipsPath)] = strToU8(serializeXml(target.relationshipsDocument));
  const anchorIndex = createImageAnchorNodes(target.drawingDocument).indexOf(anchoredClone);
  if (anchorIndex < 0) return null;

  if (source.kind === "image") {
    const image = resolved.model as XlsxImage;
    const id = options.id ?? uniqueModelId("image", targetWorkbookSheetIndex, usedImageIds(imageAssets));
    const model: XlsxImage = {
      ...image,
      anchor: targetAnchor,
      editable: true,
      id,
      sheetIndex: options.visibleSheetIndex ?? targetWorkbookSheetIndex,
      workbookSheetIndex: targetWorkbookSheetIndex,
      zIndex: nextZIndex(imageAssets.imagesByWorkbookSheetIndex[targetWorkbookSheetIndex] ?? []),
    };
    imageAssets.imagesByWorkbookSheetIndex[targetWorkbookSheetIndex] = [
      ...(imageAssets.imagesByWorkbookSheetIndex[targetWorkbookSheetIndex] ?? []),
      model,
    ];
    imageAssets.imageOriginsById.set(id, {
      anchorIndex,
      drawingPath: target.drawingPath,
      objectId: newObjectId,
      objectKind: "pic",
      workbookSheetIndex: targetWorkbookSheetIndex,
    });
    return { id, kind: "image", model };
  }

  if (source.kind === "shape") {
    const shape = resolved.model as XlsxShape;
    const id = options.id ?? uniqueModelId("shape", targetWorkbookSheetIndex, usedShapeIds(imageAssets));
    const model: XlsxShape = {
      ...shape,
      anchor: targetAnchor,
      id,
      sheetIndex: options.visibleSheetIndex ?? targetWorkbookSheetIndex,
      workbookSheetIndex: targetWorkbookSheetIndex,
      zIndex: nextZIndex(imageAssets.shapesByWorkbookSheetIndex[targetWorkbookSheetIndex] ?? []),
    };
    imageAssets.shapesByWorkbookSheetIndex[targetWorkbookSheetIndex] = [
      ...(imageAssets.shapesByWorkbookSheetIndex[targetWorkbookSheetIndex] ?? []),
      model,
    ];
    imageAssets.shapeOriginsById.set(id, {
      anchorIndex,
      drawingPath: target.drawingPath,
      objectId: newObjectId,
      objectKind: sourceObjectKind ?? "sp",
      workbookSheetIndex: targetWorkbookSheetIndex,
    });
    return { id, kind: "shape", model };
  }

  if (!chartAssets) return null;
  const chart = resolved.model as XlsxChart;
  const id = options.id ?? uniqueModelId("chart", targetWorkbookSheetIndex, usedChartIds(chartAssets));
  const model: XlsxChart = {
    ...cloneChartModel(chart, id),
    anchor: targetAnchor,
    chartPath: clonedChartPath ?? undefined,
    sheetIndex: options.visibleSheetIndex ?? targetWorkbookSheetIndex,
    workbookSheetIndex: targetWorkbookSheetIndex,
    zIndex: nextZIndex(chartAssets.chartsByWorkbookSheetIndex[targetWorkbookSheetIndex] ?? []),
  };
  chartAssets.chartsByWorkbookSheetIndex[targetWorkbookSheetIndex] = [
    ...(chartAssets.chartsByWorkbookSheetIndex[targetWorkbookSheetIndex] ?? []),
    model,
  ];
  chartAssets.chartOriginsById.set(id, {
    anchorIndex,
    anchor: targetAnchor,
    chartKind: (resolved.origin as WorkbookChartOrigin).chartKind,
    chartPath: clonedChartPath,
    drawingPath: target.drawingPath,
    objectId: newObjectId,
    workbookSheetIndex: targetWorkbookSheetIndex,
  });
  return { id, kind: "chart", model };
}

/** Deletes a drawing from the workbook assets after a cut-paste moved it. */
export function removeWorkbookDrawing(
  imageAssets: WorkbookImageAssets,
  chartAssets: WorkbookChartAssets | null,
  source: WorkbookDrawingCloneSource,
): boolean {
  if (source.kind === "chart") {
    if (!chartAssets) return false;
    const origin = chartAssets.chartOriginsById.get(source.id);
    if (!origin) return false;
    const drawingXml = readArchiveText(imageAssets.archive, origin.drawingPath);
    const drawingDocument = drawingXml ? parseXml(drawingXml) : null;
    if (!drawingDocument) return false;
    const anchors = createImageAnchorNodes(drawingDocument);
    const anchorNode = origin.objectId
      ? anchors.find((anchor) => getFirstDescendant(anchor, "cNvPr")?.getAttribute("id") === origin.objectId)
        ?? anchors[origin.anchorIndex]
      : anchors[origin.anchorIndex];
    const graphicFrame = anchorNode ? getFirstDescendant(anchorNode, "graphicFrame") : null;
    if (!anchorNode || !graphicFrame) return false;
    graphicFrame.parentNode?.removeChild(graphicFrame);
    const remaining = Array.from(anchorNode.childNodes).some(
      (child) => isElementNode(child)
        && ["pic", "sp", "cxnSp", "graphicFrame"].includes((child as Element).localName),
    );
    if (!remaining) anchorNode.parentNode?.removeChild(anchorNode);
    imageAssets.archive[normalizeArchivePath(origin.drawingPath)] = strToU8(serializeXml(drawingDocument));
    chartAssets.chartOriginsById.delete(source.id);
    chartAssets.chartsByWorkbookSheetIndex[origin.workbookSheetIndex] = (
      chartAssets.chartsByWorkbookSheetIndex[origin.workbookSheetIndex] ?? []
    ).filter((chart) => chart.id !== source.id);
    return true;
  }

  const origin = source.kind === "image"
    ? imageAssets.imageOriginsById.get(source.id)
    : imageAssets.shapeOriginsById.get(source.id);
  if (!origin) return false;
  const drawingXml = readArchiveText(imageAssets.archive, origin.drawingPath ?? "");
  const drawingDocument = drawingXml ? parseXml(drawingXml) : null;
  if (!drawingDocument) return false;
  const anchorNode = findAnchorForOrigin(drawingDocument, origin);
  if (!anchorNode) return false;
  const objectNode = origin.objectId
    ? findNonVisualObject(anchorNode, origin.objectId, origin.objectKind)
    : Array.from(anchorNode.childNodes).find(
      (child) => isElementNode(child) && (child as Element).localName === (origin.objectKind ?? "pic"),
    ) as Element | undefined;
  if (!objectNode) return false;
  objectNode.parentNode?.removeChild(objectNode);
  const remaining = Array.from(anchorNode.childNodes).some(
    (child) => isElementNode(child)
      && ["pic", "sp", "cxnSp", "graphicFrame"].includes((child as Element).localName),
  );
  if (!remaining) anchorNode.parentNode?.removeChild(anchorNode);
  imageAssets.archive[normalizeArchivePath(origin.drawingPath!)] = strToU8(serializeXml(drawingDocument));
  removeUnusedOrigin(imageAssets, source.kind, source.id, origin.workbookSheetIndex);
  return true;
}
