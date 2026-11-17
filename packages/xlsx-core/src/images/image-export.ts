import { strToU8, unzipSync, zipSync } from "fflate";
import type {
  XlsxFormControl,
  XlsxImage,
  XlsxShape,
} from "../types";
import {
  type ArchiveEntries,
  CONTENT_TYPES_NS,
  DRAWING_CONTENT_TYPE,
  DRAWING_REL_TYPE,
  PKG_REL_NS,
  SPREADSHEET_NS,
  REL_NS,
  type WorkbookFormControlOrigin,
  type WorkbookImageOrigin,
  type WorkbookImageSheetOrigin,
  type WorkbookShapeOrigin,
  cloneBytes,
  createImageAnchorNodes,
  getChildElements,
  getFirstDescendant,
  getLocalElements,
  normalizeArchivePath,
  parseRelationships,
  parseXml,
  readArchiveText,
  relativeArchivePath,
  relsPathForDocument,
  serializeXml,
  setChildText,
} from "./image-parser";
import {
  parseDrawingObjects,
  parseSheetFormControls,
  enrichFormControlsWithHiddenShapes,
} from "./drawing-parser";
import {
  type WorkbookImageAssets,
  type WorkbookSheetState,
  parseWorkbookSheets,
  parseWorkbookStructureAssetsFromArchive,
} from "./grid-render";
import type { XlsxRuntimeLimits } from "../resource-limits";
import { validateXlsxImageAssets } from "./image-budget";

// ── Revoke object URLs ─────────────────────────────────────────────────────

export function revokeWorkbookImageAssets(assets: WorkbookImageAssets | null) {
  if (!assets) {
    return;
  }

  for (const objectUrl of assets.objectUrls) {
    if (objectUrl.startsWith("blob:")) {
      URL.revokeObjectURL(objectUrl);
    }
  }
  assets.objectUrls.splice(0, assets.objectUrls.length);
}

// ── Parse workbook image assets ───────────────────────────────────────────

export function parseWorkbookStructureAssets(
  bytes: Uint8Array,
  options?: {
    includeCachedFormulaValues?: boolean;
    themePalette?: import("../types").XlsxThemePalette | null;
  }
): import("./grid-render").WorkbookStructureAssets {
  const archive = unzipSync(bytes);
  const {
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    themePalette
  } = parseWorkbookStructureAssetsFromArchive(archive, options);

  return {
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    themePalette
  };
}

export function parseWorkbookChartStyleAssets(bytes: Uint8Array): import("./grid-render").WorkbookChartStyleAssets {
  const archive = unzipSync(bytes);
  const {
    themePalette,
    workbookSheets
  } = parseWorkbookStructureAssetsFromArchive(archive);
  const sheetOrigins: Array<WorkbookImageSheetOrigin | null> = [];

  workbookSheets.forEach((sheet, workbookSheetIndex) => {
    const sheetRelationships = parseRelationships(archive, relsPathForDocument(sheet.path), sheet.path);
    const attachments: XlsxImageAttachment[] = [];

    for (const relationship of sheetRelationships.values()) {
      if (relationship.type !== DRAWING_REL_TYPE) {
        continue;
      }

      const drawingPath = relationship.target;
      const drawingRelsPath = relsPathForDocument(drawingPath);
      attachments.push({
        drawingPath,
        drawingRelsPath: archive[drawingRelsPath] ? drawingRelsPath : null,
        mediaPaths: []
      });
    }

    sheetOrigins[workbookSheetIndex] = attachments.length > 0
      ? {
          attachments,
          workbookSheetIndex
        }
      : null;
  });

  return {
    archive,
    sheetOrigins,
    themePalette
  };
}

type XlsxImageAttachment = {
  drawingPath: string;
  drawingRelsPath: string | null;
  mediaPaths: string[];
};

export function parseWorkbookImageAssets(bytes: Uint8Array, limits?: XlsxRuntimeLimits): WorkbookImageAssets {
  const archive = unzipSync(bytes);
  validateXlsxImageAssets(archive, limits);
  const {
    contentTypes,
    namedCellStyleByName,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    theme,
    themePalette,
    workbookSheets
  } = parseWorkbookStructureAssetsFromArchive(archive);
  const objectUrls: string[] = [];
  const formControlOriginsById = new Map<string, WorkbookFormControlOrigin>();
  const formControlsByWorkbookSheetIndex: XlsxFormControl[][] = [];
  const imagesByWorkbookSheetIndex: XlsxImage[][] = [];
  const shapesByWorkbookSheetIndex: XlsxShape[][] = [];
  const sheetOrigins: Array<WorkbookImageSheetOrigin | null> = [];
  const imageOriginsById = new Map<string, WorkbookImageOrigin>();
  const shapeOriginsById = new Map<string, WorkbookShapeOrigin>();
  const sheetPathsByWorkbookSheetIndex = workbookSheets.map((sheet) => sheet.path);

  try {
    workbookSheets.forEach((sheet, workbookSheetIndex) => {
    const sheetRelationships = parseRelationships(archive, relsPathForDocument(sheet.path), sheet.path);
    const attachments: XlsxImageAttachment[] = [];
    const imageList: XlsxImage[] = [];
    const shapeList: XlsxShape[] = [];
    let zIndexBase = 1;

    for (const relationship of sheetRelationships.values()) {
      if (relationship.type !== DRAWING_REL_TYPE) {
        continue;
      }

      const drawingPath = relationship.target;
      const drawingRelsPath = relsPathForDocument(drawingPath);
      const drawingImages = parseDrawingObjects(
        archive,
        contentTypes,
        drawingPath,
        objectUrls,
        workbookSheetIndex,
        zIndexBase,
        theme,
        sheetStatesByWorkbookSheetIndex[workbookSheetIndex] ?? null,
        imageOriginsById,
        shapeOriginsById
      );
      imageList.push(...drawingImages.images);
      shapeList.push(...drawingImages.shapes);
      zIndexBase += drawingImages.images.length + drawingImages.shapes.length + 10;
      attachments.push({
        drawingPath,
        drawingRelsPath: archive[drawingRelsPath] ? drawingRelsPath : null,
        mediaPaths: drawingImages.mediaPaths
      });
    }

    const formControlList = parseSheetFormControls(
      archive,
      sheet.path,
      sheetRelationships,
      workbookSheetIndex,
      zIndexBase,
      formControlOriginsById
    );
    const visibleShapeList = shapeList.filter((shape) => !shape.hidden);
    const enrichedFormControlList = enrichFormControlsWithHiddenShapes(
      formControlList,
      shapeList,
      sheetStatesByWorkbookSheetIndex[workbookSheetIndex] ?? null
    );

    formControlsByWorkbookSheetIndex[workbookSheetIndex] = enrichedFormControlList;
    imagesByWorkbookSheetIndex[workbookSheetIndex] = imageList;
    shapesByWorkbookSheetIndex[workbookSheetIndex] = visibleShapeList;
    sheetOrigins[workbookSheetIndex] = attachments.length > 0
      ? {
          attachments,
          workbookSheetIndex
        }
      : null;
    });

    return {
      archive,
      formControlOriginsById,
      formControlsByWorkbookSheetIndex,
      imageOriginsById,
      imagesByWorkbookSheetIndex,
      namedCellStyleByName,
      objectUrls,
      shapeOriginsById,
      shapesByWorkbookSheetIndex,
      sheetPathsByWorkbookSheetIndex,
      sheetOrigins,
      sheetStatesByWorkbookSheetIndex,
      styleById,
      tableMetadataByWorkbookSheetIndex,
      tableStyleByName,
      themePalette
    };
  } catch (error) {
    for (const objectUrl of objectUrls) {
      if (objectUrl.startsWith("blob:")) URL.revokeObjectURL(objectUrl);
    }
    objectUrls.splice(0, objectUrls.length);
    throw error;
  }
}

// ── Update image anchor ────────────────────────────────────────────────────

const SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const ANCHOR_CHILD_NAMES = new Set(["from", "to", "pos", "ext"]);

function createMarkerNode(document: XMLDocument, name: string, marker: { col: number; colOffsetEmu: number; row: number; rowOffsetEmu: number }) {
  const markerNode = document.createElementNS(SPREADSHEET_DRAWING_NS, name);
  for (const [childName, value] of [
    ["col", marker.col],
    ["colOff", marker.colOffsetEmu],
    ["row", marker.row],
    ["rowOff", marker.rowOffsetEmu],
  ] as const) {
    const child = document.createElementNS(SPREADSHEET_DRAWING_NS, childName);
    child.textContent = String(Math.max(0, Math.round(Number(value) || 0)));
    markerNode.appendChild(child);
  }
  return markerNode;
}

function createExtNode(document: XMLDocument, anchor: { cx: number; cy: number }) {
  const ext = document.createElementNS(SPREADSHEET_DRAWING_NS, "ext");
  ext.setAttribute("cx", String(Math.max(0, Math.round(anchor.cx))));
  ext.setAttribute("cy", String(Math.max(0, Math.round(anchor.cy))));
  return ext;
}

function updateAnchorNode(anchorNode: Element, anchor: XlsxImage["anchor"]) {
  const document = anchorNode.ownerDocument as XMLDocument | null;
  if (!document) return;
  const desiredName = anchor.kind === "two-cell"
    ? "twoCellAnchor"
    : anchor.kind === "one-cell"
      ? "oneCellAnchor"
      : "absoluteAnchor";
  let target = anchorNode;

  if (anchorNode.localName !== desiredName) {
    const replacement = document.createElementNS(SPREADSHEET_DRAWING_NS, desiredName);
    for (const attribute of Array.from(anchorNode.attributes)) {
      if (attribute.name !== "editAs") {
        replacement.setAttribute(attribute.name, attribute.value);
      }
    }
    const preserved = Array.from(anchorNode.childNodes).filter(
      (child) => child.nodeType !== 1 || !ANCHOR_CHILD_NAMES.has((child as Element).localName)
    );
    for (const child of preserved) replacement.appendChild(child);
    anchorNode.parentNode?.replaceChild(replacement, anchorNode);
    target = replacement;
  }

  for (const child of Array.from(target.childNodes)) {
    if (child.nodeType === 1 && ANCHOR_CHILD_NAMES.has((child as Element).localName)) {
      target.removeChild(child);
    }
  }

  const parts: Element[] = [];
  if (anchor.kind === "two-cell") {
    parts.push(createMarkerNode(document, "from", anchor.from));
    parts.push(createMarkerNode(document, "to", anchor.to));
  } else if (anchor.kind === "one-cell") {
    parts.push(createMarkerNode(document, "from", anchor.from));
    parts.push(createExtNode(document, anchor.sizeEmu));
  } else {
    const pos = document.createElementNS(SPREADSHEET_DRAWING_NS, "pos");
    pos.setAttribute("x", String(Math.max(0, Math.round(anchor.positionEmu.x))));
    pos.setAttribute("y", String(Math.max(0, Math.round(anchor.positionEmu.y))));
    parts.push(pos);
    parts.push(createExtNode(document, anchor.sizeEmu));
  }
  const firstChild = target.firstChild;
  for (const part of parts) {
    target.insertBefore(part, firstChild);
  }
}

function findNonVisualObject(root: Element, objectId: string, objectKind?: string): Element | null {
  for (const child of Array.from(root.childNodes)) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if (
      (!objectKind || element.localName === objectKind)
      && (element.localName === "pic" || element.localName === "sp" || element.localName === "cxnSp" || element.localName === "graphicFrame")
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
  const anchorNodes = createImageAnchorNodes(document);
  if (origin.objectId) {
    const match = anchorNodes.find((anchorNode) => Boolean(findNonVisualObject(anchorNode, origin.objectId!)));
    if (match) return match;
  }
  return anchorNodes[origin.anchorIndex] ?? null;
}

function resolveDrawingDocument(assets: WorkbookImageAssets, origin: WorkbookImageOrigin) {
  const attachments = assets.sheetOrigins[origin.workbookSheetIndex]?.attachments ?? [];
  const candidates = origin.drawingPath
    ? attachments.filter((attachment) => attachment.drawingPath === origin.drawingPath)
    : attachments;
  for (const attachment of candidates) {
    const drawingXml = readArchiveText(assets.archive, attachment.drawingPath);
    if (!drawingXml) continue;
    const drawingDocument = parseXml(drawingXml);
    if (!drawingDocument) continue;
    const anchorNode = findAnchorForOrigin(drawingDocument, origin);
    if (anchorNode) {
      return { anchorNode, drawingDocument, drawingPath: attachment.drawingPath };
    }
  }
  return null;
}

function hasDrawableChild(anchorNode: Element) {
  return Array.from(anchorNode.childNodes).some(
    (child) => child.nodeType === 1
      && ["pic", "sp", "cxnSp", "graphicFrame"].includes((child as Element).localName)
  );
}

function removeDrawingObject(assets: WorkbookImageAssets, origin: WorkbookImageOrigin): boolean {
  const resolved = resolveDrawingDocument(assets, origin);
  if (!resolved) return false;
  const { anchorNode, drawingDocument, drawingPath } = resolved;
  const objectNode = origin.objectId
    ? findNonVisualObject(anchorNode, origin.objectId, origin.objectKind)
    : Array.from(anchorNode.childNodes).find(
      (child) => child.nodeType === 1 && (child as Element).localName === (origin.objectKind ?? "pic")
    ) as Element | undefined;
  const target = objectNode ?? anchorNode;
  target.parentNode?.removeChild(target);
  if (target !== anchorNode && !hasDrawableChild(anchorNode)) {
    anchorNode.parentNode?.removeChild(anchorNode);
  }
  assets.archive[normalizeArchivePath(drawingPath)] = strToU8(serializeXml(drawingDocument));
  return true;
}

function updateDrawingAnchor(assets: WorkbookImageAssets, origin: WorkbookImageOrigin, anchor: XlsxImage["anchor"]): boolean {
  const resolved = resolveDrawingDocument(assets, origin);
  if (!resolved) return false;
  updateAnchorNode(resolved.anchorNode, anchor);
  assets.archive[normalizeArchivePath(resolved.drawingPath)] = strToU8(serializeXml(resolved.drawingDocument));
  return true;
}

export function updateWorkbookImageAnchor(
  assets: WorkbookImageAssets,
  imageId: string,
  anchor: XlsxImage["anchor"]
) {
  const origin = assets.imageOriginsById.get(imageId);
  if (!origin || !updateDrawingAnchor(assets, origin, anchor)) {
    return false;
  }
  const imageList = assets.imagesByWorkbookSheetIndex[origin.workbookSheetIndex] ?? [];
  const imageIndex = imageList.findIndex((image) => image.id === imageId);
  if (imageIndex >= 0) {
    imageList[imageIndex] = { ...imageList[imageIndex], anchor };
  }
  return true;
}

export function updateWorkbookShapeAnchor(
  assets: WorkbookImageAssets,
  shapeId: string,
  anchor: XlsxImage["anchor"]
) {
  const origin = assets.shapeOriginsById.get(shapeId);
  if (!origin || !updateDrawingAnchor(assets, origin, anchor)) {
    return false;
  }
  const shapeList = assets.shapesByWorkbookSheetIndex[origin.workbookSheetIndex] ?? [];
  const shapeIndex = shapeList.findIndex((shape) => shape.id === shapeId);
  if (shapeIndex >= 0) {
    shapeList[shapeIndex] = { ...shapeList[shapeIndex], anchor };
  }
  return true;
}

export function deleteWorkbookImage(assets: WorkbookImageAssets, imageId: string) {
  const origin = assets.imageOriginsById.get(imageId);
  if (!origin || !removeDrawingObject(assets, origin)) {
    return false;
  }
  assets.imageOriginsById.delete(imageId);
  assets.imagesByWorkbookSheetIndex[origin.workbookSheetIndex] = (
    assets.imagesByWorkbookSheetIndex[origin.workbookSheetIndex] ?? []
  ).filter((image) => image.id !== imageId);
  return true;
}

export function deleteWorkbookShape(assets: WorkbookImageAssets, shapeId: string) {
  const origin = assets.shapeOriginsById.get(shapeId);
  if (!origin || !removeDrawingObject(assets, origin)) {
    return false;
  }
  assets.shapeOriginsById.delete(shapeId);
  assets.shapesByWorkbookSheetIndex[origin.workbookSheetIndex] = (
    assets.shapesByWorkbookSheetIndex[origin.workbookSheetIndex] ?? []
  ).filter((shape) => shape.id !== shapeId);
  return true;
}

// ── Merge workbook image assets ────────────────────────────────────────────

function ensureRelationshipsDocument(archive: ArchiveEntries, relsPath: string) {
  const existing = readArchiveText(archive, relsPath);
  if (existing) {
    const parsed = parseXml(existing);
    if (parsed) {
      return parsed;
    }
  }

  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"></Relationships>`
  );
}

function ensureContentTypesDocument(archive: ArchiveEntries) {
  const existing = readArchiveText(archive, "[Content_Types].xml");
  if (existing) {
    const parsed = parseXml(existing);
    if (parsed) {
      return parsed;
    }
  }

  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CONTENT_TYPES_NS}"></Types>`
  );
}

function mergeContentTypeForPath(
  targetDocument: XMLDocument,
  originalDocument: XMLDocument | null,
  partPath: string
) {
  const normalizedPartName = `/${normalizeArchivePath(partPath)}`;
  const targetRoot = targetDocument.documentElement;
  if (!targetRoot) {
    return;
  }

  const existingOverride = getLocalElements(targetDocument, "Override").find(
    (node) => node.getAttribute("PartName") === normalizedPartName
  );
  if (!existingOverride && originalDocument) {
    const sourceOverride = getLocalElements(originalDocument, "Override").find(
      (node) => node.getAttribute("PartName") === normalizedPartName
    );
    if (sourceOverride) {
      targetRoot.appendChild(sourceOverride.cloneNode(true));
      return;
    }
  }

  const extension = normalizedPartName.split(".").pop()?.toLowerCase();
  if (!extension) {
    return;
  }

  const existingDefault = getLocalElements(targetDocument, "Default").find(
    (node) => (node.getAttribute("Extension") ?? "").toLowerCase() === extension
  );
  if (existingDefault) {
    return;
  }

  if (originalDocument) {
    const sourceDefault = getLocalElements(originalDocument, "Default").find(
      (node) => (node.getAttribute("Extension") ?? "").toLowerCase() === extension
    );
    if (sourceDefault) {
      targetRoot.appendChild(sourceDefault.cloneNode(true));
      return;
    }
  }

  if (extension === "xml") {
    const defaultNode = targetDocument.createElementNS(CONTENT_TYPES_NS, "Default");
    defaultNode.setAttribute("Extension", extension);
    defaultNode.setAttribute("ContentType", "application/xml");
    targetRoot.appendChild(defaultNode);
    return;
  }

  if (extension === "rels") {
    const defaultNode = targetDocument.createElementNS(CONTENT_TYPES_NS, "Default");
    defaultNode.setAttribute("Extension", extension);
    defaultNode.setAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
    targetRoot.appendChild(defaultNode);
  }
}

function ensureContentTypeOverride(
  targetDocument: XMLDocument,
  partPath: string,
  contentType: string
) {
  const normalizedPartName = `/${normalizeArchivePath(partPath)}`;
  const existingOverride = getLocalElements(targetDocument, "Override").find(
    (node) => node.getAttribute("PartName") === normalizedPartName
  );
  if (existingOverride) {
    existingOverride.setAttribute("ContentType", contentType);
    return;
  }

  const targetRoot = targetDocument.documentElement;
  if (!targetRoot) {
    return;
  }

  const overrideNode = targetDocument.createElementNS(CONTENT_TYPES_NS, "Override");
  overrideNode.setAttribute("PartName", normalizedPartName);
  overrideNode.setAttribute("ContentType", contentType);
  targetRoot.appendChild(overrideNode);
}

function removeDrawingReferences(sheetDocument: XMLDocument, relDocument: XMLDocument) {
  getLocalElements(sheetDocument, "drawing").forEach((node) => node.remove());
  getLocalElements(relDocument, "Relationship")
    .filter((node) => node.getAttribute("Type") === DRAWING_REL_TYPE)
    .forEach((node) => node.remove());
}

function nextRelationshipId(relDocument: XMLDocument) {
  const existingIds = new Set(
    getLocalElements(relDocument, "Relationship")
      .map((node) => node.getAttribute("Id"))
      .filter((value): value is string => Boolean(value))
  );

  let index = 1;
  while (existingIds.has(`rIdReactXlsxImage${index}`)) {
    index += 1;
  }

  return `rIdReactXlsxImage${index}`;
}

const DRAWING_SUCCESSOR_LOCAL_NAMES = new Set([
  "legacyDrawing",
  "legacyDrawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
]);

function appendSheetDrawingReference(
  sheetDocument: XMLDocument,
  relationshipId: string
) {
  const worksheet = sheetDocument.documentElement;
  if (!worksheet) {
    return;
  }

  const drawingNode = sheetDocument.createElementNS(SPREADSHEET_NS, "drawing");
  drawingNode.setAttributeNS(REL_NS, "r:id", relationshipId);

  const successor = Array.from(worksheet.childNodes).find(
    (child): child is Element => (
      child.nodeType === 1
      && DRAWING_SUCCESSOR_LOCAL_NAMES.has((child as Element).localName)
    )
  );
  if (successor) {
    worksheet.insertBefore(drawingNode, successor);
    return;
  }

  worksheet.appendChild(drawingNode);
}

function createRelationshipsDocument() {
  return parseXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"></Relationships>`
  );
}

function nextDrawingPartPath(archive: ArchiveEntries) {
  let index = 1;
  while (
    archive[`xl/drawings/drawingReactXlsx${index}.xml`]
    || archive[`xl/drawings/_rels/drawingReactXlsx${index}.xml.rels`]
  ) {
    index += 1;
  }
  return `xl/drawings/drawingReactXlsx${index}.xml`;
}

function nextDrawingRelationshipId(relDocument: XMLDocument) {
  const existingIds = new Set(
    getLocalElements(relDocument, "Relationship")
      .map((node) => node.getAttribute("Id"))
      .filter((value): value is string => Boolean(value))
  );

  let index = 1;
  while (existingIds.has(`rIdReactXlsxDrawing${index}`)) {
    index += 1;
  }

  return `rIdReactXlsxDrawing${index}`;
}

function remapRelationshipReferences(node: Node, idMap: Map<string, string>) {
  if (idMap.size === 0) {
    return;
  }

  if (node.nodeType === 1) {
    const element = node as Element;
    for (let index = 0; index < element.attributes.length; index += 1) {
      const attribute = element.attributes.item(index);
      if (!attribute) {
        continue;
      }
      const replacement = idMap.get(attribute.value);
      if (!replacement) {
        continue;
      }
      if (attribute.namespaceURI === REL_NS || attribute.name.startsWith("r:")) {
        element.setAttributeNS(attribute.namespaceURI, attribute.name, replacement);
      }
    }
  }

  for (let child = node.firstChild; child; child = child.nextSibling) {
    remapRelationshipReferences(child, idMap);
  }
}

function drawingChildElements(drawingDocument: XMLDocument) {
  return Array.from(drawingDocument.documentElement?.childNodes ?? []).filter(
    (node): node is Element => node.nodeType === 1
  );
}

// Copy a part plus every internal part reachable through its relationship file.
function copyArchivePartWithDependencies(
  targetArchive: ArchiveEntries,
  sourceArchive: ArchiveEntries,
  targetContentTypesDocument: XMLDocument,
  originalContentTypesDocument: XMLDocument | null,
  partPath: string,
  copiedPartPaths: Set<string>
) {
  const normalizedPartPath = normalizeArchivePath(partPath);
  if (!normalizedPartPath || copiedPartPaths.has(normalizedPartPath)) {
    return;
  }
  copiedPartPaths.add(normalizedPartPath);

  const partBytes = sourceArchive[normalizedPartPath];
  if (!partBytes) {
    return;
  }

  targetArchive[normalizedPartPath] = cloneBytes(partBytes);
  mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, normalizedPartPath);

  const relsPath = relsPathForDocument(normalizedPartPath);
  const relsBytes = sourceArchive[relsPath];
  if (!relsBytes) {
    return;
  }

  targetArchive[relsPath] = cloneBytes(relsBytes);
  mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, relsPath);

  const relationships = parseRelationships(sourceArchive, relsPath, normalizedPartPath);
  for (const relationship of relationships.values()) {
    if (relationship.targetMode === "External" || relationship.target.startsWith("#")) {
      continue;
    }
    copyArchivePartWithDependencies(
      targetArchive,
      sourceArchive,
      targetContentTypesDocument,
      originalContentTypesDocument,
      relationship.target,
      copiedPartPaths
    );
  }
}

function buildMergedDrawingPart(
  targetArchive: ArchiveEntries,
  sourceArchive: ArchiveEntries,
  targetContentTypesDocument: XMLDocument,
  originalContentTypesDocument: XMLDocument | null,
  attachments: XlsxImageAttachment[],
  copiedPartPaths: Set<string>
) {
  const mergedDrawingPath = nextDrawingPartPath(targetArchive);
  const mergedRelsPath = relsPathForDocument(mergedDrawingPath);
  const mergedRelDocument = createRelationshipsDocument();
  if (!mergedRelDocument) {
    return null;
  }

  let mergedDrawingDocument: XMLDocument | null = null;

  for (const attachment of attachments) {
    const drawingXml = readArchiveText(sourceArchive, attachment.drawingPath);
    const sourceDrawingDocument = drawingXml ? parseXml(drawingXml) : null;
    if (!sourceDrawingDocument?.documentElement) {
      continue;
    }

    let anchorNodes: Element[];
    if (!mergedDrawingDocument) {
      mergedDrawingDocument = sourceDrawingDocument;
      anchorNodes = drawingChildElements(mergedDrawingDocument);
    } else {
      anchorNodes = [];
      for (const child of drawingChildElements(sourceDrawingDocument)) {
        if (child.localName === "extLst") {
          continue;
        }
        const imported = mergedDrawingDocument.importNode(child, true) as Element;
        mergedDrawingDocument.documentElement?.appendChild(imported);
        anchorNodes.push(imported);
      }
    }

    const relationshipRecords = attachment.drawingRelsPath
      ? parseRelationships(sourceArchive, attachment.drawingRelsPath, attachment.drawingPath)
      : null;
    const sourceRelsXml = attachment.drawingRelsPath
      ? readArchiveText(sourceArchive, attachment.drawingRelsPath)
      : null;
    const sourceRelsDocument = sourceRelsXml ? parseXml(sourceRelsXml) : null;
    const idMap = new Map<string, string>();

    if (sourceRelsDocument) {
      for (const sourceRelationshipNode of getLocalElements(sourceRelsDocument, "Relationship")) {
        const oldId = sourceRelationshipNode.getAttribute("Id");
        const type = sourceRelationshipNode.getAttribute("Type");
        const rawTarget = sourceRelationshipNode.getAttribute("Target");
        if (!oldId || !type || !rawTarget) {
          continue;
        }
        const record = relationshipRecords?.get(oldId);
        if (!record) {
          continue;
        }

        const relationshipId = nextDrawingRelationshipId(mergedRelDocument);
        idMap.set(oldId, relationshipId);

        let target = rawTarget;
        if (record.targetMode !== "External" && !rawTarget.startsWith("#")) {
          target = relativeArchivePath(mergedDrawingPath, record.target);
          copyArchivePartWithDependencies(
            targetArchive,
            sourceArchive,
            targetContentTypesDocument,
            originalContentTypesDocument,
            record.target,
            copiedPartPaths
          );
        }

        const mergedRelationshipNode = mergedRelDocument.createElementNS(PKG_REL_NS, "Relationship");
        mergedRelationshipNode.setAttribute("Id", relationshipId);
        mergedRelationshipNode.setAttribute("Type", type);
        mergedRelationshipNode.setAttribute("Target", target);
        if (record.targetMode) {
          mergedRelationshipNode.setAttribute("TargetMode", record.targetMode);
        }
        mergedRelDocument.documentElement?.appendChild(mergedRelationshipNode);
      }
    }

    anchorNodes.forEach((node) => remapRelationshipReferences(node, idMap));

    attachment.mediaPaths.forEach((mediaPath) => {
      copyArchivePartWithDependencies(
        targetArchive,
        sourceArchive,
        targetContentTypesDocument,
        originalContentTypesDocument,
        mediaPath,
        copiedPartPaths
      );
    });
  }

  if (!mergedDrawingDocument) {
    return null;
  }

  targetArchive[mergedDrawingPath] = strToU8(serializeXml(mergedDrawingDocument));
  targetArchive[mergedRelsPath] = strToU8(serializeXml(mergedRelDocument));
  ensureContentTypeOverride(targetContentTypesDocument, mergedDrawingPath, DRAWING_CONTENT_TYPE);
  mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, mergedRelsPath);
  return mergedDrawingPath;
}

export function mergeWorkbookImageAssets(
  savedBytes: Uint8Array,
  sourceAssets: WorkbookImageAssets | null,
  sheetOrigins: Array<WorkbookImageSheetOrigin | null>
) {
  if (!sourceAssets || sheetOrigins.every((origin) => !origin?.attachments.length)) {
    return cloneBytes(savedBytes);
  }

  try {
    const archive = unzipSync(savedBytes);
    const workbookSheets = parseWorkbookSheets(archive);
    const originalContentTypesDocument = parseXml(readArchiveText(sourceAssets.archive, "[Content_Types].xml") ?? "");
    const targetContentTypesDocument = ensureContentTypesDocument(archive);
    if (!targetContentTypesDocument) {
      return cloneBytes(savedBytes);
    }

    const copiedPartPaths = new Set<string>();

    sheetOrigins.forEach((origin, workbookSheetIndex) => {
      if (!origin?.attachments.length) {
        return;
      }

      const currentSheet = workbookSheets[workbookSheetIndex];
      if (!currentSheet) {
        return;
      }

      const sheetXml = readArchiveText(archive, currentSheet.path);
      if (!sheetXml) {
        return;
      }

      const sheetDocument = parseXml(sheetXml);
      const relsPath = relsPathForDocument(currentSheet.path);
      const relDocument = ensureRelationshipsDocument(archive, relsPath);
      if (!sheetDocument || !relDocument) {
        return;
      }

      removeDrawingReferences(sheetDocument, relDocument);

      const mergedDrawingPath = buildMergedDrawingPart(
        archive,
        sourceAssets.archive,
        targetContentTypesDocument,
        originalContentTypesDocument,
        origin.attachments,
        copiedPartPaths
      );
      if (!mergedDrawingPath) {
        return;
      }

      const relationshipId = nextRelationshipId(relDocument);
      const relationshipNode = relDocument.createElementNS(PKG_REL_NS, "Relationship");
      relationshipNode.setAttribute("Id", relationshipId);
      relationshipNode.setAttribute("Type", DRAWING_REL_TYPE);
      relationshipNode.setAttribute("Target", relativeArchivePath(currentSheet.path, mergedDrawingPath));
      relDocument.documentElement?.appendChild(relationshipNode);
      appendSheetDrawingReference(sheetDocument, relationshipId);

      archive[currentSheet.path] = strToU8(serializeXml(sheetDocument));
      archive[relsPath] = strToU8(serializeXml(relDocument));
      mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, relsPath);
    });

    const hasDrawingOverride = getLocalElements(targetContentTypesDocument, "Override").some(
      (node) => node.getAttribute("ContentType") === DRAWING_CONTENT_TYPE
    );
    if (!hasDrawingOverride) {
      for (const path of Object.keys(archive)) {
        if (path.startsWith("xl/drawings/") && path.endsWith(".xml")) {
          mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, path);
        }
      }
    }

    archive["[Content_Types].xml"] = strToU8(serializeXml(targetContentTypesDocument));
    return zipSync(archive, { level: 6 });
  } catch {
    return cloneBytes(savedBytes);
  }
}
