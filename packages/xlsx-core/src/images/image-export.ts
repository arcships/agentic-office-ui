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
  type WorkbookImageOrigin,
  type WorkbookImageSheetOrigin,
  cloneBytes,
  createImageAnchorNodes,
  getChildElements,
  getFirstChild,
  getLocalElements,
  normalizeArchivePath,
  parseXml,
  readArchiveText,
  relativeArchivePath,
  relsPathForDocument,
  serializeXml,
  setChildText,
  updateMarkerElement,
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
    const sheetRelationships = parseRelationshipsWrapper(archive, relsPathForDocument(sheet.path), sheet.path);
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

import { parseRelationships } from "./image-parser";
const parseRelationshipsWrapper = parseRelationships;

type XlsxImageAttachment = {
  drawingPath: string;
  drawingRelsPath: string | null;
  mediaPaths: string[];
};

export function parseWorkbookImageAssets(bytes: Uint8Array): WorkbookImageAssets {
  const archive = unzipSync(bytes);
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
  const formControlsByWorkbookSheetIndex: XlsxFormControl[][] = [];
  const imagesByWorkbookSheetIndex: XlsxImage[][] = [];
  const shapesByWorkbookSheetIndex: XlsxShape[][] = [];
  const sheetOrigins: Array<WorkbookImageSheetOrigin | null> = [];
  const imageOriginsById = new Map<string, WorkbookImageOrigin>();

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
        imageOriginsById
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
      zIndexBase
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
    formControlsByWorkbookSheetIndex,
    imageOriginsById,
    imagesByWorkbookSheetIndex,
    namedCellStyleByName,
    objectUrls,
    shapesByWorkbookSheetIndex,
    sheetOrigins,
    sheetStatesByWorkbookSheetIndex,
    styleById,
    tableMetadataByWorkbookSheetIndex,
    tableStyleByName,
    themePalette
  };
}

// ── Update image anchor ────────────────────────────────────────────────────

function updateAnchorNode(anchorNode: Element, anchor: XlsxImage["anchor"]) {
  if (anchor.kind === "two-cell") {
    updateMarkerElement(getFirstChild(anchorNode, "from"), anchor.from);
    updateMarkerElement(getFirstChild(anchorNode, "to"), anchor.to);
    return;
  }

  if (anchor.kind === "one-cell") {
    updateMarkerElement(getFirstChild(anchorNode, "from"), anchor.from);
    const extNode = getFirstChild(anchorNode, "ext");
    if (extNode) {
      extNode.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
      extNode.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
    }
    return;
  }

  const positionNode = getFirstChild(anchorNode, "pos");
  if (positionNode) {
    positionNode.setAttribute("x", String(Math.max(0, Math.round(anchor.positionEmu.x))));
    positionNode.setAttribute("y", String(Math.max(0, Math.round(anchor.positionEmu.y))));
  }
  const extNode = getFirstChild(anchorNode, "ext");
  if (extNode) {
    extNode.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
    extNode.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
  }
}

export function updateWorkbookImageAnchor(
  assets: WorkbookImageAssets,
  imageId: string,
  anchor: XlsxImage["anchor"]
) {
  const origin = assets.imageOriginsById.get(imageId);
  if (!origin) {
    return false;
  }

  const attachments = assets.sheetOrigins[origin.workbookSheetIndex]?.attachments ?? [];
  for (const attachment of attachments) {
    const drawingXml = readArchiveText(assets.archive, attachment.drawingPath);
    if (!drawingXml) {
      continue;
    }

    const drawingDocument = parseXml(drawingXml);
    if (!drawingDocument) {
      continue;
    }

    const anchorNodes = createImageAnchorNodes(drawingDocument);
    const anchorNode = anchorNodes[origin.anchorIndex];
    if (!anchorNode || !getFirstChild(anchorNode, "pic")) {
      continue;
    }

    updateAnchorNode(anchorNode, anchor);
    assets.archive[attachment.drawingPath] = strToU8(serializeXml(drawingDocument));
    const imageList = assets.imagesByWorkbookSheetIndex[origin.workbookSheetIndex] ?? [];
    const imageIndex = imageList.findIndex((image) => image.id === imageId);
    if (imageIndex >= 0) {
      imageList[imageIndex] = {
        ...imageList[imageIndex],
        anchor
      };
    }
    return true;
  }

  return false;
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

  const extLst = getFirstChild(worksheet, "extLst");
  if (extLst) {
    worksheet.insertBefore(drawingNode, extLst);
    return;
  }

  worksheet.appendChild(drawingNode);
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

      origin.attachments.forEach((attachment) => {
        const drawingBytes = sourceAssets.archive[attachment.drawingPath];
        if (!drawingBytes) {
          return;
        }

        archive[attachment.drawingPath] = cloneBytes(drawingBytes);
        mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, attachment.drawingPath);

        if (attachment.drawingRelsPath) {
          const drawingRelsBytes = sourceAssets.archive[attachment.drawingRelsPath];
          if (drawingRelsBytes) {
            archive[attachment.drawingRelsPath] = cloneBytes(drawingRelsBytes);
            mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, attachment.drawingRelsPath);
          }
        }

        attachment.mediaPaths.forEach((mediaPath) => {
          const mediaBytes = sourceAssets.archive[mediaPath];
          if (!mediaBytes) {
            return;
          }

          archive[mediaPath] = cloneBytes(mediaBytes);
          mergeContentTypeForPath(targetContentTypesDocument, originalContentTypesDocument, mediaPath);
        });

        const relationshipId = nextRelationshipId(relDocument);
        const relationshipNode = relDocument.createElementNS(PKG_REL_NS, "Relationship");
        relationshipNode.setAttribute("Id", relationshipId);
        relationshipNode.setAttribute("Type", DRAWING_REL_TYPE);
        relationshipNode.setAttribute("Target", relativeArchivePath(currentSheet.path, attachment.drawingPath));
        relDocument.documentElement?.appendChild(relationshipNode);
        appendSheetDrawingReference(sheetDocument, relationshipId);
      });

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
